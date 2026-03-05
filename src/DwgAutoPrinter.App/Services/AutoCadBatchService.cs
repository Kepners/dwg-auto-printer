using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DwgAutoPrinter.App.Models;

namespace DwgAutoPrinter.App.Services;

public sealed class AutoCadBatchService
{
    private readonly LispDeploymentService _lispDeploymentService = new();
    private dynamic? _acad;
    private dynamic? _activeDoc;

    public Task RunAsync(RunOptions options, Action<LogEntry> log, CancellationToken cancellationToken)
    {
        var tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);

        var worker = new Thread(() =>
        {
            try
            {
                ExecuteOnStaThread(options, log, cancellationToken);
                tcs.TrySetResult(null);
            }
            catch (OperationCanceledException)
            {
                tcs.TrySetCanceled(cancellationToken);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        })
        {
            IsBackground = true,
            Name = "DWGAutoPrinter.Automation.STA"
        };

        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();
        return tcs.Task;
    }

    public void RequestStop(Action<LogEntry> log)
    {
        try
        {
            if (_activeDoc is not null)
            {
                // ESC, ESC to cancel active AutoCAD command safely.
                _activeDoc.SendCommand("\u001B\u001B\n");
            }

            log(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Warning,
                Message = "Stop requested. Attempted to cancel active AutoCAD command."
            });
        }
        catch (Exception ex)
        {
            log(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Warning,
                Message = $"Stop requested but cancel signal failed: {ex.Message}"
            });
        }
    }

    private void ExecuteOnStaThread(RunOptions options, Action<LogEntry> log, CancellationToken cancellationToken)
    {
        ValidateOptions(options);

        var copiedTargets = _lispDeploymentService.CopyToAutoCadLocations(options.LispSourcePath, log);
        if (copiedTargets.Count > 0)
        {
            log(Info($"LISP deployment completed. Updated {copiedTargets.Count} target location(s)."));
        }

        var dwgFiles = Directory.GetFiles(options.FolderPath, "*.dwg", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (dwgFiles.Count == 0)
        {
            throw new InvalidOperationException($"No DWG files found in {options.FolderPath}");
        }

        log(Info($"Found {dwgFiles.Count} DWG file(s)."));

        _acad = ConnectOrLaunch(log);
        _acad.Visible = options.AutoCadVisible;
        EnsureActiveDocument();

        WaitForIdle(TimeSpan.FromMinutes(2), "AutoCAD startup", cancellationToken);
        LoadLisp(options.LispSourcePath, cancellationToken, log);

        var total = dwgFiles.Count;
        var batchSize = Math.Max(1, options.BatchSize);
        for (var i = 0; i < total; i += batchSize)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var group = dwgFiles.Skip(i).Take(batchSize).ToList();
            log(Info($"Batch {(i / batchSize) + 1}: {group.Count} drawing(s)."));
            ProcessGroup(group, options, log, cancellationToken);
        }

        log(new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = LogLevel.Success,
            Message = "All drawings processed."
        });
    }

    private static void ValidateOptions(RunOptions options)
    {
        if (!Directory.Exists(options.FolderPath))
        {
            throw new DirectoryNotFoundException($"Folder not found: {options.FolderPath}");
        }

        if (!File.Exists(options.LispSourcePath))
        {
            throw new FileNotFoundException("LISP source file not found.", options.LispSourcePath);
        }

        if (!string.Equals(options.RevisionMode, "NEXT", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(options.RevisionMode, "EXACT", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Revision mode must be NEXT or EXACT.");
        }

        if (string.Equals(options.RevisionMode, "EXACT", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(options.ExactRevision))
        {
            throw new InvalidOperationException("Exact revision is required when mode is EXACT.");
        }
    }

    private dynamic ConnectOrLaunch(Action<LogEntry> log)
    {
        if (ComInterop.TryGetActiveObject("AutoCAD.Application", out var active))
        {
            log(Info("Attached to running AutoCAD instance."));
            return active!;
        }

        var progIdType = Type.GetTypeFromProgID("AutoCAD.Application", throwOnError: false);
        if (progIdType is null)
        {
            throw new InvalidOperationException("AutoCAD COM ProgID was not found.");
        }

        var created = Activator.CreateInstance(progIdType);
        if (created is null)
        {
            throw new InvalidOperationException("Unable to launch AutoCAD COM instance.");
        }

        log(Info("Launched new AutoCAD instance."));
        return created;
    }

    private void EnsureActiveDocument()
    {
        var docs = _acad!.Documents;
        if (docs.Count == 0)
        {
            docs.Add();
        }

        _activeDoc = _acad.ActiveDocument;
        if (_activeDoc is null)
        {
            throw new InvalidOperationException("No active AutoCAD document available.");
        }
    }

    private void LoadLisp(string lispPath, CancellationToken cancellationToken, Action<LogEntry> log)
    {
        var escaped = EscapeLispString(ToLispPath(lispPath));
        SendToActiveDoc($"(load \"{escaped}\")");
        WaitForIdle(TimeSpan.FromMinutes(1), "LISP load", cancellationToken);
        log(Info($"Loaded LISP from {lispPath}"));
    }

    private void ProcessGroup(
        List<string> group,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        var openedDocs = new List<dynamic>();

        foreach (var path in group)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var doc = TryOpen(path, log, cancellationToken);
            if (doc is not null)
            {
                openedDocs.Add(doc);
            }
        }

        foreach (dynamic doc in openedDocs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ProcessDocument(doc, options, log, cancellationToken);
        }

        if (options.CloseAfterProcess)
        {
            foreach (dynamic doc in openedDocs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                TryClose(doc, log, cancellationToken);
            }
        }
    }

    private dynamic? TryOpen(string path, Action<LogEntry> log, CancellationToken cancellationToken)
    {
        try
        {
            var docs = _acad!.Documents;
            dynamic doc = docs.Open(path);
            _activeDoc = doc;
            WaitForIdle(TimeSpan.FromMinutes(2), $"Open {Path.GetFileName(path)}", cancellationToken);
            log(Info($"Opened {path}"));
            return doc;
        }
        catch (Exception ex)
        {
            log(Error($"Failed to open {path}: {ex.Message}"));
            return null;
        }
    }

    private void ProcessDocument(
        dynamic doc,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        string drawingName = doc.Name;

        try
        {
            doc.Activate();
            _activeDoc = doc;
            WaitForIdle(TimeSpan.FromMinutes(1), $"Activate {drawingName}", cancellationToken);

            var cmd = BuildExeRunCommand(options);
            SendToActiveDoc(cmd);
            WaitForIdle(TimeSpan.FromMinutes(options.TimeoutMinutes), $"Revision run {drawingName}", cancellationToken);

            doc.Save();
            WaitForIdle(TimeSpan.FromMinutes(1), $"Save {drawingName}", cancellationToken);
            log(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Success,
                Message = $"Processed {drawingName}"
            });
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            log(Error($"Processing failed for {drawingName}: {ex.Message}"));
        }
    }

    private void TryClose(dynamic doc, Action<LogEntry> log, CancellationToken cancellationToken)
    {
        var name = (string)doc.Name;
        try
        {
            doc.Close(true);
            WaitForIdle(TimeSpan.FromMinutes(1), $"Close {name}", cancellationToken);
            log(Info($"Closed {name}"));
        }
        catch (Exception ex)
        {
            log(Error($"Failed to close {name}: {ex.Message}"));
        }
    }

    private string BuildExeRunCommand(RunOptions options)
    {
        var folder = EscapeLispString(ToLispPath(options.FolderPath));
        var mode = EscapeLispString(options.RevisionMode.ToUpperInvariant());
        var exact = EscapeLispString(options.ExactRevision ?? string.Empty);
        var description = EscapeLispString(options.Description);
        var closeAfter = options.CloseAfterProcess ? "T" : "NIL";

        return $"(dap:exe-run \"{folder}\" {options.BatchSize} \"{mode}\" \"{exact}\" \"{description}\" {closeAfter})";
    }

    private void WaitForIdle(TimeSpan timeout, string phase, CancellationToken cancellationToken)
    {
        var started = DateTime.UtcNow;
        while (DateTime.UtcNow - started < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                bool isIdle = _acad!.GetAcadState().IsQuiescent;
                if (isIdle)
                {
                    return;
                }
            }
            catch
            {
                // transient COM busy state while AutoCAD processes commands.
            }

            Thread.Sleep(300);
        }

        throw new TimeoutException($"Timed out waiting for AutoCAD idle at phase: {phase}");
    }

    private void SendToActiveDoc(string command)
    {
        if (_activeDoc is null)
        {
            throw new InvalidOperationException("No active AutoCAD document for command dispatch.");
        }

        _activeDoc.SendCommand(command + "\n");
    }

    private static string EscapeLispString(string value)
    {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static string ToLispPath(string inputPath)
    {
        return Path.GetFullPath(inputPath).Replace('\\', '/');
    }

    private static LogEntry Info(string message) => new()
    {
        Timestamp = DateTime.Now,
        Message = message,
        Level = LogLevel.Info
    };

    private static LogEntry Error(string message) => new()
    {
        Timestamp = DateTime.Now,
        Message = message,
        Level = LogLevel.Error
    };
}
