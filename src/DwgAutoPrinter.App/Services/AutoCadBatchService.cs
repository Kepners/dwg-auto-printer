using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
                using var _ = ComMessageFilter.Register();
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
        log(Info("Preflight: validating options."));
        ValidateOptions(options);
        log(Info("Preflight: options valid."));

        var copiedTargets = _lispDeploymentService.CopyToAutoCadLocations(options.LispSourcePath, log);
        if (copiedTargets.Count > 0)
        {
            log(Info($"LISP deployment completed. Updated {copiedTargets.Count} target location(s)."));
        }
        else
        {
            log(Info("LISP deployment did not find writable targets. Will fallback to bundled path if needed."));
        }

        var lispLoadPath = SelectLispLoadPath(options.LispSourcePath, copiedTargets, log);
        var lispFingerprint = TryGetFileFingerprint(lispLoadPath);
        if (!string.IsNullOrWhiteSpace(lispFingerprint))
        {
            log(Info($"LISP fingerprint: {lispFingerprint}"));
        }

        var dwgFiles = Directory.GetFiles(options.FolderPath, "*.dwg", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (dwgFiles.Count == 0)
        {
            throw new InvalidOperationException($"No DWG files found in {options.FolderPath}");
        }

        log(Info($"Found {dwgFiles.Count} DWG file(s)."));
        if (string.Equals(options.RevisionMode, "NEXT", StringComparison.OrdinalIgnoreCase))
        {
            log(Info("Revision mode NEXT: each drawing reads current revision and increments it before PDF plotting."));
        }

        _acad = ConnectOrLaunch(log);
        _acad.Visible = options.AutoCadVisible;
        EnsureActiveDocument();

        WaitForIdle(TimeSpan.FromMinutes(2), "AutoCAD startup", cancellationToken);
        EnsureTrustedPathForLisp(lispLoadPath, cancellationToken, log);
        LoadLisp(lispLoadPath, cancellationToken, log);

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
        log(Info($"Loading LISP from {lispPath}"));
        SendToActiveDoc($"(load \"{escaped}\")");
        WaitForIdle(TimeSpan.FromMinutes(1), "LISP load", cancellationToken);
        log(Info($"Loaded LISP from {lispPath}"));
    }

    private void EnsureTrustedPathForLisp(string lispPath, CancellationToken cancellationToken, Action<LogEntry> log)
    {
        var directory = Path.GetDirectoryName(lispPath);
        if (string.IsNullOrWhiteSpace(directory))
        {
            return;
        }

        var lispDir = EscapeLispString(ToLispPath(directory));
        var expr =
            $"(if (not (vl-string-search \"{lispDir}\" (getvar \"TRUSTEDPATHS\"))) (setvar \"TRUSTEDPATHS\" (strcat (getvar \"TRUSTEDPATHS\") \";{lispDir}\")))";

        SendToActiveDoc(expr);
        WaitForIdle(TimeSpan.FromSeconds(20), "Update TRUSTEDPATHS", cancellationToken);
        log(Info($"Trusted path ensured for LISP: {directory}"));
    }

    private void ProcessGroup(
        List<string> group,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        var openedDocs = new List<OpenedDocument>();

        foreach (var path in group)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var doc = TryOpen(path, log, cancellationToken);
            if (doc is not null)
            {
                var layouts = GetPaperSpaceLayouts(doc, log);
                var paperSpaceCount = layouts.Count;
                var docName = TryGetDocumentName(doc);
                log(Info($"{docName}: detected {paperSpaceCount} paper space layout(s)."));

                if (paperSpaceCount > 3)
                {
                    if (openedDocs.Count > 0)
                    {
                        ProcessLoadedDocs(openedDocs, options, log, cancellationToken);
                        openedDocs.Clear();
                    }

                    log(Info($"Single-DWG mode: {docName} has {paperSpaceCount} paper space layouts (>3)."));
                    ProcessDocument(doc, layouts, options, log, cancellationToken);
                    if (options.CloseAfterProcess)
                    {
                        TryClose(doc, log, cancellationToken);
                    }
                }
                else
                {
                    openedDocs.Add(new OpenedDocument(doc, layouts));
                    if (openedDocs.Count >= options.BatchSize)
                    {
                        ProcessLoadedDocs(openedDocs, options, log, cancellationToken);
                        openedDocs.Clear();
                    }
                }
            }
        }

        ProcessLoadedDocs(openedDocs, options, log, cancellationToken);
    }

    private void ProcessLoadedDocs(
        List<OpenedDocument> openedDocs,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        foreach (var opened in openedDocs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ProcessDocument(opened.Document, opened.PaperSpaceLayouts, options, log, cancellationToken);
        }

        if (options.CloseAfterProcess)
        {
            foreach (var opened in openedDocs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                TryClose(opened.Document, log, cancellationToken);
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

    private IReadOnlyList<string> GetPaperSpaceLayouts(dynamic doc, Action<LogEntry> log)
    {
        var names = new List<string>();

        try
        {
            dynamic layouts = doc.Layouts;
            int total = layouts.Count;
            if (!TryCollectLayoutNamesByIndex(layouts, total, 0, names))
            {
                names.Clear();
                TryCollectLayoutNamesByIndex(layouts, total, 1, names);
            }
        }
        catch (Exception ex)
        {
            var docName = TryGetDocumentName(doc);
            log(Error($"Layout index enumeration failed for {docName}: {ex.Message}"));
        }

        if (names.Count > 0)
        {
            return names
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        try
        {
            foreach (dynamic layout in doc.Layouts)
            {
                string name = Convert.ToString(layout.Name) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(name) &&
                    !string.Equals(name, "Model", StringComparison.OrdinalIgnoreCase))
                {
                    names.Add(name);
                }
            }
        }
        catch (Exception ex)
        {
            var docName = TryGetDocumentName(doc);
            log(Error($"Layout fallback enumeration failed for {docName}: {ex.Message}"));
        }

        return names
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool TryCollectLayoutNamesByIndex(dynamic layouts, int total, int startIndex, List<string> names)
    {
        try
        {
            for (var i = 0; i < total; i++)
            {
                dynamic layout = layouts.Item(i + startIndex);
                string name = Convert.ToString(layout.Name) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(name) &&
                    !string.Equals(name, "Model", StringComparison.OrdinalIgnoreCase))
                {
                    names.Add(name);
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private void ProcessDocument(
        dynamic doc,
        IReadOnlyList<string> paperSpaceLayouts,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        string drawingName = TryGetDocumentName(doc);

        try
        {
            doc.Activate();
            _activeDoc = doc;
            WaitForIdle(TimeSpan.FromMinutes(1), $"Activate {drawingName}", cancellationToken);

            var cmd = BuildExeRunCommand(options);
            log(Info($"AutoCAD command -> {cmd}"));
            SendToActiveDoc(cmd);
            WaitForIdle(TimeSpan.FromMinutes(options.TimeoutMinutes), $"Revision run {drawingName}", cancellationToken);

            doc.Save();
            WaitForIdle(TimeSpan.FromMinutes(1), $"Save {drawingName}", cancellationToken);

            PlotPaperSpaceLayoutsToPdf(doc, paperSpaceLayouts, options, log, cancellationToken);

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

    private void PlotPaperSpaceLayoutsToPdf(
        dynamic doc,
        IReadOnlyList<string> paperSpaceLayouts,
        RunOptions options,
        Action<LogEntry> log,
        CancellationToken cancellationToken)
    {
        string drawingName = TryGetDocumentName(doc);
        string drawingStem = Path.GetFileNameWithoutExtension(drawingName);
        string fullName = TryGetDocumentFullName(doc);
        string outputFolder = string.IsNullOrWhiteSpace(fullName)
            ? options.FolderPath
            : (Path.GetDirectoryName(fullName) ?? options.FolderPath);
        var layouts = paperSpaceLayouts.Count > 0
            ? paperSpaceLayouts
            : GetPaperSpaceLayouts(doc, log);
        var originalLayoutName = TryGetActiveLayoutName(doc);
        var originalBackgroundPlot = TryGetVariableInt(doc, "BACKGROUNDPLOT");

        try
        {
            if (layouts.Count == 0)
            {
                log(Error($"No paper space layouts found for plotting in {drawingName}."));
                return;
            }

            try
            {
                doc.Activate();
                _activeDoc = doc;
                WaitForIdle(TimeSpan.FromMinutes(1), $"Activate for plot {drawingName}", cancellationToken);
            }
            catch (Exception ex)
            {
                log(Error($"Failed to activate {drawingName} before plotting: {ex.Message}"));
            }

            SetVariableIfPossible(doc, "BACKGROUNDPLOT", 0, log);

            foreach (var layoutName in layouts)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ActivateLayout(doc, layoutName, cancellationToken, log);

                var safeLayout = SanitizeFileName(layoutName);
                var outputPath = Path.Combine(outputFolder, $"{drawingStem}-{safeLayout}.pdf");
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                log(Info($"Plotting {drawingName} [{layoutName}] -> {outputPath}"));

                dynamic plot = doc.Plot;
                try
                {
                    plot.QuietErrorMode = true;
                }
                catch
                {
                    // Optional COM property.
                }

                try
                {
                    doc.Regen(1);
                }
                catch
                {
                    // Best-effort, not fatal.
                }

                var layoutsToPlot = new[] { layoutName };
                try
                {
                    plot.SetLayoutsToPlot(layoutsToPlot);
                }
                catch
                {
                    // Active-layout fallback still works if SetLayoutsToPlot is unavailable.
                }

                bool plotAccepted = true;
                var plotResult = plot.PlotToFile(outputPath);
                if (plotResult is bool boolResult)
                {
                    plotAccepted = boolResult;
                }
                else
                {
                    string resultText = Convert.ToString(plotResult) ?? string.Empty;
                    if (bool.TryParse(resultText, out bool parsed))
                    {
                        plotAccepted = parsed;
                    }
                }

                WaitForIdle(TimeSpan.FromMinutes(2), $"Plot {drawingName} [{layoutName}]", cancellationToken);

                if (plotAccepted && File.Exists(outputPath))
                {
                    log(new LogEntry
                    {
                        Timestamp = DateTime.Now,
                        Level = LogLevel.Success,
                        Message = $"PDF created: {Path.GetFileName(outputPath)}"
                    });
                }
                else
                {
                    log(Error($"Plot failed for {drawingName} [{layoutName}] (accepted={plotAccepted}, file={File.Exists(outputPath)})."));
                }
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            log(Error($"PDF plotting failed for {drawingName}: {ex.Message}"));
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(originalLayoutName))
            {
                try
                {
                    ActivateLayout(doc, originalLayoutName, cancellationToken, log);
                }
                catch
                {
                    // best-effort cleanup only
                }
            }

            if (originalBackgroundPlot.HasValue)
            {
                SetVariableIfPossible(doc, "BACKGROUNDPLOT", originalBackgroundPlot.Value, log);
            }
        }
    }

    private void ActivateLayout(dynamic doc, string layoutName, CancellationToken cancellationToken, Action<LogEntry> log)
    {
        try
        {
            dynamic layouts = doc.Layouts;
            dynamic layout = layouts.Item(layoutName);
            doc.ActiveLayout = layout;
        }
        catch (Exception ex)
        {
            try
            {
                doc.SetVariable("CTAB", layoutName);
            }
            catch
            {
                log(Error($"Failed to activate layout {layoutName}: {ex.Message}"));
                throw;
            }
        }

        WaitForIdle(TimeSpan.FromSeconds(30), $"Activate layout {layoutName}", cancellationToken);
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = value.ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            if (invalid.Contains(chars[i]))
            {
                chars[i] = '-';
            }
        }

        var sanitized = new string(chars).Trim();
        return string.IsNullOrWhiteSpace(sanitized) ? "Layout" : sanitized;
    }

    private static string? TryGetActiveLayoutName(dynamic doc)
    {
        try
        {
            return Convert.ToString(doc.ActiveLayout.Name);
        }
        catch
        {
            return null;
        }
    }

    private static int? TryGetVariableInt(dynamic doc, string variable)
    {
        try
        {
            var raw = doc.GetVariable(variable);
            if (raw is int i)
            {
                return i;
            }

            var rawText = Convert.ToString(raw);
            if (int.TryParse(rawText, out int parsed))
            {
                return parsed;
            }
        }
        catch
        {
            // best-effort only
        }

        return null;
    }

    private static void SetVariableIfPossible(dynamic doc, string variable, int value, Action<LogEntry> log)
    {
        try
        {
            doc.SetVariable(variable, value);
        }
        catch (Exception ex)
        {
            log(Error($"Unable to set {variable}={value}: {ex.Message}"));
        }
    }

    private void TryClose(dynamic doc, Action<LogEntry> log, CancellationToken cancellationToken)
    {
        var name = TryGetDocumentName(doc);
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

    private static string TryGetDocumentName(dynamic doc)
    {
        try
        {
            var name = Convert.ToString(doc.Name);
            return string.IsNullOrWhiteSpace(name) ? "<unknown.dwg>" : name;
        }
        catch
        {
            return "<unknown.dwg>";
        }
    }

    private static string TryGetDocumentFullName(dynamic doc)
    {
        try
        {
            return Convert.ToString(doc.FullName) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string EscapeLispString(string value)
    {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static string ToLispPath(string inputPath)
    {
        return Path.GetFullPath(inputPath).Replace('\\', '/');
    }

    private static string SelectLispLoadPath(string bundledPath, IReadOnlyList<string> copiedTargets, Action<LogEntry> log)
    {
        _ = copiedTargets;
        log(Info($"Preflight: EXE-controlled mode; loading bundled LISP: {bundledPath}"));
        return bundledPath;
    }

    private static string TryGetFileFingerprint(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return string.Empty;
            }

            using var stream = File.OpenRead(path);
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(stream);
            var shortHash = Convert.ToHexString(hash)[..12];
            return $"{shortHash} ({path})";
        }
        catch
        {
            return string.Empty;
        }
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

    private sealed record OpenedDocument(dynamic Document, IReadOnlyList<string> PaperSpaceLayouts);
}
