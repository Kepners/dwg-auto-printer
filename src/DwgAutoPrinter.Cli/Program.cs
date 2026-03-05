using System.Runtime.InteropServices;

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("DWGAutoPrinter requires Windows.");
    return 1;
}

AppOptions options;
try
{
    options = AppOptions.Parse(args);
}
catch (ArgumentException ex)
{
    Console.Error.WriteLine(ex.Message);
    AppOptions.PrintUsage();
    return 2;
}

if (options.ShowHelp)
{
    AppOptions.PrintUsage();
    return 0;
}

try
{
    var runner = new AutoCadBatchRunner(options);
    runner.Run();
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"ERROR: {ex.Message}");
    return 1;
}

internal sealed record AppOptions(
    string Folder,
    int BatchSize,
    string RevisionMode,
    string ExactRevision,
    string Description,
    bool CloseAfter,
    string LispPath,
    bool Visible,
    int TimeoutMinutes,
    bool ShowHelp)
{
    public static AppOptions Parse(string[] args)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var token = args[i];
            if (!token.StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Unexpected argument '{token}'.");
            }

            if (token.Equals("--help", StringComparison.OrdinalIgnoreCase) || token.Equals("-h", StringComparison.OrdinalIgnoreCase))
            {
                values["help"] = "true";
                continue;
            }

            if (i + 1 >= args.Length)
            {
                throw new ArgumentException($"Missing value for '{token}'.");
            }

            values[token[2..]] = args[++i];
        }

        var showHelp = values.ContainsKey("help");
        var folder = Path.GetFullPath(values.GetValueOrDefault("folder", Environment.CurrentDirectory));
        var batchSize = ParsePositiveInt(values.GetValueOrDefault("batch-size", "5"), "--batch-size");
        var revisionMode = values.GetValueOrDefault("revision-mode", "NEXT").Trim().ToUpperInvariant();
        if (revisionMode is not ("NEXT" or "EXACT"))
        {
            throw new ArgumentException("Invalid --revision-mode. Use NEXT or EXACT.");
        }

        var exactRevision = values.GetValueOrDefault("exact-revision", string.Empty).Trim();
        if (revisionMode == "EXACT" && string.IsNullOrWhiteSpace(exactRevision))
        {
            throw new ArgumentException("--exact-revision is required when --revision-mode EXACT is used.");
        }

        var description = values.GetValueOrDefault("description", "Construction Issue");
        var closeAfter = ParseBool(values.GetValueOrDefault("close-after", "true"), "--close-after");
        var visible = ParseBool(values.GetValueOrDefault("visible", "true"), "--visible");
        var timeoutMinutes = ParsePositiveInt(values.GetValueOrDefault("timeout-minutes", "30"), "--timeout-minutes");

        var defaultLsp = Path.Combine(AppContext.BaseDirectory, "smart-revision-update.lsp");
        var lispPath = Path.GetFullPath(values.GetValueOrDefault("lsp-path", defaultLsp));

        return new AppOptions(
            folder,
            batchSize,
            revisionMode,
            exactRevision,
            description,
            closeAfter,
            lispPath,
            visible,
            timeoutMinutes,
            showHelp);
    }

    private static int ParsePositiveInt(string raw, string key)
    {
        if (!int.TryParse(raw, out var value) || value <= 0)
        {
            throw new ArgumentException($"Invalid {key} value '{raw}'. Must be a positive integer.");
        }

        return value;
    }

    private static bool ParseBool(string raw, string key)
    {
        if (bool.TryParse(raw, out var value))
        {
            return value;
        }

        return raw.Trim().ToLowerInvariant() switch
        {
            "1" or "yes" or "y" => true,
            "0" or "no" or "n" => false,
            _ => throw new ArgumentException($"Invalid {key} value '{raw}'. Use true or false.")
        };
    }

    public static void PrintUsage()
    {
        Console.WriteLine("DWGAutoPrinter.Cli");
        Console.WriteLine("Batch-controls AutoCAD via COM and runs smart-revision-update.lsp on each DWG.");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  DwgAutoPrinter.Cli [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --folder <path>            Root folder containing DWG files (default: current folder)");
        Console.WriteLine("  --batch-size <n>           Open/process drawings in groups (default: 5)");
        Console.WriteLine("  --revision-mode NEXT|EXACT Revision strategy (default: NEXT)");
        Console.WriteLine("  --exact-revision <value>   Required if --revision-mode EXACT");
        Console.WriteLine("  --description <text>       Revision description (default: Construction Issue)");
        Console.WriteLine("  --close-after true|false   Close drawings after processing (default: true)");
        Console.WriteLine("  --lsp-path <path>          Path to smart-revision-update.lsp");
        Console.WriteLine("  --visible true|false       AutoCAD visibility (default: true)");
        Console.WriteLine("  --timeout-minutes <n>      Idle wait timeout per drawing (default: 30)");
        Console.WriteLine("  --help                     Show this help");
    }
}

internal sealed class AutoCadBatchRunner
{
    private readonly AppOptions _options;
    private dynamic? _acad;
    private dynamic? _activeDoc;

    public AutoCadBatchRunner(AppOptions options)
    {
        _options = options;
    }

    public void Run()
    {
        ValidateInputs();
        var files = GetDwgFiles();
        Console.WriteLine($"Found {files.Count} DWG file(s) in {_options.Folder}");

        _acad = ConnectOrLaunch();
        _acad.Visible = _options.Visible;
        EnsureActiveDocument();

        WaitForIdle(TimeSpan.FromMinutes(2), "AutoCAD startup");
        LoadLispIntoActiveDoc();

        var total = files.Count;
        for (var i = 0; i < total; i += _options.BatchSize)
        {
            var group = files.Skip(i).Take(_options.BatchSize).ToList();
            Console.WriteLine();
            Console.WriteLine($"Batch {(i / _options.BatchSize) + 1}: {group.Count} drawing(s)");
            ProcessGroup(group);
        }

        Console.WriteLine("All drawings processed.");
    }

    private void ValidateInputs()
    {
        if (!Directory.Exists(_options.Folder))
        {
            throw new DirectoryNotFoundException($"Folder not found: {_options.Folder}");
        }

        if (!File.Exists(_options.LispPath))
        {
            throw new FileNotFoundException("LISP file not found.", _options.LispPath);
        }
    }

    private List<string> GetDwgFiles()
    {
        var files = Directory.GetFiles(_options.Folder, "*.dwg", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (files.Count == 0)
        {
            throw new InvalidOperationException($"No DWG files found in folder: {_options.Folder}");
        }

        return files;
    }

    private dynamic ConnectOrLaunch()
    {
        if (ComInterop.TryGetActiveObject("AutoCAD.Application", out var active))
        {
            Console.WriteLine("Attached to running AutoCAD instance.");
            return active;
        }

        var progId = Type.GetTypeFromProgID("AutoCAD.Application", throwOnError: false);
        if (progId is null)
        {
            throw new InvalidOperationException("AutoCAD COM ProgID not found. Confirm AutoCAD is installed.");
        }

        var created = Activator.CreateInstance(progId);
        if (created is null)
        {
            throw new InvalidOperationException("Failed to launch AutoCAD instance.");
        }

        Console.WriteLine("Launched new AutoCAD instance.");
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
            throw new InvalidOperationException("Unable to obtain active AutoCAD document.");
        }
    }

    private void LoadLispIntoActiveDoc()
    {
        var lispPath = EscapeLispString(ToLispPath(_options.LispPath));
        SendToActiveDoc($"(load \"{lispPath}\")");
        WaitForIdle(TimeSpan.FromMinutes(1), "LISP load");
    }

    private void ProcessGroup(List<string> group)
    {
        var openedDocs = new List<dynamic>();
        foreach (var path in group)
        {
            var doc = TryOpenDocument(path);
            if (doc is not null)
            {
                openedDocs.Add(doc);
            }
        }

        foreach (dynamic doc in openedDocs)
        {
            ProcessDocument(doc);
        }

        if (_options.CloseAfter)
        {
            foreach (dynamic doc in openedDocs)
            {
                TryCloseDocument(doc);
            }
        }
    }

    private dynamic? TryOpenDocument(string path)
    {
        try
        {
            var docs = _acad!.Documents;
            dynamic doc = docs.Open(path);
            _activeDoc = doc;
            WaitForIdle(TimeSpan.FromMinutes(2), $"open {Path.GetFileName(path)}");
            Console.WriteLine($"Opened: {path}");
            return doc;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Open failed: {path} ({ex.Message})");
            return null;
        }
    }

    private void ProcessDocument(dynamic doc)
    {
        string name = doc.Name;
        try
        {
            doc.Activate();
            _activeDoc = doc;
            WaitForIdle(TimeSpan.FromMinutes(1), $"activate {name}");

            SendToActiveDoc(BuildExeRunExpression());
            WaitForIdle(TimeSpan.FromMinutes(_options.TimeoutMinutes), $"revise {name}");

            doc.Save();
            WaitForIdle(TimeSpan.FromMinutes(1), $"save {name}");
            Console.WriteLine($"Processed: {name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Process failed: {name} ({ex.Message})");
        }
    }

    private void TryCloseDocument(dynamic doc)
    {
        string name = doc.Name;
        try
        {
            doc.Close(true);
            WaitForIdle(TimeSpan.FromMinutes(1), $"close {name}");
            Console.WriteLine($"Closed: {name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Close failed: {name} ({ex.Message})");
        }
    }

    private string BuildExeRunExpression()
    {
        var folder = EscapeLispString(ToLispPath(_options.Folder));
        var mode = EscapeLispString(_options.RevisionMode);
        var exact = EscapeLispString(_options.ExactRevision);
        var description = EscapeLispString(_options.Description);
        var closeAfter = _options.CloseAfter ? "T" : "NIL";

        return
            $"(dap:exe-run \"{folder}\" {_options.BatchSize} \"{mode}\" \"{exact}\" \"{description}\" {closeAfter})";
    }

    private void SendToActiveDoc(string command)
    {
        if (_activeDoc is null)
        {
            throw new InvalidOperationException("No active AutoCAD document to receive command.");
        }

        Console.WriteLine($"> {command}");
        _activeDoc.SendCommand(command + "\n");
    }

    private void WaitForIdle(TimeSpan timeout, string phase)
    {
        var start = DateTime.UtcNow;
        while (DateTime.UtcNow - start < timeout)
        {
            try
            {
                var state = _acad!.GetAcadState();
                bool isIdle = state.IsQuiescent;
                if (isIdle)
                {
                    return;
                }
            }
            catch
            {
                // Transient COM busy states are expected during command execution.
            }

            Thread.Sleep(350);
        }

        throw new TimeoutException($"Timed out waiting for AutoCAD idle during phase '{phase}'.");
    }

    private static string ToLispPath(string inputPath)
    {
        return Path.GetFullPath(inputPath).Replace('\\', '/');
    }

    private static string EscapeLispString(string value)
    {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }
}

internal static class ComInterop
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    public static bool TryGetActiveObject(string progId, out object? instance)
    {
        instance = null;
        var hr = CLSIDFromProgID(progId, out var clsid);
        if (hr < 0)
        {
            return false;
        }

        hr = GetActiveObject(ref clsid, IntPtr.Zero, out instance);
        return hr >= 0 && instance is not null;
    }
}
