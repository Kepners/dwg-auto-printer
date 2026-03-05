using System;
using System.Collections.Generic;
using System.IO;
using DwgAutoPrinter.App.Models;

namespace DwgAutoPrinter.App.Services;

public sealed class LispDeploymentService
{
    public IReadOnlyList<string> CopyToAutoCadLocations(string lispPath, Action<LogEntry> log)
    {
        if (!File.Exists(lispPath))
        {
            throw new FileNotFoundException("LISP file not found.", lispPath);
        }

        var targetDirectories = DiscoverTargetDirectories();
        var copiedFiles = new List<string>();

        foreach (var directory in targetDirectories)
        {
            var destination = Path.Combine(directory, "smart-revision-update.lsp");
            try
            {
                Directory.CreateDirectory(directory);
                File.Copy(lispPath, destination, overwrite: true);
                copiedFiles.Add(destination);
                log(new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = LogLevel.Success,
                    Message = $"LISP copied: {destination}"
                });
            }
            catch (UnauthorizedAccessException ex)
            {
                log(new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = LogLevel.Error,
                    Message = $"Permission denied copying to {destination}: {ex.Message}"
                });
            }
            catch (Exception ex)
            {
                log(new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = LogLevel.Error,
                    Message = $"Copy failed for {destination}: {ex.Message}"
                });
            }
        }

        if (copiedFiles.Count == 0)
        {
            log(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Warning,
                Message = "No AutoCAD locations were updated. Run the app as Administrator for Program Files copy."
            });
        }

        return copiedFiles;
    }

    private static IEnumerable<string> DiscoverTargetDirectories()
    {
        var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var autodeskRoot = Path.Combine(programFiles, "Autodesk");
        if (Directory.Exists(autodeskRoot))
        {
            foreach (var dir in Directory.GetDirectories(autodeskRoot, "AutoCAD *", SearchOption.TopDirectoryOnly))
            {
                targets.Add(dir);

                var support = Path.Combine(dir, "Support");
                if (Directory.Exists(support))
                {
                    targets.Add(support);
                }
            }
        }

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var roamingAutodesk = Path.Combine(appData, "Autodesk");
        if (Directory.Exists(roamingAutodesk))
        {
            foreach (var cadDir in Directory.GetDirectories(roamingAutodesk, "AutoCAD *", SearchOption.TopDirectoryOnly))
            {
                foreach (var releaseDir in Directory.GetDirectories(cadDir, "R*", SearchOption.TopDirectoryOnly))
                {
                    var support = Path.Combine(releaseDir, "enu", "Support");
                    if (Directory.Exists(support))
                    {
                        targets.Add(support);
                    }
                }
            }
        }

        return targets;
    }
}
