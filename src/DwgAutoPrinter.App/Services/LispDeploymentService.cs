using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Principal;
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
        var programFilesRoot = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var isAdmin = IsRunningAsAdministrator();
        var skippedProgramFiles = 0;

        foreach (var directory in targetDirectories)
        {
            if (!isAdmin && directory.StartsWith(programFilesRoot, StringComparison.OrdinalIgnoreCase))
            {
                skippedProgramFiles++;
                continue;
            }

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
                    Level = LogLevel.Warning,
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

        if (skippedProgramFiles > 0)
        {
            log(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Warning,
                Message = $"Skipped {skippedProgramFiles} Program Files target(s). Run as Administrator to deploy there."
            });
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

    private static bool IsRunningAsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
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
