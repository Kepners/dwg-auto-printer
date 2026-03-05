using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DwgAutoPrinter.App.Models;
using DwgAutoPrinter.App.Services;
using Forms = System.Windows.Forms;

namespace DwgAutoPrinter.App;

public partial class MainWindow : Window
{
    private readonly AutoCadBatchService _batchService = new();
    private readonly LispDeploymentService _lispDeploymentService = new();
    private CancellationTokenSource? _runCancellation;

    public ObservableCollection<LogEntry> MonitorLogs { get; } = [];
    public ObservableCollection<LogEntry> ErrorLogs { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        FolderPathTextBox.Text = Environment.CurrentDirectory;
        DescriptionTextBox.Text = "Construction Issue";
        AlwaysOnTopCheckBox.IsChecked = true;
        Topmost = true;
    }

    private void BrowseFolderButton_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Select DWG job root folder",
            SelectedPath = Directory.Exists(FolderPathTextBox.Text)
                ? FolderPathTextBox.Text
                : Environment.CurrentDirectory
        };

        var result = dialog.ShowDialog();
        if (result == Forms.DialogResult.OK && Directory.Exists(dialog.SelectedPath))
        {
            FolderPathTextBox.Text = dialog.SelectedPath;
        }
    }

    private void RevisionModeComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ExactRevisionTextBox is null)
        {
            return;
        }

        var mode = GetRevisionMode();
        ExactRevisionTextBox.IsEnabled = string.Equals(mode, "EXACT", StringComparison.OrdinalIgnoreCase);
    }

    private void AlwaysOnTopCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        Topmost = AlwaysOnTopCheckBox.IsChecked == true;
    }

    private async void DeployLispButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            SetStatus("Deploying LISP...", isRunning: true);
            var lispSource = ResolveBundledLispPath();
            var copied = await Task.Run(() => _lispDeploymentService.CopyToAutoCadLocations(lispSource, AppendLog));
            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Success,
                Message = $"LISP deployment complete. Copied to {copied.Count} location(s)."
            });
            SetStatus("Idle", isRunning: false);
        }
        catch (Exception ex)
        {
            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Error,
                Message = $"Deployment failed: {ex.Message}"
            });
            SetStatus("Error", isRunning: false);
        }
    }

    private async void StartButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_runCancellation is not null)
        {
            return;
        }

        try
        {
            var options = BuildRunOptions();
            _runCancellation = new CancellationTokenSource();

            SetStatus("Running", isRunning: true);
            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Info,
                Message = "Run started."
            });

            await _batchService.RunAsync(options, AppendLog, _runCancellation.Token);

            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Success,
                Message = "Run completed successfully."
            });
            SetStatus("Completed", isRunning: false);
        }
        catch (OperationCanceledException)
        {
            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Warning,
                Message = "Run cancelled by user."
            });
            SetStatus("Stopped", isRunning: false);
        }
        catch (Exception ex)
        {
            AppendLog(new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = LogLevel.Error,
                Message = $"Run failed: {ex.Message}"
            });
            SetStatus("Error", isRunning: false);
        }
        finally
        {
            _runCancellation?.Dispose();
            _runCancellation = null;
            SetStatus("Idle", isRunning: false);
        }
    }

    private void StopButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_runCancellation is null)
        {
            return;
        }

        SetStatus("Stopping...", isRunning: true);
        _runCancellation.Cancel();
        _batchService.RequestStop(AppendLog);
    }

    private RunOptions BuildRunOptions()
    {
        var folder = FolderPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(folder))
        {
            throw new InvalidOperationException("Folder path is required.");
        }

        if (!Directory.Exists(folder))
        {
            throw new DirectoryNotFoundException($"Folder not found: {folder}");
        }

        var batchText = (BatchSizeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "5";
        if (!int.TryParse(batchText, out var batchSize) || batchSize < 1)
        {
            batchSize = 5;
        }

        var mode = GetRevisionMode();
        var exactRevision = ExactRevisionTextBox.Text.Trim();
        if (string.Equals(mode, "EXACT", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(exactRevision))
        {
            throw new InvalidOperationException("Exact revision value is required when revision mode is EXACT.");
        }

        if (!int.TryParse(TimeoutMinutesTextBox.Text.Trim(), out var timeoutMinutes) || timeoutMinutes < 1)
        {
            timeoutMinutes = 30;
        }

        return new RunOptions
        {
            FolderPath = folder,
            BatchSize = batchSize,
            RevisionMode = mode,
            ExactRevision = exactRevision,
            Description = DescriptionTextBox.Text.Trim(),
            CloseAfterProcess = CloseAfterCheckBox.IsChecked == true,
            AutoCadVisible = AutoCadVisibleCheckBox.IsChecked == true,
            TimeoutMinutes = timeoutMinutes,
            LispSourcePath = ResolveBundledLispPath()
        };
    }

    private string GetRevisionMode()
    {
        return (RevisionModeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Trim().ToUpperInvariant() ?? "NEXT";
    }

    private string ResolveBundledLispPath()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "smart-revision-update.lsp");
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Bundled LISP file not found next to executable.", path);
        }

        return path;
    }

    private void SetStatus(string value, bool isRunning)
    {
        Dispatcher.Invoke(() =>
        {
            StatusText.Text = value;
            RunProgress.Visibility = isRunning ? Visibility.Visible : Visibility.Collapsed;
            RunProgress.IsIndeterminate = isRunning;
            StartButton.IsEnabled = !isRunning;
            StopButton.IsEnabled = isRunning;
            DeployLispButton.IsEnabled = !isRunning;
        });
    }

    private void AppendLog(LogEntry entry)
    {
        Dispatcher.Invoke(() =>
        {
            if (entry.Level is LogLevel.Error or LogLevel.Warning)
            {
                ErrorLogs.Add(entry);
                TrimCollection(ErrorLogs);
                if (ErrorLogs.Count > 0)
                {
                    ErrorListBox.ScrollIntoView(ErrorLogs[^1]);
                }
            }
            else
            {
                MonitorLogs.Add(entry);
                TrimCollection(MonitorLogs);
                if (MonitorLogs.Count > 0)
                {
                    MonitorListBox.ScrollIntoView(MonitorLogs[^1]);
                }
            }
        });
    }

    private static void TrimCollection(ObservableCollection<LogEntry> collection, int max = 1500)
    {
        while (collection.Count > max)
        {
            collection.RemoveAt(0);
        }
    }
}
