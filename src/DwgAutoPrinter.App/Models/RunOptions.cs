namespace DwgAutoPrinter.App.Models;

public sealed class RunOptions
{
    public required string FolderPath { get; init; }
    public required int BatchSize { get; init; }
    public required string RevisionMode { get; init; }
    public required string ExactRevision { get; init; }
    public required string Description { get; init; }
    public required bool CloseAfterProcess { get; init; }
    public required bool AutoCadVisible { get; init; }
    public required int TimeoutMinutes { get; init; }
    public required string LispSourcePath { get; init; }
}
