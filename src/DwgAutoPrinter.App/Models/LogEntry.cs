namespace DwgAutoPrinter.App.Models;

public enum LogLevel
{
    Info,
    Success,
    Warning,
    Error
}

public sealed class LogEntry
{
    public required DateTime Timestamp { get; init; }
    public required string Message { get; init; }
    public required LogLevel Level { get; init; }
}
