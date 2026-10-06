namespace WordAutoSaveAssistant.Models;

public enum DocumentSaveStatus
{
    Saved,
    Unchanged,
    Skipped,
    Failed
}

public enum DocumentOutcomeSource
{
    Document,
    WordInstance,
    Assistant
}

public sealed record DocumentSaveOutcome(
    string DocumentName,
    DocumentSaveStatus Status,
    string Message,
    string? ErrorCode = null,
    string? FullPath = null,
    DocumentOutcomeSource Source = DocumentOutcomeSource.Document);

public sealed record SaveRoundResult(
    DateTimeOffset CompletedAt,
    int DetectedInstances,
    int DetectedDocuments,
    IReadOnlyList<DocumentSaveOutcome> Outcomes,
    bool Cancelled = false)
{
    public int SavedCount => Outcomes.Count(item => item.Status == DocumentSaveStatus.Saved);
    public int UnchangedCount => Outcomes.Count(item => item.Status == DocumentSaveStatus.Unchanged);
    public int SkippedCount => Outcomes.Count(item => item.Status == DocumentSaveStatus.Skipped);
    public int FailedCount => Outcomes.Count(item => item.Status == DocumentSaveStatus.Failed);

    public static SaveRoundResult Empty() =>
        new(DateTimeOffset.Now, 0, 0, Array.Empty<DocumentSaveOutcome>());
}
