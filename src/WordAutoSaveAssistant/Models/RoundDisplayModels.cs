namespace WordAutoSaveAssistant.Models;

public sealed record DocumentResultListItem(
    string DisplayName,
    string HelpText,
    string? FullPath,
    DocumentOutcomeSource Source,
    int OriginalOrder,
    bool IsPlaceholder = false)
{
    public string AutomationName => DisplayName;
}

public sealed record ResultSectionSnapshot(
    int Count,
    IReadOnlyList<DocumentResultListItem> Items);

public sealed record RoundDisplaySnapshot(
    bool HasCompletedRound,
    DateTimeOffset? CompletedAt,
    int DetectedInstances,
    int DetectedDocuments,
    ResultSectionSnapshot Saved,
    ResultSectionSnapshot Unchanged,
    ResultSectionSnapshot Skipped,
    ResultSectionSnapshot Failed)
{
    public static RoundDisplaySnapshot Initial { get; } = new(
        false,
        null,
        0,
        0,
        PlaceholderSection("尚未执行"),
        PlaceholderSection("尚未执行"),
        PlaceholderSection("尚未执行"),
        PlaceholderSection("尚未执行"));

    public string DetectionSummary => HasCompletedRound
        ? $"最近一轮 {CompletedAt?.LocalDateTime:HH:mm:ss} · {DetectedInstances} 个可访问 Word 实例、{DetectedDocuments} 个文档"
        : "尚未执行保存轮次；实时检测请查看“当前文档”。";

    internal static ResultSectionSnapshot PlaceholderSection(string text) =>
        new(0, [new DocumentResultListItem(text, text, null, DocumentOutcomeSource.Assistant, 0, true)]);
}
