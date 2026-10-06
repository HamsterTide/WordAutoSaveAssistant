namespace WordAutoSaveAssistant.Models;

public enum WordConnectionIssue
{
    None,
    PermissionMismatch,
    Busy,
    AccessDenied,
    ObjectModelUnavailable,
    Unknown
}

public sealed record WordInventoryItem(
    string DocumentName,
    string? FullPath,
    string State,
    string Detail,
    bool CanAutoSave,
    bool? HasUnsavedChanges,
    DocumentOutcomeSource Source = DocumentOutcomeSource.Document,
    bool IsSkipped = false,
    WordConnectionIssue ConnectionIssue = WordConnectionIssue.None)
{
    public string HelpText => $"{FullPath ?? DocumentName}\n{Detail}";
}

public sealed record WordInventorySnapshot(
    DateTimeOffset CheckedAt,
    int DetectedInstances,
    int DetectedDocuments,
    IReadOnlyList<WordInventoryItem> Items,
    bool Cancelled = false)
{
    public int UnavailableCount => Items.Count(item => item.Source == DocumentOutcomeSource.WordInstance);
    public int EligibleCount => Items.Count(item => item.Source == DocumentOutcomeSource.Document && item.CanAutoSave);
    public int PendingSaveCount => Items.Count(item => item.Source == DocumentOutcomeSource.Document && item.CanAutoSave && item.HasUnsavedChanges == true);
    public int SkippedCount => Items.Count(item => item.Source == DocumentOutcomeSource.Document && item.IsSkipped);
    public int DocumentErrorCount => Items.Count(item => item.Source == DocumentOutcomeSource.Document && !item.CanAutoSave && !item.IsSkipped);
    public int DetectionErrorCount => Items.Count(item => item.Source == DocumentOutcomeSource.Assistant);
    public string Summary => $"可连接 {DetectedInstances} 个实例 · 已检测 {DetectedDocuments} 个文档\n"
        + $"符合保存条件 {EligibleCount}（有修改 {PendingSaveCount}） · 跳过 {SkippedCount} · 文档异常 {DocumentErrorCount} · 无法连接 {UnavailableCount}"
        + (DetectionErrorCount > 0 ? $" · 检测异常 {DetectionErrorCount}（清单可能不完整）" : string.Empty);

    public string ConnectionWarning
    {
        get
        {
            if (UnavailableCount == 0) return string.Empty;
            List<string> reasons = new();
            Add(WordConnectionIssue.PermissionMismatch, "权限差异已确认");
            Add(WordConnectionIssue.Busy, "Word 暂时忙碌");
            Add(WordConnectionIssue.AccessDenied, "拒绝访问，权限差异未确认");
            Add(WordConnectionIssue.ObjectModelUnavailable, "对象模型不可用");
            int other = Items.Count(item => item.Source == DocumentOutcomeSource.WordInstance
                && item.ConnectionIssue is WordConnectionIssue.Unknown or WordConnectionIssue.None);
            if (other > 0) reasons.Add($"其他原因未确认 {other}");
            string guidance = Items.Any(item => item.Source == DocumentOutcomeSource.WordInstance
                && item.ConnectionIssue == WordConnectionIssue.PermissionMismatch)
                ? "已确认权限不同的窗口需手动对齐权限；其他项请查看具体原因。"
                : "不能将连接失败直接归因于权限，请查看具体原因。";
            return $"部分 Word 连接项当前不能自动保存：{string.Join("；", reasons)}。{guidance}悬停查看详情。";

            void Add(WordConnectionIssue issue, string label)
            {
                int count = Items.Count(item => item.Source == DocumentOutcomeSource.WordInstance && item.ConnectionIssue == issue);
                if (count > 0) reasons.Add($"{label} {count}");
            }
        }
    }
}

public sealed record WordScanResult(SaveRoundResult Round, WordInventorySnapshot Inventory);
