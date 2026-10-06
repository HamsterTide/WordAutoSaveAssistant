using System.Text;
using WordAutoSaveAssistant.Models;

namespace WordAutoSaveAssistant.Core;

public static class ResultPresentation
{
    public static RoundDisplaySnapshot CreateSnapshot(SaveRoundResult result)
    {
        ArgumentNullException.ThrowIfNull(result);

        List<(DocumentSaveOutcome Outcome, int Order)> indexed = result.Outcomes
            .Select((outcome, index) => (outcome, index))
            .ToList();

        return new RoundDisplaySnapshot(
            true,
            result.CompletedAt,
            result.DetectedInstances,
            result.DetectedDocuments,
            CreateSection(indexed, DocumentSaveStatus.Saved),
            CreateSection(indexed, DocumentSaveStatus.Unchanged),
            CreateSection(indexed, DocumentSaveStatus.Skipped),
            CreateSection(indexed, DocumentSaveStatus.Failed));
    }

    private static ResultSectionSnapshot CreateSection(
        IEnumerable<(DocumentSaveOutcome Outcome, int Order)> indexed,
        DocumentSaveStatus status)
    {
        List<DocumentResultListItem> items = indexed
            .Where(item => item.Outcome.Status == status)
            .Select(item => CreateItem(item.Outcome, item.Order))
            .OrderBy(item => item, DocumentResultListItemComparer.Instance)
            .ToList();

        return items.Count == 0
            ? RoundDisplaySnapshot.PlaceholderSection("无")
            : new ResultSectionSnapshot(items.Count, items);
    }

    private static DocumentResultListItem CreateItem(DocumentSaveOutcome outcome, int order)
    {
        string displayName = outcome.Source switch
        {
            DocumentOutcomeSource.WordInstance => string.IsNullOrWhiteSpace(outcome.DocumentName) ? "(Word 实例)" : outcome.DocumentName,
            DocumentOutcomeSource.Assistant => "(助手)",
            _ => string.IsNullOrWhiteSpace(outcome.DocumentName) ? "(未命名文档)" : outcome.DocumentName
        };

        StringBuilder helpText = new();
        helpText.Append(string.IsNullOrWhiteSpace(outcome.FullPath) ? displayName : outcome.FullPath);
        if (outcome.Status is DocumentSaveStatus.Skipped or DocumentSaveStatus.Failed)
        {
            helpText.AppendLine();
            helpText.Append("原因：");
            helpText.Append(outcome.Message);
            if (!string.IsNullOrWhiteSpace(outcome.ErrorCode))
            {
                helpText.Append("（");
                helpText.Append(outcome.ErrorCode);
                helpText.Append('）');
            }
        }

        return new DocumentResultListItem(
            displayName,
            helpText.ToString(),
            outcome.FullPath,
            outcome.Source,
            order);
    }
}

public sealed class RoundDisplayState
{
    public RoundDisplaySnapshot Snapshot { get; private set; } = RoundDisplaySnapshot.Initial;
    public int ReplacementCount { get; private set; }

    public bool TryApply(SaveRoundResult result, long completedGeneration, long currentGeneration)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (result.Cancelled || completedGeneration != currentGeneration)
        {
            return false;
        }

        Snapshot = ResultPresentation.CreateSnapshot(result);
        ReplacementCount++;
        return true;
    }
}

internal sealed class DocumentResultListItemComparer : IComparer<DocumentResultListItem>
{
    public static DocumentResultListItemComparer Instance { get; } = new();

    public int Compare(DocumentResultListItem? left, DocumentResultListItem? right)
    {
        if (ReferenceEquals(left, right)) return 0;
        if (left is null) return -1;
        if (right is null) return 1;

        int source = left.Source.CompareTo(right.Source);
        if (source != 0) return source;

        int name = NaturalTextComparer.Compare(left.DisplayName, right.DisplayName);
        if (name != 0) return name;

        int path = StringComparer.OrdinalIgnoreCase.Compare(left.FullPath, right.FullPath);
        if (path != 0) return path;

        return left.OriginalOrder.CompareTo(right.OriginalOrder);
    }
}

internal static class NaturalTextComparer
{
    public static int Compare(string? left, string? right)
    {
        left ??= string.Empty;
        right ??= string.Empty;
        int leftIndex = 0;
        int rightIndex = 0;

        while (leftIndex < left.Length && rightIndex < right.Length)
        {
            bool leftDigit = char.IsDigit(left[leftIndex]);
            bool rightDigit = char.IsDigit(right[rightIndex]);
            if (leftDigit && rightDigit)
            {
                int result = CompareDigitRuns(left, ref leftIndex, right, ref rightIndex);
                if (result != 0) return result;
                continue;
            }

            if (leftDigit != rightDigit)
            {
                return char.ToUpperInvariant(left[leftIndex]).CompareTo(char.ToUpperInvariant(right[rightIndex]));
            }

            int character = char.ToUpperInvariant(left[leftIndex])
                .CompareTo(char.ToUpperInvariant(right[rightIndex]));
            if (character != 0) return character;
            leftIndex++;
            rightIndex++;
        }

        return (left.Length - leftIndex).CompareTo(right.Length - rightIndex);
    }

    private static int CompareDigitRuns(
        string left,
        ref int leftIndex,
        string right,
        ref int rightIndex)
    {
        int leftStart = leftIndex;
        int rightStart = rightIndex;
        while (leftIndex < left.Length && char.IsDigit(left[leftIndex])) leftIndex++;
        while (rightIndex < right.Length && char.IsDigit(right[rightIndex])) rightIndex++;

        int leftSignificant = leftStart;
        int rightSignificant = rightStart;
        while (leftSignificant < leftIndex - 1 && left[leftSignificant] == '0') leftSignificant++;
        while (rightSignificant < rightIndex - 1 && right[rightSignificant] == '0') rightSignificant++;

        int leftLength = leftIndex - leftSignificant;
        int rightLength = rightIndex - rightSignificant;
        int length = leftLength.CompareTo(rightLength);
        if (length != 0) return length;

        for (int offset = 0; offset < leftLength; offset++)
        {
            int digit = left[leftSignificant + offset].CompareTo(right[rightSignificant + offset]);
            if (digit != 0) return digit;
        }

        int leftRawLength = leftIndex - leftStart;
        int rightRawLength = rightIndex - rightStart;
        return leftRawLength.CompareTo(rightRawLength);
    }
}

public sealed class RecentEventHistory
{
    private const int Capacity = 5;
    private readonly List<string> _items = new();
    private long _nextSequence;

    public IReadOnlyList<string> Items => _items;

    public void AddImmediate(string message)
    {
        _nextSequence++;
        _items.Insert(0, message);
        Trim();
    }

    public void AddRound(SaveRoundResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        List<string> messages = result.Outcomes
            .Where(item => item.Status is DocumentSaveStatus.Skipped or DocumentSaveStatus.Failed)
            .Select(outcome => FormatRoundEvent(result, outcome))
            .ToList();

        if (messages.Count == 0) return;
        _nextSequence++;
        _items.InsertRange(0, messages);
        Trim();
    }

    private static string FormatRoundEvent(SaveRoundResult result, DocumentSaveOutcome outcome)
    {
        string status = outcome.Status == DocumentSaveStatus.Skipped ? "跳过" : "失败";
        string error = string.IsNullOrWhiteSpace(outcome.ErrorCode) ? string.Empty : $"（{outcome.ErrorCode}）";
        return $"{result.CompletedAt.LocalDateTime:HH:mm:ss}  {status}｜{outcome.DocumentName}：{outcome.Message}{error}";
    }

    private void Trim()
    {
        if (_items.Count > Capacity)
        {
            _items.RemoveRange(Capacity, _items.Count - Capacity);
        }
    }
}
