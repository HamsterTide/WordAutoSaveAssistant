using System.Runtime.InteropServices;

namespace WordAutoSaveAssistant.Tests;

public sealed class FakeWordApplication
{
    public FakeWordApplication(params FakeWordDocument[] documents)
    {
        _documents = new FakeWordDocuments(documents);
    }

    public string Name => "Microsoft Word";
    private readonly FakeWordDocuments _documents;
    public bool BusyDocuments { get; set; }
    public FakeWordDocuments Documents => BusyDocuments
        ? throw new COMException("模拟文档集合忙碌", unchecked((int)0x80010001)) : _documents;
}

public sealed class FakeWordDocuments
{
    private readonly List<FakeWordDocument> _documents;

    public FakeWordDocuments(FakeWordDocument[] documents)
    {
        _documents = documents.ToList();
    }

    public int Count => _documents.Count;
    public FakeWordDocument this[int oneBasedIndex] => _documents[oneBasedIndex - 1];
    public void Remove(FakeWordDocument document) => _documents.Remove(document);
}

public sealed class FakeWordDocument
{
    private readonly string _fullName;
    private readonly bool _readOnly;

    public FakeWordDocument(
        string name,
        string path,
        string fullName,
        bool saved,
        bool readOnly = false,
        int protectionType = -1,
        bool throwOnFullName = false,
        bool throwOnReadOnly = false)
    {
        Name = name;
        Path = path;
        _fullName = fullName;
        _saved = saved;
        _readOnly = readOnly;
        ProtectionType = protectionType;
        ThrowOnFullName = throwOnFullName;
        ThrowOnReadOnly = throwOnReadOnly;
    }

    public string Name { get; }
    public string Path { get; }
    private bool _saved;
    public bool BusyConfirmation { get; set; }
    public bool Saved => BusyConfirmation && CompletedSaves > 0
        ? throw new COMException("模拟保存后确认忙碌", unchecked((int)0x8001010A)) : _saved;
    public int ProtectionType { get; }
    public bool ThrowOnFullName { get; }
    public bool ThrowOnReadOnly { get; }

    public string FullName => ThrowOnFullName
        ? throw new COMException("模拟 FullName 读取失败", unchecked((int)0x800A0001))
        : _fullName;

    public bool ReadOnly => ThrowOnReadOnly
        ? throw new COMException("模拟 ReadOnly 读取失败", unchecked((int)0x800A0002))
        : _readOnly;

    public int SaveCalls { get; private set; }
    public Action? AfterSave { get; set; }
    public bool ConfirmSave { get; set; } = true;
    public bool FailSave { get; set; }
    public bool BusySave { get; set; }
    public int CompletedSaves { get; private set; }

    public void Save()
    {
        SaveCalls++;
        if (BusySave) throw new COMException("模拟 Word 忙碌", unchecked((int)0x8001010A));
        if (FailSave) throw new COMException("模拟磁盘保存失败", unchecked((int)0x800A0003));
        CompletedSaves++;
        _saved = ConfirmSave;
        AfterSave?.Invoke();
    }
}
