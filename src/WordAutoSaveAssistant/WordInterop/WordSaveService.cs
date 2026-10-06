using WordAutoSaveAssistant.Models;

namespace WordAutoSaveAssistant.WordInterop;

internal sealed class WordSaveService
{
    private readonly ComStaHost _host = new();
    private readonly WordSaveEngine _engine = new();

    public bool IsBusy => _host.IsBusy;

    public Task<WordScanResult> SaveAllAsync(CancellationToken cancellationToken, IProgress<string>? progress = null) =>
        _host.InvokeAsync(token => _engine.SaveAndInspectAll(token, progress), cancellationToken);

    public Task<WordInventorySnapshot> InspectAllAsync(CancellationToken cancellationToken) =>
        _host.InvokeAsync(_engine.InspectAll, cancellationToken);

    public Task<bool> TryShutdownAsync(TimeSpan timeout) =>
        _host.TryShutdownAsync(timeout);
}
