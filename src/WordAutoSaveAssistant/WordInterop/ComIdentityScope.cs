using System.Runtime.InteropServices;

namespace WordAutoSaveAssistant.WordInterop;

// Addresses are not durable IDs. Keep IUnknown alive until enumeration ends so
// a released proxy address cannot be reused by the next document.
internal sealed class ComIdentityScope : IDisposable
{
    private readonly Func<object, nint> _acquire;
    private readonly Action<nint> _release;
    private readonly HashSet<nint> _identities = new();

    public ComIdentityScope()
        : this(Marshal.GetIUnknownForObject, pointer => Marshal.Release(pointer)) { }

    internal ComIdentityScope(Func<object, nint> acquire, Action<nint> release)
    {
        _acquire = acquire;
        _release = release;
    }

    public bool TryAdd(object value)
    {
        nint identity = _acquire(value);
        if (_identities.Add(identity)) return true;
        _release(identity);
        return false;
    }

    public void Dispose()
    {
        foreach (nint identity in _identities) _release(identity);
        _identities.Clear();
    }
}
