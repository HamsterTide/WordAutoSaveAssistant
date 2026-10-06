using Microsoft.Win32;

namespace WordAutoSaveAssistant.Services;

internal sealed class SystemSessionEvents : IDisposable
{
    private bool _started;

    public event Action<bool>? SessionLockedChanged;
    public event Action<bool>? SystemSuspendedChanged;

    public void Start()
    {
        if (_started)
        {
            return;
        }

        SystemEvents.SessionSwitch += OnSessionSwitch;
        SystemEvents.PowerModeChanged += OnPowerModeChanged;
        _started = true;
    }

    public void Dispose()
    {
        if (!_started)
        {
            return;
        }

        SystemEvents.SessionSwitch -= OnSessionSwitch;
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        _started = false;
    }

    private void OnSessionSwitch(object sender, SessionSwitchEventArgs args)
    {
        if (args.Reason == SessionSwitchReason.SessionLock)
        {
            SessionLockedChanged?.Invoke(true);
        }
        else if (args.Reason == SessionSwitchReason.SessionUnlock)
        {
            SessionLockedChanged?.Invoke(false);
        }
    }

    private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs args)
    {
        if (args.Mode == PowerModes.Suspend)
        {
            SystemSuspendedChanged?.Invoke(true);
        }
        else if (args.Mode == PowerModes.Resume)
        {
            SystemSuspendedChanged?.Invoke(false);
        }
    }
}
