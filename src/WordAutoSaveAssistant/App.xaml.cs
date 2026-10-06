using System.Diagnostics;
using System.Windows;
using WordAutoSaveAssistant.Services;

namespace WordAutoSaveAssistant;

public partial class App : System.Windows.Application
{
    private SingleInstanceGuard? _singleInstance;
    private MainWindow? _mainWindow;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _singleInstance = new SingleInstanceGuard(Process.GetCurrentProcess().SessionId);
        if (!_singleInstance.IsPrimary)
        {
            System.Windows.MessageBox.Show(
                "Word 定时保存助手已经在运行。",
                "Word 定时保存助手",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            Shutdown();
            return;
        }

        _mainWindow = new MainWindow();
#if DEBUG
        if (e.Args.Contains("--ui-fixture", StringComparer.OrdinalIgnoreCase))
        {
            _mainWindow.LoadUiFixture();
        }
#endif
        MainWindow = _mainWindow;
        _mainWindow.Show();
    }

    internal void CompleteShutdown()
    {
        _singleInstance?.Dispose();
        _singleInstance = null;
        Shutdown();
    }
}
