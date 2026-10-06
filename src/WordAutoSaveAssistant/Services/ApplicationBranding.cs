using System.IO;

namespace WordAutoSaveAssistant.Services;

internal static class ApplicationBranding
{
    internal static System.Drawing.Icon LoadTrayIcon()
    {
        using Stream stream = typeof(ApplicationBranding).Assembly.GetManifestResourceStream(
            "WordAutoSaveAssistant.Assets.AppIcon.ico")
            ?? throw new InvalidOperationException("应用图标资源缺失。");
        using System.Drawing.Icon source = new(stream, 32, 32);
        return (System.Drawing.Icon)source.Clone();
    }
}
