using System.Text.Json;
using System.IO;

namespace WordAutoSaveAssistant.Services;

public sealed class SettingsStore
{
    public const int DefaultIntervalMinutes = 10;
    private const string TempPrefix = "settings.json.";
    private readonly string _directory;
    private readonly string _path;

    public SettingsStore(string? directory = null)
    {
        _directory = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Word定时保存助手");
        _path = Path.Combine(_directory, "settings.json");
    }

    public int LoadInterval()
    {
        CleanupTemporaryFiles();
        try
        {
            if (!File.Exists(_path))
            {
                return DefaultIntervalMinutes;
            }

            string json = File.ReadAllText(_path);
            SettingsDocument? settings = JsonSerializer.Deserialize<SettingsDocument>(json);
            return settings?.IntervalMinutes is >= 1 and <= 1440
                ? settings.IntervalMinutes
                : DefaultIntervalMinutes;
        }
        catch (JsonException)
        {
            return DefaultIntervalMinutes;
        }
        catch (IOException)
        {
            return DefaultIntervalMinutes;
        }
        catch (UnauthorizedAccessException)
        {
            return DefaultIntervalMinutes;
        }
    }

    public bool TrySaveInterval(int intervalMinutes, out string? error)
    {
        if (intervalMinutes is < 1 or > 1440)
        {
            error = "保存间隔必须在 1 到 1440 分钟之间。";
            return false;
        }

        string tempPath = Path.Combine(
            _directory,
            $"{TempPrefix}{Environment.ProcessId}.{Guid.NewGuid():N}.tmp");

        try
        {
            Directory.CreateDirectory(_directory);
            string json = JsonSerializer.Serialize(
                new SettingsDocument(intervalMinutes),
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(tempPath, json);

            if (File.Exists(_path))
            {
                File.Replace(tempPath, _path, null, true);
            }
            else
            {
                File.Move(tempPath, _path, false);
            }

            error = null;
            return true;
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            TryDelete(tempPath);
            error = exception.Message;
            return false;
        }
    }

    private void CleanupTemporaryFiles()
    {
        try
        {
            if (!Directory.Exists(_directory))
            {
                return;
            }

            foreach (string file in Directory.EnumerateFiles(_directory, $"{TempPrefix}*.tmp"))
            {
                TryDelete(file);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private sealed record SettingsDocument(int IntervalMinutes);
}
