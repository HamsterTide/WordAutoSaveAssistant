using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace WordAutoSaveAssistant.Tests;

// Explicit Debug-only diagnostic: fake data, disabled save controls, no Word calls.
internal static class UiRefreshBenchmark
{
    internal static int Run()
    {
#if DEBUG
        App app = new() { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        MainWindow window = new();
        window.LoadUiFixture();
        app.MainWindow = window;
        DispatcherFrame frame = new();
        app.Exit += (_, _) => frame.Continue = false;
        using Process process = Process.GetCurrentProcess();
        DependencyPropertyDescriptor brushProperty = DependencyPropertyDescriptor.FromProperty(
            Control.BackgroundProperty, typeof(Button));
        int brushChanges = 0;
        EventHandler brushChanged = (_, _) => brushChanges++;
        brushProperty.AddValueChanged(window.StartStopButton, brushChanged);
        long allocatedAtStart = 0;
        TimeSpan cpuAtStart = default;
        Stopwatch sample = new();
        bool measuring = false;
        DispatcherTimer timer = new(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        timer.Tick += (_, _) =>
        {
            if (!measuring)
            {
                measuring = true;
                brushChanges = 0;
                cpuAtStart = process.TotalProcessorTime;
                allocatedAtStart = GC.GetAllocatedBytesForCurrentThread();
                sample.Start();
                timer.Interval = TimeSpan.FromSeconds(30);
                return;
            }

            timer.Stop();
            sample.Stop();
            long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedAtStart;
            TimeSpan cpu = process.TotalProcessorTime - cpuAtStart;
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                Scenario = "stopped-visible-debug-fixture-no-word",
                ElapsedSeconds = sample.Elapsed.TotalSeconds,
                ProcessCpuSeconds = cpu.TotalSeconds,
                UiThreadAllocatedBytes = allocated,
                StartButtonBrushChanges = brushChanges
            }));
            window.Close();
        };
        try
        {
            window.Show();
            timer.Start();
            Dispatcher.PushFrame(frame);
            return measuring ? 0 : 1;
        }
        finally
        {
            timer.Stop();
            brushProperty.RemoveValueChanged(window.StartStopButton, brushChanged);
        }
#else
        Console.Error.WriteLine("此诊断仅支持 Debug；不会在 Release 中连接或操作 Word。");
        return 2;
#endif
    }
}
