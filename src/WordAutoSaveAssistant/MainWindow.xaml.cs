using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WordAutoSaveAssistant.Core;
using WordAutoSaveAssistant.Models;
using WordAutoSaveAssistant.Services;
using WordAutoSaveAssistant.WordInterop;
using Forms = System.Windows.Forms;
using MediaColor = System.Windows.Media.Color;

namespace WordAutoSaveAssistant;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private static readonly SolidColorBrush BlueBrush = CreateFrozenBrush(37, 99, 235);
    private static readonly SolidColorBrush RedBrush = CreateFrozenBrush(217, 45, 32);
    private static readonly SolidColorBrush AmberBrush = CreateFrozenBrush(247, 144, 9);
    private static readonly SolidColorBrush GreenBrush = CreateFrozenBrush(18, 183, 106);
    private static readonly SolidColorBrush GrayBrush = CreateFrozenBrush(152, 162, 179);
    private readonly CountdownRefreshState _countdownRefresh = new();
    private readonly SettingsStore _settingsStore = new();
    private readonly MonotonicClock _clock = new();
    private readonly SaveScheduleState _schedule;
    private readonly WordSaveService _wordSaveService = new();
    private readonly SystemSessionEvents _systemEvents = new();
    private readonly DispatcherTimer _uiTimer;
    private readonly RoundDisplayState _roundDisplayState = new();
    private readonly RecentEventHistory _eventHistory = new();
    private readonly ObservableCollection<string> _recentEvents = new();
    private readonly Forms.NotifyIcon _trayIcon;
    private readonly System.Drawing.Icon _applicationIcon;
    private readonly Forms.ContextMenuStrip _trayMenu;
    private readonly Forms.ToolStripMenuItem _showMenuItem;
    private readonly Forms.ToolStripMenuItem _exitMenuItem;
    private CancellationTokenSource? _runCancellation;
    private CancellationTokenSource? _attemptCancellation;
    private CancellationTokenSource? _inspectionCancellation;
    private string? _saveProgress;
    private string? _settingsWarning;
    private bool _allowClose;
    private bool _isExiting;
    private bool _initializing = true;
#if DEBUG
    private bool _isUiFixture;
#endif
    private RoundDisplaySnapshot _roundSnapshot = RoundDisplaySnapshot.Initial;
    private readonly CancellationTokenSource _lifetimeCancellation = new();
    private WordInventorySnapshot? _inventorySnapshot;
    private bool _inspectionInProgress;
    private long _nextInspectionAt;

    public WordInventorySnapshot? InventorySnapshot
    {
        get => _inventorySnapshot;
        private set
        {
            _inventorySnapshot = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(InventorySnapshot)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public RoundDisplaySnapshot RoundSnapshot
    {
        get => _roundSnapshot;
        private set
        {
            if (ReferenceEquals(_roundSnapshot, value))
            {
                return;
            }

            _roundSnapshot = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RoundSnapshot)));
        }
    }

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;

        int interval = _settingsStore.LoadInterval();
        _schedule = new SaveScheduleState(_clock, interval);
        IntervalBox.Text = interval.ToString();
        RecentEventsList.ItemsSource = _recentEvents;

        _trayMenu = new Forms.ContextMenuStrip();
        _showMenuItem = new Forms.ToolStripMenuItem("显示主窗口");
        _exitMenuItem = new Forms.ToolStripMenuItem("退出");
        _showMenuItem.Click += ShowMenuItem_Click;
        _exitMenuItem.Click += ExitMenuItem_Click;
        _trayMenu.Items.AddRange([_showMenuItem, _exitMenuItem]);

        _applicationIcon = ApplicationBranding.LoadTrayIcon();
        _trayIcon = new Forms.NotifyIcon
        {
            Icon = _applicationIcon,
            Text = "Word 定时保存助手",
            ContextMenuStrip = _trayMenu,
            Visible = true
        };
        _trayIcon.DoubleClick += TrayIcon_DoubleClick;

        _systemEvents.SessionLockedChanged += SystemEvents_SessionLockedChanged;
        _systemEvents.SystemSuspendedChanged += SystemEvents_SystemSuspendedChanged;
        try
        {
            _systemEvents.Start();
        }
        catch (Exception exception)
        {
            AddRecentEvent($"系统事件监听未启动：{exception.Message}");
        }

        _uiTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _uiTimer.Tick += UiTimer_Tick;
        _uiTimer.Start();

        _initializing = false;
        UpdateUi();
        Loaded += (_, _) =>
        {
            UpdateUi();
            _ = RefreshInventoryAsync();
        };
    }

#if DEBUG
    internal void LoadUiFixture()
    {
        _isUiFixture = true;
        DateTimeOffset now = DateTimeOffset.Now;
        SaveRoundResult fixture = new(
            now,
            2,
            15,
            [
                new("报告2.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"C:\资料\项目A\报告2.docx"),
                new("报告10.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"C:\资料\项目A\报告10.docx"),
                new("2026年度特别长的中文项目阶段性总结与后续工作安排说明.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"D:\很长的目录名称\2026年度特别长的中文项目阶段性总结与后续工作安排说明.docx"),
                new("report003.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"C:\English\report003.docx"),
                new("分析1.docx", DocumentSaveStatus.Unchanged, "没有未保存的修改", FullPath: @"D:\A\分析1.docx"),
                new("分析01.docx", DocumentSaveStatus.Unchanged, "没有未保存的修改", FullPath: @"D:\A\分析01.docx"),
                new("same-name.docx", DocumentSaveStatus.Unchanged, "没有未保存的修改", FullPath: @"C:\一组\same-name.docx"),
                new("same-name.docx", DocumentSaveStatus.Unchanged, "没有未保存的修改", FullPath: @"D:\二组\same-name.docx"),
                new("未命名文档1", DocumentSaveStatus.Skipped, "未命名文档没有现有保存路径"),
                new("只读材料.docx", DocumentSaveStatus.Skipped, "文档为只读状态", FullPath: @"C:\资料\只读材料.docx"),
                new("受保护材料.docx", DocumentSaveStatus.Skipped, "文档处于受保护状态", FullPath: @"C:\资料\受保护材料.docx"),
                new("云端同步中的超长英文文件名-document-being-synchronized-and-temporarily-read-only.docx", DocumentSaveStatus.Skipped, "文档为只读状态", FullPath: @"C:\Cloud\云端同步中的超长英文文件名-document-being-synchronized-and-temporarily-read-only.docx"),
                new("损坏文档.docx", DocumentSaveStatus.Failed, "Word 返回错误", "0x800A175D", @"D:\异常\损坏文档.docx"),
                new("(Word 实例)", DocumentSaveStatus.Failed, "Word 正忙，本轮未能完成", "0x8001010A", Source: DocumentOutcomeSource.WordInstance),
                new("(助手)", DocumentSaveStatus.Failed, "保存轮次异常：调试验收示例", "0x80131500", Source: DocumentOutcomeSource.Assistant)
            ]);

        ApplyResult(fixture, notifyFailure: false);
        InventorySnapshot = new(now, 2, 15, fixture.Outcomes.Select(outcome => new WordInventoryItem(
            outcome.DocumentName, outcome.FullPath,
            outcome.Status == DocumentSaveStatus.Failed ? "无法连接" : outcome.Status == DocumentSaveStatus.Skipped ? "不自动保存" : "无未保存修改",
            outcome.Message, outcome.Status is DocumentSaveStatus.Saved or DocumentSaveStatus.Unchanged,
            false, outcome.Source, IsSkipped: outcome.Status == DocumentSaveStatus.Skipped,
            ConnectionIssue: outcome.Source == DocumentOutcomeSource.WordInstance
                ? outcome.ErrorCode is "0x80010001" or "0x8001010A" ? WordConnectionIssue.Busy : WordConnectionIssue.Unknown
                : WordConnectionIssue.None)).ToArray());
        UpdateUi();
    }
#endif

    private void StartStopButton_Click(object sender, RoutedEventArgs e)
    {
#if DEBUG
        if (_isUiFixture)
        {
            return;
        }
#endif
        if (_schedule.IsRunning || _schedule.IsSaving)
        {
            StopSchedule();
            return;
        }

        if (!TryReadInterval(out int interval))
        {
            return;
        }

        PersistInterval(interval);
        _schedule.ChangeInterval(interval, _clock.Timestamp);
        _runCancellation?.Dispose();
        _runCancellation = new CancellationTokenSource();
        ScheduleAction action = _schedule.Start(_clock.Timestamp);
        UpdateUi();
        QueueSave(action);
    }

    private void StopSchedule()
    {
        _schedule.Stop();
        _runCancellation?.Cancel();
        _attemptCancellation?.Cancel();
        UpdateUi();
    }

    private void SaveOnceButton_Click(object sender, RoutedEventArgs e)
    {
        ScheduleAction action = _schedule.SaveOnce();
        if (!action.QueueSave) return;
        if (!_schedule.IsRunning)
        {
            _runCancellation?.Dispose();
            _runCancellation = new CancellationTokenSource();
        }
        UpdateUi();
        QueueSave(action);
    }

    private void IntervalBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_initializing || _schedule is null)
        {
            return;
        }

        if (TryReadInterval(out int interval))
        {
            IntervalErrorText.Visibility = Visibility.Collapsed;
            StartStopButton.IsEnabled = true;
            if (interval != _schedule.IntervalMinutes)
            {
                _schedule.ChangeInterval(interval, _clock.Timestamp);
                PersistInterval(interval);
            }
        }
        else
        {
            IntervalErrorText.Text = "请输入 1–1440 之间的整数分钟。";
            IntervalErrorText.Visibility = Visibility.Visible;
            StartStopButton.IsEnabled = _schedule.IsRunning;
        }

        UpdateUi();
    }

    private bool TryReadInterval(out int interval) =>
        int.TryParse(IntervalBox.Text, out interval) && interval is >= 1 and <= 1440;

    private void PersistInterval(int interval)
    {
        if (!_settingsStore.TrySaveInterval(interval, out string? error))
        {
            _settingsWarning = $"设置未能写入，本次仍使用 {interval} 分钟；下次启动可能恢复为此前设置或默认值。";
            AddRecentEvent($"{_settingsWarning} 原因：{error}");
        }
        else _settingsWarning = null;
    }

    private void UiTimer_Tick(object? sender, EventArgs e)
    {
        long now = _clock.Timestamp;
        ScheduleAction action = _schedule.Tick(now);
        QueueSave(action);
        UpdateCountdown(now);
        if (IsVisible && now >= _nextInspectionAt)
            _ = RefreshInventoryAsync();
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e) => _ = RefreshInventoryAsync();

    private async Task RefreshInventoryAsync()
    {
#if DEBUG
        if (_isUiFixture) return;
#endif
        if (_isExiting || _inspectionInProgress || _schedule.IsSaving || _schedule.IsPaused) return;
        _inspectionInProgress = true;
        using CancellationTokenSource inspection = CancellationTokenSource.CreateLinkedTokenSource(_lifetimeCancellation.Token);
        _inspectionCancellation = inspection;
        RefreshButton.IsEnabled = false;
        InventoryTimeText.Text = "正在只读检测…";
        try
        {
            WordInventorySnapshot snapshot = await _wordSaveService.InspectAllAsync(inspection.Token);
            // A save can be queued while inspection is running. Its newer result wins.
            if (!_isExiting && !_schedule.IsSaving && !snapshot.Cancelled) InventorySnapshot = snapshot;
        }
        catch (OperationCanceledException) { }
        catch (Exception exception)
        {
            if (!_isExiting)
                InventorySnapshot = new(DateTimeOffset.Now, 0, 0,
                    [new("(检测异常)", null, "检测失败", exception.Message, false, null, DocumentOutcomeSource.Assistant)]);
        }
        finally
        {
            _inspectionInProgress = false;
            _inspectionCancellation = null;
            _nextInspectionAt = _clock.Add(_clock.Timestamp, TimeSpan.FromSeconds(10));
            UpdateUi();
        }
    }

    private void QueueSave(ScheduleAction action)
    {
        if (!action.QueueSave)
        {
            return;
        }

        _inspectionCancellation?.Cancel();
        CancellationToken token = _runCancellation?.Token ?? _lifetimeCancellation.Token;
        _ = RunSaveAsync(action.Generation, token);
    }

    private async Task RunSaveAsync(long generation, CancellationToken cancellationToken)
    {
        using CancellationTokenSource attempt = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetimeCancellation.Token);
        _attemptCancellation = attempt;
        _saveProgress = null;
        Progress<string> progress = new(message =>
        {
            if (!cancellationToken.IsCancellationRequested && generation == _schedule.Generation && !_isExiting && _schedule.IsSaving && !_schedule.IsPaused)
            {
                _saveProgress = message;
                UpdateUi();
            }
        });
        UpdateUi();
        SaveRoundResult? result = null;
        try
        {
            WordScanResult scan = await _wordSaveService.SaveAllAsync(attempt.Token, progress);
            result = attempt.IsCancellationRequested ? scan.Round with { Cancelled = true } : scan.Round;
            if (!attempt.IsCancellationRequested && !scan.Inventory.Cancelled && generation == _schedule.Generation && !_isExiting)
                InventorySnapshot = scan.Inventory;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            result = new SaveRoundResult(
                DateTimeOffset.Now,
                0,
                0,
                [
                    new DocumentSaveOutcome(
                        "(助手)",
                        DocumentSaveStatus.Failed,
                        $"保存轮次异常：{exception.Message}",
                        $"0x{exception.HResult:X8}",
                        Source: DocumentOutcomeSource.Assistant)
                ]);
        }

        _attemptCancellation = null;
        _saveProgress = null;
        ScheduleAction followUp = _schedule.CompleteSave(generation, _clock.Timestamp);
        if (result is not null && _roundDisplayState.TryApply(result, generation, _schedule.Generation))
        {
            ApplyResult(result, _roundDisplayState.Snapshot);
        }

        UpdateUi();
        QueueSave(followUp);
    }

    private void ApplyResult(
        SaveRoundResult result,
        RoundDisplaySnapshot? snapshot = null,
        bool notifyFailure = true)
    {
        RoundSnapshot = snapshot ?? ResultPresentation.CreateSnapshot(result);
        LastSaveText.Text = result.CompletedAt.LocalDateTime.ToString("HH:mm:ss");
        _eventHistory.AddRound(result);
        RefreshRecentEvents();

        if (notifyFailure && result.FailedCount > 0)
        {
            _trayIcon.ShowBalloonTip(
                4000,
                "Word 定时保存助手",
                $"本轮有 {result.FailedCount} 个保存失败，请打开主窗口查看。",
                Forms.ToolTipIcon.Warning);
        }
    }

    private void AddRecentEvent(string message)
    {
        _eventHistory.AddImmediate(message);
        RefreshRecentEvents();
    }

    private void RefreshRecentEvents()
    {
        _recentEvents.Clear();
        foreach (string item in _eventHistory.Items)
        {
            _recentEvents.Add(item);
        }
    }

    private void SystemEvents_SessionLockedChanged(bool locked)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (_isExiting)
            {
                return;
            }

            ScheduleAction action = _schedule.SetSessionLocked(locked, _clock.Timestamp);
            if (locked) { _attemptCancellation?.Cancel(); _inspectionCancellation?.Cancel(); }
            UpdateUi();
            QueueSave(action);
        });
    }

    private void SystemEvents_SystemSuspendedChanged(bool suspended)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (_isExiting)
            {
                return;
            }

            ScheduleAction action = _schedule.SetSystemSuspended(suspended, _clock.Timestamp);
            if (suspended) { _attemptCancellation?.Cancel(); _inspectionCancellation?.Cancel(); }
            UpdateUi();
            QueueSave(action);
        });
    }

    private void UpdateUi()
    {
        // State events refresh the full visible UI. Timer ticks only refresh changed seconds.
        // Hidden state is caught up immediately by RestoreWindow, without polling controls.
        if (!IsVisible) return;
        SettingsWarningText.Text = _settingsWarning ?? string.Empty;
        SettingsWarningText.Visibility = _settingsWarning is null ? Visibility.Collapsed : Visibility.Visible;
        SaveOnceButton.IsEnabled = !_schedule.IsSaving && !_schedule.IsPaused && !_isExiting;
        StartStopButton.IsEnabled = !_isExiting && (_schedule.IsRunning || _schedule.IsSaving || TryReadInterval(out _));
        RefreshButton.IsEnabled = !_inspectionInProgress && !_schedule.IsSaving && !_schedule.IsPaused && !_isExiting;
        InventorySummaryText.Text = InventorySnapshot?.Summary ?? "尚未完成检测";
        InventoryTimeText.Text = _inspectionInProgress ? "正在只读检测…"
            : InventorySnapshot is null ? "打开后自动检测，不会保存文档"
            : $"更新于 {InventorySnapshot.CheckedAt.LocalDateTime:HH:mm:ss} · 可见时每 10 秒刷新";
        InventoryEmptyText.Visibility = InventorySnapshot is null || InventorySnapshot.Items.Count == 0
            ? Visibility.Visible : Visibility.Collapsed;
        InventoryEmptyText.Text = InventorySnapshot is null ? "正在检测已打开的 Word 文档…" : "未检测到文档。请确认已打开 Microsoft Word 桌面版。";
        ConnectionWarningText.Visibility = InventorySnapshot?.UnavailableCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        ConnectionWarningText.Text = InventorySnapshot?.ConnectionWarning ?? string.Empty;
        DetectionText.Text = _schedule.IsSaving
            ? $"{RoundSnapshot.DetectionSummary}；正在更新"
            : RoundSnapshot.DetectionSummary;

        if (_schedule.IsRunning)
        {
            StartStopButton.Content = "停止";
            StartStopButton.Background = RedBrush;

            if (_schedule.IsPaused)
            {
                StatusText.Text = "已暂停";
                StatusDot.Fill = AmberBrush;
                CountdownText.Text = "--";
                FooterText.Text = "锁屏或休眠期间不执行；解锁或唤醒后立即检查一次。";
            }
            else if (_schedule.IsSaving)
            {
                StatusText.Text = "正在保存";
                StatusDot.Fill = BlueBrush;
                CountdownText.Text = "保存中";
                FooterText.Text = _saveProgress ?? "正在连接 Word。停止操作不会中断已进入 Word 的单次保存。";
            }
            else
            {
                StatusText.Text = "运行中";
                StatusDot.Fill = GreenBrush;
                FooterText.Text = "成功时静默运行；只有保存失败才显示托盘通知。";
            }
        }
        else
        {
            StartStopButton.Content = _schedule.IsSaving ? "停止本次保存" : "开始";
            StartStopButton.Background = BlueBrush;
            StatusText.Text = _schedule.IsSaving ? "单次保存中" : "已停止";
            StatusDot.Fill = GrayBrush;
            CountdownText.Text = "--";
            FooterText.Text = _schedule.IsSaving ? _saveProgress ?? "正在单次保存，定时任务仍为停止；可停止以取消等待和后续文档。"
                : "“立即保存一次”不启动定时；点击“开始”会立即保存并启动定时。最小化后继续运行。";
        }

        UpdateCountdown(_clock.Timestamp, force: true);

#if DEBUG
        if (_isUiFixture)
        {
            StartStopButton.IsEnabled = false;
            SaveOnceButton.IsEnabled = false;
            IntervalBox.IsEnabled = false;
            RefreshButton.IsEnabled = false;
            StatusText.Text = "界面验收";
            FooterText.Text = "调试假数据模式：不会连接 Word，也不能启动保存。";
        }
#endif
    }

    private void UpdateCountdown(long now, bool force = false)
    {
        if (_countdownRefresh.TryGetChange(IsVisible, _schedule.Remaining(now), _schedule.IsSaving,
            out int totalSeconds, force))
            CountdownText.Text = $"{totalSeconds / 60:00}:{totalSeconds % 60:00}";
    }

    private static SolidColorBrush CreateFrozenBrush(byte red, byte green, byte blue)
    {
        SolidColorBrush brush = new(MediaColor.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }

    private void ResultList_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not System.Windows.Controls.ListBox listBox)
        {
            return;
        }

        ScrollViewer? inner = FindVisualChild<ScrollViewer>(listBox);
        if (inner is null)
        {
            return;
        }

        bool atTop = e.Delta > 0 && inner.VerticalOffset <= 0;
        bool atBottom = e.Delta < 0 && inner.VerticalOffset >= inner.ScrollableHeight;
        if (!atTop && !atBottom)
        {
            return;
        }

        MainScrollViewer.ScrollToVerticalOffset(MainScrollViewer.VerticalOffset - e.Delta);
        e.Handled = true;
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, index);
            if (child is T match)
            {
                return match;
            }

            T? descendant = FindVisualChild<T>(child);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }

    private void Window_StateChanged(object? sender, EventArgs e)
    {
        if (WindowState != WindowState.Minimized)
        {
            return;
        }

        Hide();
        ShowInTaskbar = false;
        _trayIcon.ShowBalloonTip(
            1800,
            "Word 定时保存助手",
            _schedule.IsRunning ? "已最小化，定时保存仍在运行。" : "已最小化到系统托盘。",
            Forms.ToolTipIcon.Info);
    }

    private void RestoreWindow()
    {
        Show();
        ShowInTaskbar = true;
        WindowState = WindowState.Normal;
        Activate();
        UpdateUi();
        _ = RefreshInventoryAsync();
    }

    private void TrayIcon_DoubleClick(object? sender, EventArgs e) =>
        Dispatcher.BeginInvoke(RestoreWindow);

    private void ShowMenuItem_Click(object? sender, EventArgs e) =>
        Dispatcher.BeginInvoke(RestoreWindow);

    private void ExitMenuItem_Click(object? sender, EventArgs e) =>
        Dispatcher.BeginInvoke(() => _ = BeginExitAsync());

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowClose)
        {
            return;
        }

        e.Cancel = true;
        Dispatcher.BeginInvoke(() => _ = BeginExitAsync());
    }

    private async Task BeginExitAsync()
    {
        if (_isExiting)
        {
            return;
        }

        if (_schedule.IsRunning || _schedule.IsSaving)
        {
            MessageBoxResult confirm = System.Windows.MessageBox.Show(
                "保存任务正在运行。确定要停止任务并退出吗？",
                "退出确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }
        }

        _isExiting = true;
        _lifetimeCancellation.Cancel();
        StopSchedule();
        _uiTimer.Stop();
        _uiTimer.Tick -= UiTimer_Tick;
        _systemEvents.SessionLockedChanged -= SystemEvents_SessionLockedChanged;
        _systemEvents.SystemSuspendedChanged -= SystemEvents_SystemSuspendedChanged;
        _systemEvents.Dispose();

        bool stopped = await _wordSaveService.TryShutdownAsync(TimeSpan.FromSeconds(3));
        while (!stopped)
        {
            MessageBoxResult waitAgain = System.Windows.MessageBox.Show(
                "Word 调用尚未返回。\n\n选择“是”继续等待 3 秒；选择“否”将强制退出助手。强制退出不会关闭 Word，但当前保存结果可能未知。",
                "Word 仍在响应中",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.Yes);
            if (waitAgain == MessageBoxResult.No)
            {
                DisposeTray();
                Environment.Exit(0);
                return;
            }

            stopped = await _wordSaveService.TryShutdownAsync(TimeSpan.FromSeconds(3));
        }

        DisposeTray();
        _runCancellation?.Dispose();
        _lifetimeCancellation.Dispose();
        _allowClose = true;
        Close();
        ((App)System.Windows.Application.Current).CompleteShutdown();
    }

    private void DisposeTray()
    {
        _trayIcon.Visible = false;
        _trayIcon.DoubleClick -= TrayIcon_DoubleClick;
        _showMenuItem.Click -= ShowMenuItem_Click;
        _exitMenuItem.Click -= ExitMenuItem_Click;
        _trayMenu.Items.Clear();
        _trayIcon.Dispose();
        _applicationIcon.Dispose();
        _trayMenu.Dispose();
    }
}
