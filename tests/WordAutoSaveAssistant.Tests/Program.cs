using WordAutoSaveAssistant.Core;
using WordAutoSaveAssistant.Models;
using WordAutoSaveAssistant.Services;
using WordAutoSaveAssistant.WordInterop;
using System.Runtime.InteropServices;

namespace WordAutoSaveAssistant.Tests;

internal static class Program
{
    private static int _passed;

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Contains("--ui-refresh-benchmark", StringComparer.OrdinalIgnoreCase))
        {
            return UiRefreshBenchmark.Run();
        }

        if (args.Contains("--word-integration", StringComparer.OrdinalIgnoreCase))
        {
            return WordIntegration.Run(args.Contains("--timed", StringComparer.OrdinalIgnoreCase));
        }

        if (args.Contains("--word-discovery", StringComparer.OrdinalIgnoreCase))
        {
            return WordDiscoveryDiagnostic.Run(args);
        }

        string root = Path.Combine(Path.GetTempPath(), $"word-save-assistant-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            Run("署名：作者 LGH、赞助商 ZHZ", () =>
            {
                var metadata = typeof(SettingsStore).Assembly.GetCustomAttributes(
                    typeof(System.Reflection.AssemblyMetadataAttribute), false)
                    .Cast<System.Reflection.AssemblyMetadataAttribute>().ToDictionary(a => a.Key, a => a.Value);
                Equal("LGH", metadata["Author"]!);
                Equal("ZHZ", metadata["Sponsor"]!);
            });
            Run("图标：托盘资源可独立加载", () =>
            {
                using var icon = ApplicationBranding.LoadTrayIcon();
                Equal(32, icon.Width);
                Equal(32, icon.Height);
            });
            Run("图标：包含八种尺寸", () =>
            {
                using var stream = typeof(SettingsStore).Assembly.GetManifestResourceStream("WordAutoSaveAssistant.Assets.AppIcon.ico")!;
                using var reader = new BinaryReader(stream);
                Equal((ushort)0, reader.ReadUInt16());
                Equal((ushort)1, reader.ReadUInt16());
                Equal((ushort)8, reader.ReadUInt16());
                foreach (int size in new[] {16,20,24,32,48,64,128,256})
                {
                    Equal((byte)(size % 256), reader.ReadByte());
                    Equal((byte)(size % 256), reader.ReadByte());
                    stream.Seek(14, SeekOrigin.Current);
                }
            });
            Run("设置：首次启动为 10 分钟", () =>
            {
                SettingsStore store = new(Path.Combine(root, "default"));
                Equal(10, store.LoadInterval());
            });

            Run("设置：原子写入并重新读取", () =>
            {
                string directory = Path.Combine(root, "persist");
                SettingsStore store = new(directory);
                True(store.TrySaveInterval(25, out string? error), error);
                Equal(25, new SettingsStore(directory).LoadInterval());
                True(store.TrySaveInterval(30, out error), error);
                Equal(30, new SettingsStore(directory).LoadInterval());
                Equal(0, Directory.EnumerateFiles(directory, "settings.json.*.tmp").Count());
            });

            Run("设置：启动读取时只清理本应用残留临时文件", () =>
            {
                string directory = Path.Combine(root, "cleanup");
                Directory.CreateDirectory(directory);
                string leftover = Path.Combine(directory, "settings.json.123.deadbeef.tmp");
                File.WriteAllText(leftover, "partial");
                File.WriteAllText(Path.Combine(directory, "unrelated.tmp"), "keep");
                Equal(10, new SettingsStore(directory).LoadInterval());
                False(File.Exists(leftover));
                True(File.Exists(Path.Combine(directory, "unrelated.tmp")));
            });

            Run("设置：损坏或越界配置回退默认值", () =>
            {
                string directory = Path.Combine(root, "fallback");
                Directory.CreateDirectory(directory);
                File.WriteAllText(Path.Combine(directory, "settings.json"), "{broken");
                Equal(10, new SettingsStore(directory).LoadInterval());
                File.WriteAllText(Path.Combine(directory, "settings.json"), "{\"IntervalMinutes\":1441}");
                Equal(10, new SettingsStore(directory).LoadInterval());
            });

            Run("设置：拒绝 1–1440 以外的间隔", () =>
            {
                SettingsStore store = new(Path.Combine(root, "bounds"));
                False(store.TrySaveInterval(0, out _));
                False(store.TrySaveInterval(1441, out _));
            });

            Run("设置：目录无法创建时返回失败，不影响本次调度", () =>
            {
                string blocked = Path.Combine(root, "blocked-directory");
                File.WriteAllText(blocked, "keep");
                SettingsStore store = new(blocked);
                False(store.TrySaveInterval(23, out string? error));
                True(!string.IsNullOrWhiteSpace(error));
                Equal("keep", File.ReadAllText(blocked));
                Equal(10, store.LoadInterval());
                FakeClock clock = new();
                SaveScheduleState schedule = new(clock, 23);
                True(schedule.Start(clock.Timestamp).QueueSave);
                Equal(23, schedule.IntervalMinutes);
            });

            Run("设置：目标不可替换时保留既有设置，并清理自己的临时文件", () =>
            {
                string directory = Path.Combine(root, "locked-settings");
                SettingsStore store = new(directory);
                True(store.TrySaveInterval(25, out _));
                using (FileStream locked = File.Open(Path.Combine(directory, "settings.json"), FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    False(store.TrySaveInterval(30, out string? error));
                    True(!string.IsNullOrWhiteSpace(error));
                }
                Equal(25, store.LoadInterval());
                Equal(0, Directory.EnumerateFiles(directory, "settings.json.*.tmp").Count());
            });

            Run("调度：开始立即保存，之后按间隔触发", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                ScheduleAction first = state.Start(clock.Timestamp);
                True(first.QueueSave);
                state.CompleteSave(first.Generation, clock.Timestamp);
                clock.Advance(TimeSpan.FromMinutes(9));
                False(state.Tick(clock.Timestamp).QueueSave);
                clock.Advance(TimeSpan.FromMinutes(1));
                True(state.Tick(clock.Timestamp).QueueSave);
            });

            Run("调度：修改间隔会从修改时刻重新倒计时", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                ScheduleAction first = state.Start(clock.Timestamp);
                state.CompleteSave(first.Generation, clock.Timestamp);
                clock.Advance(TimeSpan.FromMinutes(8));
                state.ChangeInterval(5, clock.Timestamp);
                clock.Advance(TimeSpan.FromMinutes(4));
                False(state.Tick(clock.Timestamp).QueueSave);
                clock.Advance(TimeSpan.FromMinutes(1));
                True(state.Tick(clock.Timestamp).QueueSave);
            });

            Run("调度：停止不保存，重新启动才立即保存", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                ScheduleAction first = state.Start(clock.Timestamp);
                state.CompleteSave(first.Generation, clock.Timestamp);
                state.Stop();
                clock.Advance(TimeSpan.FromHours(1));
                False(state.Tick(clock.Timestamp).QueueSave);
                True(state.Start(clock.Timestamp).QueueSave);
            });

            Run("调度：锁屏和睡眠暂停，全部恢复后仅保存一次", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                ScheduleAction first = state.Start(clock.Timestamp);
                state.CompleteSave(first.Generation, clock.Timestamp);
                False(state.SetSessionLocked(true, clock.Timestamp).QueueSave);
                False(state.SetSystemSuspended(true, clock.Timestamp).QueueSave);
                False(state.SetSessionLocked(false, clock.Timestamp).QueueSave);
                True(state.SetSystemSuspended(false, clock.Timestamp).QueueSave);
            });

            Run("调度：暂停时点击开始，恢复后保存一次", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                state.SetSessionLocked(true, clock.Timestamp);
                False(state.Start(clock.Timestamp).QueueSave);
                True(state.SetSessionLocked(false, clock.Timestamp).QueueSave);
            });

            Run("调度：旧一轮保存完成不会吞掉重新开始请求", () =>
            {
                FakeClock clock = new();
                SaveScheduleState state = new(clock, 10);
                ScheduleAction oldRound = state.Start(clock.Timestamp);
                state.Stop();
                False(state.Start(clock.Timestamp).QueueSave);
                ScheduleAction replacement = state.CompleteSave(oldRound.Generation, clock.Timestamp);
                True(replacement.QueueSave);
                True(replacement.Generation != oldRound.Generation);
            });

            Run("结果：初始状态与完成后空分类文案明确区分", () =>
            {
                RoundDisplaySnapshot initial = RoundDisplaySnapshot.Initial;
                False(initial.HasCompletedRound);
                Equal(0, initial.Saved.Count);
                Equal("尚未执行", initial.Saved.Items.Single().DisplayName);

                RoundDisplaySnapshot completed = ResultPresentation.CreateSnapshot(SaveRoundResult.Empty());
                True(completed.HasCompletedRound);
                Equal(0, completed.Saved.Count);
                Equal("无", completed.Saved.Items.Single().DisplayName);
            });

            Run("结果：四类投影、重名保留及路径原因提示", () =>
            {
                SaveRoundResult result = new(
                    DateTimeOffset.Now,
                    1,
                    5,
                    [
                        new("同名.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"C:\A\同名.docx"),
                        new("同名.docx", DocumentSaveStatus.Saved, "保存成功", FullPath: @"D:\B\同名.docx"),
                        new("无修改.docx", DocumentSaveStatus.Unchanged, "没有未保存的修改", FullPath: @"C:\无修改.docx"),
                        new("只读.docx", DocumentSaveStatus.Skipped, "文档为只读状态", FullPath: @"C:\只读.docx"),
                        new("失败.docx", DocumentSaveStatus.Failed, "Word 返回错误", "0x800A0001", @"C:\失败.docx")
                    ]);

                RoundDisplaySnapshot snapshot = ResultPresentation.CreateSnapshot(result);
                Equal(2, snapshot.Saved.Count);
                Equal(2, snapshot.Saved.Items.Count);
                Equal(1, snapshot.Unchanged.Count);
                Equal(1, snapshot.Skipped.Count);
                Equal(1, snapshot.Failed.Count);
                Contains(@"C:\只读.docx", snapshot.Skipped.Items.Single().HelpText);
                Contains("原因：文档为只读状态", snapshot.Skipped.Items.Single().HelpText);
                Contains("0x800A0001", snapshot.Failed.Items.Single().HelpText);
            });

            Run("结果：自然排序支持中文、前导零和超长数字", () =>
            {
                SaveRoundResult result = new(
                    DateTimeOffset.Now,
                    1,
                    7,
                    [
                        new("文档100000000000000000000000000000.docx", DocumentSaveStatus.Saved, "ok"),
                        new("文档10.docx", DocumentSaveStatus.Saved, "ok"),
                        new("文档0002.docx", DocumentSaveStatus.Saved, "ok"),
                        new("文档2.docx", DocumentSaveStatus.Saved, "ok"),
                        new("文档99999999999999999999999999999.docx", DocumentSaveStatus.Saved, "ok"),
                        new("ignored", DocumentSaveStatus.Saved, "ok", Source: DocumentOutcomeSource.WordInstance),
                        new("ignored", DocumentSaveStatus.Saved, "ok", Source: DocumentOutcomeSource.Assistant)
                    ]);

                string[] names = ResultPresentation.CreateSnapshot(result).Saved.Items
                    .Select(item => item.DisplayName)
                    .ToArray();
                SequenceEqual(
                    ["文档2.docx", "文档0002.docx", "文档10.docx", "文档99999999999999999999999999999.docx", "文档100000000000000000000000000000.docx", "ignored", "(助手)"],
                    names);
            });

            Run("结果：保留失败窗口名称，助手保持伪来源标签", () =>
            {
                SaveRoundResult result = new(
                    DateTimeOffset.Now,
                    0,
                    0,
                    [
                        new("不是标签", DocumentSaveStatus.Failed, "实例错误", Source: DocumentOutcomeSource.WordInstance),
                        new("也不是标签", DocumentSaveStatus.Failed, "助手错误", Source: DocumentOutcomeSource.Assistant)
                    ]);
                string[] names = ResultPresentation.CreateSnapshot(result).Failed.Items
                    .Select(item => item.DisplayName)
                    .ToArray();
                SequenceEqual(["不是标签", "(助手)"], names);
            });

            Run("结果：有效轮次仅整体替换一次，取消和过时代次零变更", () =>
            {
                RoundDisplayState state = new();
                SaveRoundResult valid = new(
                    DateTimeOffset.Now,
                    1,
                    1,
                    [new("A.docx", DocumentSaveStatus.Saved, "保存成功")]);
                True(state.TryApply(valid, 4, 4));
                RoundDisplaySnapshot accepted = state.Snapshot;
                Equal(1, state.ReplacementCount);

                False(state.TryApply(valid, 4, 5));
                True(ReferenceEquals(accepted, state.Snapshot));
                Equal(1, state.ReplacementCount);

                SaveRoundResult cancelled = valid with { Cancelled = true };
                False(state.TryApply(cancelled, 5, 5));
                True(ReferenceEquals(accepted, state.Snapshot));
                Equal(1, state.ReplacementCount);
            });

            Run("历史：跨轮累计、同轮顺序和最近五条上限", () =>
            {
                RecentEventHistory history = new();
                history.AddRound(new SaveRoundResult(
                    DateTimeOffset.Now.AddMinutes(-1),
                    1,
                    2,
                    [
                        new("旧1.docx", DocumentSaveStatus.Skipped, "旧原因1"),
                        new("旧2.docx", DocumentSaveStatus.Failed, "旧原因2")
                    ]));
                history.AddRound(new SaveRoundResult(
                    DateTimeOffset.Now,
                    1,
                    4,
                    [
                        new("新1.docx", DocumentSaveStatus.Skipped, "新原因1"),
                        new("新2.docx", DocumentSaveStatus.Failed, "新原因2"),
                        new("新3.docx", DocumentSaveStatus.Skipped, "新原因3"),
                        new("新4.docx", DocumentSaveStatus.Failed, "新原因4")
                    ]));

                Equal(5, history.Items.Count);
                Contains("新1.docx", history.Items[0]);
                Contains("新2.docx", history.Items[1]);
                Contains("新3.docx", history.Items[2]);
                Contains("新4.docx", history.Items[3]);
                Contains("旧1.docx", history.Items[4]);

                history.AddImmediate("助手即时错误");
                Equal(5, history.Items.Count);
                Equal("助手即时错误", history.Items[0]);
                Contains("新4.docx", history.Items[4]);
            });

            Run("引擎：路径、来源及 FullName 失败边界", () =>
            {
                FakeWordApplication application = new(
                    new("修改.docx", @"C:\测试", @"C:\测试\修改.docx", saved: false),
                    new("无修改.docx", @"C:\测试", @"C:\测试\无修改.docx", saved: true),
                    new("未命名文档1", string.Empty, "未命名文档1", saved: false),
                    new("只读.docx", @"C:\测试", @"C:\测试\只读.docx", saved: false, readOnly: true),
                    new("受保护.docx", @"C:\测试", @"C:\测试\受保护.docx", saved: false, protectionType: 2),
                    new("路径读取失败.docx", @"C:\测试", @"C:\测试\路径读取失败.docx", saved: true, throwOnFullName: true),
                    new("属性读取失败.docx", @"C:\测试", @"C:\测试\属性读取失败.docx", saved: false, throwOnReadOnly: true));

                SaveRoundResult result = new WordSaveEngine().ProcessKnownApplicationsForTest(
                    [application],
                    CancellationToken.None);

                Equal(7, result.DetectedDocuments);
                Equal(1, result.SavedCount);
                Equal(2, result.UnchangedCount);
                Equal(3, result.SkippedCount);
                Equal(1, result.FailedCount);
                True(result.Outcomes.All(outcome => outcome.Source == DocumentOutcomeSource.Document));
                Equal(@"C:\测试\修改.docx", result.Outcomes.Single(outcome => outcome.DocumentName == "修改.docx").FullPath);
                Equal(null, result.Outcomes.Single(outcome => outcome.DocumentName == "未命名文档1").FullPath);
                DocumentSaveOutcome pathFailure = result.Outcomes.Single(outcome => outcome.DocumentName == "路径读取失败.docx");
                Equal(DocumentSaveStatus.Unchanged, pathFailure.Status);
                Equal(null, pathFailure.FullPath);
                DocumentSaveOutcome propertyFailure = result.Outcomes.Single(outcome => outcome.DocumentName == "属性读取失败.docx");
                Equal(DocumentSaveStatus.Failed, propertyFailure.Status);
                Equal(@"C:\测试\属性读取失败.docx", propertyFailure.FullPath);
            });

            Run("发现：按窗口覆盖多个实例、去重并压制同进程偶发失败", () =>
            {
                FakeWordApplication first = new(
                    new FakeWordDocument("一.docx", @"C:\测试", @"C:\测试\一.docx", saved: true));
                FakeWordApplication second = new(
                    new FakeWordDocument("二.docx", @"D:\测试", @"D:\测试\二.docx", saved: true));
                FakeWordWindowNativeApi api = new(
                    [
                        (new WordWindowCandidate((nint)1, 100, "一.docx - Word"), () => first),
                        (new WordWindowCandidate((nint)2, 100, "同进程另一窗口 - Word"), () =>
                            throw new COMException("窗口暂时忙", unchecked((int)0x8001010A))),
                        (new WordWindowCandidate((nint)3, 200, "二.docx - Word"), () => second),
                        (new WordWindowCandidate((nint)4, 300, "受限.docx - Word"), () =>
                            throw new COMException("拒绝访问", unchecked((int)0x80070005))),
                        (new WordWindowCandidate((nint)5, 400, "未公开.docx - Word"), () =>
                            throw new COMException("一般失败", unchecked((int)0x80004005)))
                    ]);

                WordDiscoveryResult discovery = RunningObjectTableWordDiscovery.DiscoverApplicationsForTest(api);
                Equal(2, discovery.Applications.Count);
                Equal(2, discovery.Failures.Count);
                WordDiscoveryFailure accessDenied = discovery.Failures.Single(failure => failure.ProcessId == 300);
                Contains("权限级别可能不同", accessDenied.Message);
                Equal("0x80070005", accessDenied.ErrorCode!);
                WordDiscoveryFailure unavailable = discovery.Failures.Single(failure => failure.ProcessId == 400);
                Contains("未公开桌面对象模型", unavailable.Message);
                Equal("0x80004005", unavailable.ErrorCode!);

                SaveRoundResult result = new WordSaveEngine().ProcessDiscoveryForTest(
                    discovery,
                    CancellationToken.None);
                Equal(2, result.DetectedInstances);
                Equal(2, result.DetectedDocuments);
                Equal(2, result.UnchangedCount);
                Equal(2, result.FailedCount);
                True(result.Outcomes.Where(outcome => outcome.Status == DocumentSaveStatus.Failed)
                    .All(outcome => outcome.Source == DocumentOutcomeSource.WordInstance));
            });

            Run("检测：跨实例、非活动文档、重名路径及跳过状态，绝不保存", () =>
            {
                FakeWordDocument changed = new("相同.docx", @"C:\A", @"C:\A\相同.docx", false);
                FakeWordDocument other = new("相同.docx", @"D:\B", @"D:\B\相同.docx", false);
                FakeWordDocument unnamed = new("文档1", "", "文档1", false);
                FakeWordDocument readOnly = new("只读.docx", @"C:\A", @"C:\A\只读.docx", false, readOnly: true);
                FakeWordApplication first = new(changed, unnamed, readOnly);
                FakeWordApplication second = new(other);
                WordInventorySnapshot result = new WordSaveEngine().InspectKnownApplicationsForTest([first, second], CancellationToken.None);
                Equal(2, result.DetectedInstances);
                Equal(4, result.DetectedDocuments);
                Equal(2, result.Items.Count(item => item.State == "待保存"));
                Equal(2, result.Items.Count(item => !item.CanAutoSave));
                Equal(2, result.SkippedCount);
                Equal(2, result.EligibleCount);
                Equal(2, result.PendingSaveCount);
                Equal(0, result.DocumentErrorCount);
                Equal(0, changed.SaveCalls + other.SaveCalls + unnamed.SaveCalls + readOnly.SaveCalls);
                False(changed.Saved);
                False(other.Saved);
            });

            Run("身份：枚举结束前保留 IUnknown，重复获取仅释放额外引用", () =>
            {
                int retained = 0;
                object first = new();
                object second = new();
                // Simulate a native allocator that reuses address 1 when no refs exist.
                ComIdentityScope scope = new(value =>
                {
                    nint identity = ReferenceEquals(value, first) || retained == 0 ? 1 : 2;
                    retained++;
                    return identity;
                }, _ => retained--);
                True(scope.TryAdd(first));
                Equal(1, retained);
                True(scope.TryAdd(second));
                False(scope.TryAdd(first));
                Equal(2, retained);
                scope.Dispose();
                Equal(0, retained);
                scope.Dispose();
                Equal(0, retained);
            });

            Run("引擎：先固定文档集合，保存期间集合变化也不会漏掉后续文档", () =>
            {
                FakeWordDocument first = new("前台.docx", @"C:\A", @"C:\A\前台.docx", false);
                FakeWordDocument background = new("非活动.docx", @"C:\A", @"C:\A\非活动.docx", false);
                FakeWordApplication app = new(first, background);
                first.AfterSave = () => app.Documents.Remove(first);
                SaveRoundResult round = new WordSaveEngine().ProcessKnownApplicationsForTest([app], CancellationToken.None);
                Equal(2, round.DetectedDocuments);
                Equal(2, round.SavedCount);
                Equal(1, background.SaveCalls);
            });

            Run("引擎：重复应用引用只保存一次，同名不同实例仍分别保存", () =>
            {
                FakeWordDocument first = new("报告.docx", @"C:\A", @"C:\A\报告.docx", false);
                FakeWordDocument second = new("报告.docx", @"D:\B", @"D:\B\报告.docx", false);
                FakeWordApplication app = new(first);
                SaveRoundResult round = new WordSaveEngine().ProcessKnownApplicationsForTest([app, app, new FakeWordApplication(second)], CancellationToken.None);
                Equal(2, round.DetectedDocuments);
                Equal(2, round.SavedCount);
                Equal(1, first.SaveCalls);
                Equal(1, second.SaveCalls);
            });

            Run("引擎：保存失败和未确认不能伪报成功，后续文档继续保存", () =>
            {
                FakeWordDocument failed = new("失败.docx", @"C:\A", @"C:\A\失败.docx", false) { FailSave = true };
                FakeWordDocument pending = new("未确认.docx", @"C:\A", @"C:\A\未确认.docx", false) { ConfirmSave = false };
                FakeWordDocument valid = new("成功.docx", @"C:\A", @"C:\A\成功.docx", false);
                SaveRoundResult round = new WordSaveEngine().ProcessKnownApplicationsForTest([new FakeWordApplication(failed, pending, valid)], CancellationToken.None);
                Equal(2, round.FailedCount);
                Equal(1, round.SavedCount);
                Contains("模拟磁盘保存失败", round.Outcomes.Single(item => item.DocumentName == "失败.docx").Message);
                Equal(1, valid.SaveCalls);
            });

            Run("检测：取消不保存，空列表不伪造文档", () =>
            {
                WordSaveEngine engine = new();
                Equal(0, engine.InspectKnownApplicationsForTest([], CancellationToken.None).Items.Count);
                FakeWordDocument document = new("取消.docx", @"C:\A", @"C:\A\取消.docx", false);
                WordInventorySnapshot snapshot = engine.InspectKnownApplicationsForTest([new FakeWordApplication(document)], new CancellationToken(true));
                True(snapshot.Cancelled);
                Equal(0, document.SaveCalls);
            });

            Run("单次：停止状态保存后仍停止，不创建倒计时", () =>
            {
                FakeClock clock = new();
                SaveScheduleState schedule = new(clock, 10);
                ScheduleAction action = schedule.SaveOnce();
                True(action.QueueSave);
                False(schedule.IsRunning);
                Equal(null, schedule.Remaining(clock.Timestamp));
                False(schedule.SaveOnce().QueueSave);
                False(schedule.CompleteSave(action.Generation, clock.Timestamp).QueueSave);
                False(schedule.IsSaving);
                False(schedule.IsRunning);
                True(schedule.SaveOnce().QueueSave);
            });
            Run("单次：运行中不移动下次定时截止点", () =>
            {
                FakeClock clock = new();
                SaveScheduleState schedule = new(clock, 10);
                ScheduleAction start = schedule.Start(clock.Timestamp);
                schedule.CompleteSave(start.Generation, clock.Timestamp);
                clock.Advance(TimeSpan.FromMinutes(2));
                ScheduleAction manual = schedule.SaveOnce();
                True(manual.QueueSave);
                clock.Advance(TimeSpan.FromSeconds(5));
                schedule.CompleteSave(manual.Generation, clock.Timestamp);
                Equal(TimeSpan.FromMinutes(8) - TimeSpan.FromSeconds(5), schedule.Remaining(clock.Timestamp));
            });
            Run("单次：锁屏拒绝新轮次，停止使旧结果失效", () =>
            {
                FakeClock clock = new();
                SaveScheduleState schedule = new(clock, 10);
                schedule.SetSessionLocked(true, clock.Timestamp);
                False(schedule.SaveOnce().QueueSave);
                schedule.SetSessionLocked(false, clock.Timestamp);
                ScheduleAction manual = schedule.SaveOnce();
                schedule.Stop();
                False(new RoundDisplayState().TryApply(SaveRoundResult.Empty(), manual.Generation, schedule.Generation));
                False(schedule.CompleteSave(manual.Generation, clock.Timestamp).QueueSave);
            });
            Run("忙碌：仅延迟五秒一次，不重复保存已成功文档", () =>
            {
                FakeWordDocument busy = new("忙.docx", @"C:\A", @"C:\A\忙.docx", false) { BusySave = true };
                FakeWordDocument valid = new("成功.docx", @"D:\B", @"D:\B\成功.docx", false);
                int delays = 0;
                WordSaveEngine engine = new((duration, _) =>
                {
                    Equal(TimeSpan.FromSeconds(5), duration);
                    Equal(1, valid.CompletedSaves);
                    delays++;
                    busy.BusySave = false;
                });
                SaveRoundResult result = engine.ProcessKnownApplicationsForTest([new FakeWordApplication(busy, valid)], CancellationToken.None);
                Equal(1, delays);
                Equal(2, result.SavedCount);
                Equal(0, result.FailedCount);
                Equal(1, valid.SaveCalls);
                Equal(1, busy.CompletedSaves);
                Contains("延迟重试一次", result.Outcomes.Single(x => x.DocumentName == "忙.docx").Message);
            });
            Run("忙碌：Save 返回后仅重试确认，已保存分类不误记为无需保存", () =>
            {
                FakeWordDocument doc = new("确认.docx", @"C:\A", @"C:\A\确认.docx", false);
                doc.AfterSave = () => doc.BusyConfirmation = true;
                int delays = 0;
                SaveRoundResult result = new WordSaveEngine((_, _) => { delays++; doc.BusyConfirmation = false; })
                    .ProcessKnownApplicationsForTest([new FakeWordApplication(doc)], CancellationToken.None);
                Equal(1, delays);
                Equal(1, doc.SaveCalls);
                Equal(1, result.SavedCount);
                Equal(0, result.UnchangedCount);
                Contains("延迟读取后已确认", result.Outcomes.Single().Message);
            });
            Run("忙碌：延迟确认仍有修改时保留未确认失败，不重复调用 Save", () =>
            {
                FakeWordDocument doc = new("仍未确认.docx", @"C:\A", @"C:\A\仍未确认.docx", false) { ConfirmSave = false };
                doc.AfterSave = () => doc.BusyConfirmation = true;
                SaveRoundResult result = new WordSaveEngine((_, _) => doc.BusyConfirmation = false)
                    .ProcessKnownApplicationsForTest([new FakeWordApplication(doc)], CancellationToken.None);
                Equal(1, doc.SaveCalls);
                Equal(1, result.FailedCount);
                Equal(0, result.SavedCount);
                Contains("Word 仍报告未保存修改", result.Outcomes.Single().Message);
            });
            Run("忙碌：集合暂时拒绝调用也能延迟恢复", () =>
            {
                FakeWordDocument doc = new("恢复.docx", @"C:\A", @"C:\A\恢复.docx", false);
                FakeWordApplication app = new(doc) { BusyDocuments = true };
                int delays = 0;
                WordSaveEngine engine = new((_, _) => { delays++; app.BusyDocuments = false; });
                SaveRoundResult result = engine.ProcessKnownApplicationsForTest([app], CancellationToken.None);
                Equal(1, delays);
                Equal(1, result.DetectedInstances);
                Equal(1, result.DetectedDocuments);
                Equal(1, result.SavedCount);
            });
            Run("忙碌：仅重连忙碌进程，权限错误直接报告", () =>
            {
                FakeWordDocument doc = new("恢复.docx", @"C:\A", @"C:\A\恢复.docx", false);
                int delays = 0, connections = 0;
                WordSaveEngine engine = new((_, _) => delays++, (pid, _) =>
                {
                    Equal(123, pid);
                    connections++;
                    return new([new FakeWordApplication(doc)], []);
                });
                SaveRoundResult result = engine.ProcessDiscoveryForTest(new([], [
                    new(123, "忙.docx - Word", "暂时忙碌", "0x8001010A"),
                    new(456, "受限.docx - Word", "权限不同", "0x80070005")]), CancellationToken.None);
                Equal(1, delays);
                Equal(1, connections);
                Equal(1, result.SavedCount);
                Equal(1, result.FailedCount);
                Equal("受限.docx - Word", result.Outcomes.Single(x => x.Status == DocumentSaveStatus.Failed).DocumentName);
            });
            Run("忙碌：等待被取消后绝不进入延迟重试", () =>
            {
                FakeWordDocument doc = new("取消.docx", @"C:\A", @"C:\A\取消.docx", false) { BusySave = true };
                using CancellationTokenSource cts = new();
                WordSaveEngine engine = new((_, _) => cts.Cancel());
                SaveRoundResult result = engine.ProcessKnownApplicationsForTest([new FakeWordApplication(doc)], cts.Token);
                True(result.Cancelled);
                Equal(0, doc.CompletedSaves);
                True(doc.BusySave);
            });
            Run("忙碌：持续忙碌仅一轮延迟，保留失败而不无限循环", () =>
            {
                FakeWordDocument doc = new("仍忙.docx", @"C:\A", @"C:\A\仍忙.docx", false) { BusySave = true };
                int delays = 0;
                SaveRoundResult result = new WordSaveEngine((_, _) => delays++)
                    .ProcessKnownApplicationsForTest([new FakeWordApplication(doc)], CancellationToken.None);
                Equal(1, delays);
                Equal(1, result.FailedCount);
                Contains("延迟重试一次", result.Outcomes.Single().Message);
            });
            Run("忙碌：磁盘错误与未确认保存不延迟重试", () =>
            {
                FakeWordDocument failed = new("磁盘.docx", @"C:\A", @"C:\A\磁盘.docx", false) { FailSave = true };
                FakeWordDocument pending = new("未确认.docx", @"C:\A", @"C:\A\未确认.docx", false) { ConfirmSave = false };
                SaveRoundResult result = new WordSaveEngine((_, _) => throw new Exception("不得重试"))
                    .ProcessKnownApplicationsForTest([new FakeWordApplication(failed, pending)], CancellationToken.None);
                Equal(2, result.FailedCount);
                Equal(1, failed.SaveCalls);
                Equal(1, pending.SaveCalls);
            });
            Run("忙碌：只读检测绝不启动延迟保存", () =>
            {
                FakeWordApplication app = new() { BusyDocuments = true };
                WordInventorySnapshot result = new WordSaveEngine((_, _) => throw new Exception("不得重试"))
                    .InspectKnownApplicationsForTest([app], CancellationToken.None);
                Equal(1, result.UnavailableCount);
            });
            Run("忙碌：真实延迟等待可及时取消，不等满五秒", () =>
            {
                FakeWordDocument doc = new("等待.docx", @"C:\A", @"C:\A\等待.docx", false) { BusySave = true };
                using CancellationTokenSource cts = new(TimeSpan.FromSeconds(4));
                System.Diagnostics.Stopwatch elapsed = System.Diagnostics.Stopwatch.StartNew();
                SaveRoundResult result = new WordSaveEngine().ProcessKnownApplicationsForTest([new FakeWordApplication(doc)], cts.Token);
                True(result.Cancelled);
                Equal(0, doc.CompletedSaves);
                True(elapsed.Elapsed < TimeSpan.FromSeconds(6));
            });

            Run("刷新：停止和隐藏期间不刷新倒计时，热循环不分配内存", () =>
            {
                CountdownRefreshState refresh = new();
                for (int i = 0; i < 100; i++) refresh.TryGetChange(true, null, false, out _);
                long before = GC.GetAllocatedBytesForCurrentThread();
                int changes = 0;
                for (int i = 0; i < 120_000; i++)
                {
                    if (refresh.TryGetChange(true, null, false, out _)) changes++;
                    if (refresh.TryGetChange(false, TimeSpan.FromSeconds(60), false, out _)) changes++;
                }
                long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
                Equal(0, changes);
                Equal(0L, allocated);
            });
            Run("刷新：同一显示秒只更新一次，跨秒立即更新", () =>
            {
                CountdownRefreshState refresh = new();
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(120), false, out int seconds));
                Equal(120, seconds);
                False(refresh.TryGetChange(true, TimeSpan.FromSeconds(119.75), false, out _));
                False(refresh.TryGetChange(true, TimeSpan.FromSeconds(119.5), false, out _));
                False(refresh.TryGetChange(true, TimeSpan.FromSeconds(119.25), false, out _));
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(119), false, out seconds));
                Equal(119, seconds);
            });
            Run("刷新：保存和暂停状态不覆盖提示，恢复后立即更新", () =>
            {
                CountdownRefreshState refresh = new();
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(120), false, out _));
                False(refresh.TryGetChange(true, TimeSpan.FromSeconds(120), true, out _));
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(120), false, out _));
                False(refresh.TryGetChange(true, null, false, out _));
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(120), false, out _));
            });
            Run("刷新：托盘恢复和事件强制刷新不被同秒缓存抑制", () =>
            {
                CountdownRefreshState refresh = new();
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(60), false, out _));
                False(refresh.TryGetChange(false, TimeSpan.FromSeconds(35), false, out _, force: true));
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(35), false, out int seconds, force: true));
                Equal(35, seconds);
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(35), false, out _, force: true));
                True(refresh.TryGetChange(true, TimeSpan.FromMinutes(25), false, out seconds));
                Equal(1500, seconds);
                True(refresh.TryGetChange(true, TimeSpan.FromSeconds(-1), false, out seconds));
                Equal(0, seconds);
            });
            Run("刷新：隐藏不改变调度，到期仍只排队一次", () =>
            {
                FakeClock clock = new();
                SaveScheduleState schedule = new(clock, 1);
                ScheduleAction first = schedule.Start(clock.Timestamp);
                schedule.CompleteSave(first.Generation, clock.Timestamp);
                CountdownRefreshState refresh = new();
                int queued = 0;
                for (int tick = 0; tick < 240; tick++)
                {
                    clock.Advance(TimeSpan.FromMilliseconds(250));
                    if (schedule.Tick(clock.Timestamp).QueueSave) queued++;
                    False(refresh.TryGetChange(false, schedule.Remaining(clock.Timestamp), schedule.IsSaving, out _));
                }
                Equal(1, queued);
                True(schedule.IsSaving);
                False(schedule.Tick(clock.Timestamp).QueueSave);
            });
            Run("覆盖：符合条件、修改、跳过和文档异常不相互混淆", () =>
            {
                WordInventorySnapshot snapshot = new(DateTimeOffset.Now, 1, 4, [
                    new("修改.docx", null, "待保存", "", true, true),
                    new("未改.docx", null, "无未保存修改", "", true, false),
                    new("只读.docx", null, "只读", "", false, null, IsSkipped: true),
                    new("异常.docx", null, "无法读取", "", false, null),
                    new("受限 - Word", null, "无法连接", "", false, null, DocumentOutcomeSource.WordInstance,
                        ConnectionIssue: WordConnectionIssue.PermissionMismatch)]);
                Equal(2, snapshot.EligibleCount);
                Equal(1, snapshot.PendingSaveCount);
                Equal(1, snapshot.SkippedCount);
                Equal(1, snapshot.DocumentErrorCount);
                Equal(1, snapshot.UnavailableCount);
                Contains("符合保存条件 2（有修改 1）", snapshot.Summary);
                Contains("权限差异已确认 1", snapshot.ConnectionWarning);
            });
            Run("覆盖：对象模型或未知故障不统一建议权限对齐", () =>
            {
                WordInventorySnapshot snapshot = new(DateTimeOffset.Now, 0, 0, [
                    new("Word", null, "无法连接", "", false, null, DocumentOutcomeSource.WordInstance,
                        ConnectionIssue: WordConnectionIssue.ObjectModelUnavailable),
                    new("未知", null, "无法连接", "", false, null, DocumentOutcomeSource.WordInstance)]);
                Contains("对象模型不可用 1", snapshot.ConnectionWarning);
                Contains("其他原因未确认 1", snapshot.ConnectionWarning);
                False(snapshot.ConnectionWarning.Contains("需手动对齐权限", StringComparison.Ordinal));
                Equal(string.Empty, new WordInventorySnapshot(DateTimeOffset.Now, 0, 0, []).ConnectionWarning);
                WordInventorySnapshot failedInspection = new(DateTimeOffset.Now, 0, 0,
                    [new("(检测异常)", null, "检测失败", "模拟故障", false, null, DocumentOutcomeSource.Assistant)]);
                Equal(1, failedInspection.DetectionErrorCount);
                Contains("清单可能不完整", failedInspection.Summary);
            });
            Run("覆盖：忙碌和拒绝访问提示不冒充已确认权限差异", () =>
            {
                WordInventorySnapshot snapshot = new(DateTimeOffset.Now, 0, 0, [
                    new("忙", null, "无法连接", "", false, null, DocumentOutcomeSource.WordInstance,
                        ConnectionIssue: WordConnectionIssue.Busy),
                    new("拒绝", null, "无法连接", "", false, null, DocumentOutcomeSource.WordInstance,
                        ConnectionIssue: WordConnectionIssue.AccessDenied)]);
                Contains("Word 暂时忙碌 1", snapshot.ConnectionWarning);
                Contains("拒绝访问，权限差异未确认 1", snapshot.ConnectionWarning);
                False(snapshot.ConnectionWarning.Contains("权限差异已确认", StringComparison.Ordinal));
            });
            Run("诊断：权限差异必须有两端证据，相同权限不误归因", () =>
            {
                Equal(WordConnectionIssue.PermissionMismatch, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x80004005), 0x3000, 0x2000));
                Equal(WordConnectionIssue.ObjectModelUnavailable, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x80004005), 0x3000, 0x3000));
                Equal(WordConnectionIssue.AccessDenied, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x80070005), null, 0x2000));
                Equal(WordConnectionIssue.Unknown, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x80004004), null, null));
                Contains("权限级别一致", WordProcessSecurity.DescribeConnectionBoundary(0x3000, 0x3000));
                Contains("未能完整读取", WordProcessSecurity.DescribeConnectionBoundary(0x3000, null));
            });
            Run("诊断：忙碌错误明确归为暂时忙碌", () =>
            {
                Equal(WordConnectionIssue.Busy, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x8001010A), 0x3000, 0x2000));
                Equal(WordConnectionIssue.Busy, WordProcessSecurity.ClassifyConnectionIssue(unchecked((int)0x80010001), null, null));
                WordDiscoveryFailure failure = new(123, "忙 - Word", "忙碌", "0x8001010A", WordConnectionIssue.Busy);
                WordInventorySnapshot snapshot = new WordSaveEngine().InspectDiscoveryForTest(new([], [failure]), CancellationToken.None);
                Equal(WordConnectionIssue.Busy, snapshot.Items.Single().ConnectionIssue);
            });
            Run("结果：最近一轮明确标注保存时间，不覆盖实时检测含义", () =>
            {
                Contains("尚未执行保存轮次", RoundDisplaySnapshot.Initial.DetectionSummary);
                Contains("当前文档", RoundDisplaySnapshot.Initial.DetectionSummary);
                RoundDisplaySnapshot completed = ResultPresentation.CreateSnapshot(SaveRoundResult.Empty());
                Contains("最近一轮", completed.DetectionSummary);
                Contains(completed.CompletedAt!.Value.LocalDateTime.ToString("HH:mm:ss"), completed.DetectionSummary);
            });

            Console.WriteLine($"PASS: {_passed} 个测试全部通过");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"FAIL: {exception.Message}");
            return 1;
        }
        finally
        {
            try { Directory.Delete(root, true); } catch { }
        }
    }

    private static void Run(string name, Action test)
    {
        test();
        _passed++;
        Console.WriteLine($"PASS: {name}");
    }

    private static void True(bool value, string? message = null)
    {
        if (!value) throw new InvalidOperationException(message ?? "预期为 true，实际为 false。");
    }

    private static void False(bool value) => True(!value, "预期为 false，实际为 true。");

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException($"预期 {expected}，实际 {actual}。");
        }
    }

    private static void Contains(string expectedSubstring, string actual)
    {
        if (!actual.Contains(expectedSubstring, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"预期包含 {expectedSubstring}，实际为 {actual}。");
        }
    }

    private static void SequenceEqual<T>(IReadOnlyList<T> expected, IReadOnlyList<T> actual)
    {
        if (!expected.SequenceEqual(actual))
        {
            throw new InvalidOperationException(
                $"序列不一致。预期 [{string.Join(", ", expected)}]，实际 [{string.Join(", ", actual)}]。");
        }
    }

    private sealed class FakeClock : IMonotonicClock
    {
        public long Timestamp { get; private set; }
        public TimeSpan Elapsed(long start, long end) => TimeSpan.FromTicks(end - start);
        public long Add(long timestamp, TimeSpan duration) => timestamp + duration.Ticks;
        public void Advance(TimeSpan duration) => Timestamp += duration.Ticks;
    }

    private sealed class FakeWordWindowNativeApi : IWordWindowNativeApi
    {
        private readonly IReadOnlyList<(WordWindowCandidate Candidate, Func<object> Resolve)> _entries;

        public FakeWordWindowNativeApi(
            IReadOnlyList<(WordWindowCandidate Candidate, Func<object> Resolve)> entries)
        {
            _entries = entries;
        }

        public IReadOnlyList<WordWindowCandidate> EnumerateWordWindows() =>
            _entries.Select(entry => entry.Candidate).ToArray();

        public object GetNativeObject(WordWindowCandidate candidate) =>
            _entries.Single(entry => entry.Candidate.Handle == candidate.Handle).Resolve();
    }
}
