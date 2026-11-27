using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Docx2PDFLive.Core.Interop;
using Docx2PDFLive.Core.Services;
using Docx2PDFLive.UI.Views;
using WpfApplication = System.Windows.Application;

namespace Docx2PDFLive.UI;

/// <summary>
/// Interaction logic for App.xaml
/// Enforces single-instance execution, mandatory online auto-updates, and prerequisite detection.
/// </summary>
public partial class App : WpfApplication
{
    private static Mutex? _singleInstanceMutex;
    private static EventWaitHandle? _bringToFrontEvent;
    private static RegisteredWaitHandle? _registeredWait;
    private static System.Threading.Timer? _backgroundUpdateTimer;

    public const string MutexName = @"Local\Docx2PDF_SingleInstance_Mutex_9F842FDEE031";
    public const string EventName = @"Local\Docx2PDF_SingleInstance_Event_9F842FDEE031";

    protected override void OnStartup(StartupEventArgs e)
    {
        bool isNewInstance;
        try
        {
            _singleInstanceMutex = new Mutex(true, MutexName, out isNewInstance);
        }
        catch
        {
            isNewInstance = true;
        }

        if (!isNewInstance)
        {
            // Signal running primary instance to bring its window to the foreground
            try
            {
                using var wakeEvent = EventWaitHandle.OpenExisting(EventName);
                wakeEvent.Set();
            }
            catch { }

            // Immediately exit the duplicate clone instance
            Shutdown();
            return;
        }

        try
        {
            _bringToFrontEvent = new EventWaitHandle(false, EventResetMode.AutoReset, EventName);

            // Register background listener on threadpool to wake up main window
            _registeredWait = ThreadPool.RegisterWaitForSingleObject(
                _bringToFrontEvent,
                (state, timedOut) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (MainWindow is MainWindow win)
                        {
                            win.RestoreWindow();
                        }
                    });
                },
                null,
                Timeout.Infinite,
                false);
        }
        catch { }

        // Online mandatory auto-update check (Option 1: Graceful offline fallback)
        try
        {
            var updateService = new UpdateService();
            var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);
            var updateInfo = Task.Run(() => updateService.CheckForUpdateAsync(currentVersion, timeoutSeconds: 3)).GetAwaiter().GetResult();

            if (updateInfo.IsUpdateAvailable && updateInfo.IsMandatory)
            {
                var updateModal = new AutoUpdateWindow(updateInfo, updateService);
                updateModal.ShowDialog();

                // Shutdown instance so installer can apply files in-place cleanly
                ReleaseSingleInstanceResources();
                Shutdown();
                return;
            }
        }
        catch { }

        try
        {
            // Initialize Windows Job Object with KILL_ON_JOB_CLOSE to guarantee zero ghost worker processes
            JobObjectManager.InitializeGlobalJobObject();
        }
        catch { }

        // Mandatory prerequisite check: User must have Microsoft Word or LibreOffice
        var engines = OfficeEngineDetector.GetAvailableEngines();
        if (engines == OfficeEngineType.None)
        {
            var prereqModal = new PrerequisiteCheckWindow();
            var result = prereqModal.ShowDialog();

            OfficeEngineDetector.ResetCache();
            if (result != true || OfficeEngineDetector.GetAvailableEngines() == OfficeEngineType.None)
            {
                // Engine requirement was not satisfied; close application without opening main window
                Shutdown();
                return;
            }
        }

        base.OnStartup(e);

        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();

        // Start non-intrusive periodic background update timer
        StartPeriodicUpdateChecks();
    }

    private void StartPeriodicUpdateChecks()
    {
        // First check after 30 minutes, then repeat every 60 minutes
        _backgroundUpdateTimer = new System.Threading.Timer(async _ =>
        {
            try
            {
                var updateService = new UpdateService();
                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);
                var updateInfo = await updateService.CheckForUpdateAsync(currentVersion, timeoutSeconds: 4).ConfigureAwait(false);

                if (updateInfo.IsUpdateAvailable && updateInfo.IsMandatory)
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (MainWindow is MainWindow win)
                        {
                            // If actively converting, wait until queue is idle before prompting
                            if (win.IsQueueBusy) return;

                            var updateModal = new AutoUpdateWindow(updateInfo, updateService);
                            if (win.IsVisible)
                            {
                                updateModal.Owner = win;
                            }
                            updateModal.ShowDialog();
                            win.PrepareForShutdown();
                            ReleaseSingleInstanceResources();
                            Shutdown();
                        }
                    });
                }
            }
            catch { }
        }, null, TimeSpan.FromMinutes(30), TimeSpan.FromHours(1));
    }

    private void ReleaseSingleInstanceResources()
    {
        try
        {
            _backgroundUpdateTimer?.Dispose();
            _backgroundUpdateTimer = null;
            _registeredWait?.Unregister(null);
            _registeredWait = null;
            _bringToFrontEvent?.Dispose();
            _bringToFrontEvent = null;
            if (_singleInstanceMutex != null)
            {
                _singleInstanceMutex.ReleaseMutex();
                _singleInstanceMutex.Dispose();
                _singleInstanceMutex = null;
            }
        }
        catch { }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        ReleaseSingleInstanceResources();
        base.OnExit(e);
    }
}
