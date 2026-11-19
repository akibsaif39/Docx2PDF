using System.Threading;
using System.Windows;
using WpfApplication = System.Windows.Application;

namespace Docx2PDFLive.UI;

/// <summary>
/// Interaction logic for App.xaml
/// Enforces single-instance execution and wakes up the primary instance when clone is clicked.
/// </summary>
public partial class App : WpfApplication
{
    private static Mutex? _singleInstanceMutex;
    private static EventWaitHandle? _bringToFrontEvent;
    private static RegisteredWaitHandle? _registeredWait;

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
            try
            {
                using var wakeEvent = EventWaitHandle.OpenExisting(EventName);
                wakeEvent.Set();
            }
            catch { }

            Shutdown();
            return;
        }

        try
        {
            _bringToFrontEvent = new EventWaitHandle(false, EventResetMode.AutoReset, EventName);

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

        base.OnStartup(e);

        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            _registeredWait?.Unregister(null);
            _bringToFrontEvent?.Dispose();
            if (_singleInstanceMutex != null)
            {
                _singleInstanceMutex.ReleaseMutex();
                _singleInstanceMutex.Dispose();
            }
        }
        catch { }

        base.OnExit(e);
    }
}
