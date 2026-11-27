using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Docx2PDFLive.Core.Services;
using Docx2PDFLive.UI.ViewModels;
using Docx2PDFLive.UI.Views;
using Application = System.Windows.Application;

namespace Docx2PDFLive.UI;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    private NotifyIcon? _notifyIcon;
    private bool _isExplicitExit;
    private bool _hasShownBalloon;

    public bool IsQueueBusy => _viewModel.IsQueueBusy;

    public MainWindow()
    {
        InitializeComponent();

        try
        {
            this.Icon = new System.Windows.Media.Imaging.BitmapImage(
                new Uri("pack://application:,,,/Resources/app_icon.png"));
        }
        catch { }

        _viewModel = (MainViewModel)DataContext;

        Loaded += OnWindowLoaded;
        Closing += OnWindowClosing;
        StateChanged += OnWindowStateChanged;

        InitializeTrayIcon();
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        // One-time first-launch welcome modal check
        if (!FirstRunMarker.HasBeenWelcomed())
        {
            var welcomeModal = new FirstRunWelcomeWindow
            {
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var result = welcomeModal.ShowDialog();

            // After Get Started is clicked, prompt for folder
            if (result == true)
            {
                _viewModel.PromptSelectFolder();
            }
        }
        else
        {
            // Auto-resume last remembered folder if enabled and still present; otherwise prompt
            _viewModel.InitializeFolderOnStartup();
        }
    }

    private void InitializeTrayIcon()
    {
        try
        {
            _notifyIcon = new NotifyIcon
            {
                Text = "Docx2PDF - Created by Akib",
                Visible = true
            };

            // Load icon from embedded resources or disk
            var icoStream = Application.GetResourceStream(new Uri("pack://application:,,,/Resources/app.ico"))?.Stream;
            if (icoStream != null)
            {
                _notifyIcon.Icon = new Icon(icoStream);
            }
            else
            {
                _notifyIcon.Icon = SystemIcons.Application;
            }

            var contextMenu = new ContextMenuStrip();

            var openItem = new ToolStripMenuItem("Open Docx2PDF", null, (s, e) => RestoreWindow());
            var aboutItem = new ToolStripMenuItem("About & Credits...", null, (s, e) => _viewModel.ShowAboutDialog());
            var pauseItem = new ToolStripMenuItem("Pause / Resume", null, (s, e) => _viewModel.PauseResumeCommand.Execute(null));
            var separator = new ToolStripSeparator();
            var exitItem = new ToolStripMenuItem("Exit", null, (s, e) => ExitApplication());

            contextMenu.Items.Add(openItem);
            contextMenu.Items.Add(aboutItem);
            contextMenu.Items.Add(pauseItem);
            contextMenu.Items.Add(separator);
            contextMenu.Items.Add(exitItem);

            _notifyIcon.ContextMenuStrip = contextMenu;
            _notifyIcon.DoubleClick += (s, e) => RestoreWindow();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to initialize tray icon: {ex.Message}");
        }
    }

    private void OnWindowStateChanged(object? sender, EventArgs e)
    {
        // When minimized, hide to system tray
        if (WindowState == WindowState.Minimized)
        {
            HideToTray();
        }
    }

    private void OnWindowClosing(object? sender, CancelEventArgs e)
    {
        if (_isExplicitExit)
        {
            // Clean exit
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }

            _viewModel.Dispose();
        }
        else
        {
            // When user clicks the close [X] button, hide to system tray instead of exiting
            e.Cancel = true;
            HideToTray();
        }
    }

    private void HideToTray()
    {
        Hide();
        ShowInTaskbar = false;

        if (!_hasShownBalloon && _notifyIcon != null)
        {
            _hasShownBalloon = true;
            _notifyIcon.ShowBalloonTip(
                2500,
                "Docx2PDF",
                "App is running in the background. Double-click this icon to reopen.",
                ToolTipIcon.Info);
        }
    }

    public void RestoreWindow()
    {
        if (!IsVisible)
        {
            Show();
        }
        ShowInTaskbar = true;
        if (WindowState == WindowState.Minimized)
        {
            WindowState = WindowState.Normal;
        }
        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }

    private void ExitApplication()
    {
        _isExplicitExit = true;
        Close();
    }
}
