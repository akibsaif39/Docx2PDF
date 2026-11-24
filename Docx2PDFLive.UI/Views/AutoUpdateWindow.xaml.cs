using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;

namespace Docx2PDFLive.UI.Views;

public partial class AutoUpdateWindow : Window
{
    private readonly UpdateInfo _updateInfo;
    private readonly UpdateService _updateService;
    private readonly CancellationTokenSource _cts = new();
    private bool _isDownloading = false;

    public AutoUpdateWindow(UpdateInfo updateInfo, UpdateService? updateService = null)
    {
        InitializeComponent();
        _updateInfo = updateInfo;
        _updateService = updateService ?? new UpdateService();

        VersionSubtitle.Text = $"Updating from v{_updateInfo.CurrentVersion} to v{_updateInfo.LatestVersion}";
        if (!string.IsNullOrWhiteSpace(_updateInfo.ReleaseNotes))
        {
            ReleaseNotesText.Text = _updateInfo.ReleaseNotes;
        }

        Loaded += AutoUpdateWindow_Loaded;
        Closing += AutoUpdateWindow_Closing;
    }

    private async void AutoUpdateWindow_Loaded(object sender, RoutedEventArgs e)
    {
        // Automatically start the download on open for smooth zero-touch experience
        await StartUpdateAsync();
    }

    private void AutoUpdateWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_isDownloading)
        {
            _cts.Cancel();
        }
    }

    private async void OnUpdateClicked(object sender, RoutedEventArgs e)
    {
        await StartUpdateAsync();
    }

    private void OnExitClicked(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        DialogResult = false;
        Close();
    }

    private async Task StartUpdateAsync()
    {
        if (_isDownloading) return;
        _isDownloading = true;

        BtnUpdate.IsEnabled = false;
        BtnUpdate.Content = "Updating...";
        StatusLabel.Text = "Connecting to download server...";

        try
        {
            var tempInstaller = Path.Combine(Path.GetTempPath(), $"Docx2PDF_Install_v{_updateInfo.LatestVersion}.exe");
            
            var progress = new Progress<DownloadProgressReport>(report =>
            {
                Dispatcher.Invoke(() =>
                {
                    DownloadProgressBar.Value = report.ProgressPercentage;
                    
                    var mbReceived = report.BytesReceived / 1048576.0;
                    var mbTotal = report.TotalBytes / 1048576.0;
                    var speedMb = report.SpeedBytesPerSecond / 1048576.0;

                    if (report.TotalBytes > 0)
                    {
                        ProgressDetailLabel.Text = $"{report.ProgressPercentage:0}% ({mbReceived:0.0} / {mbTotal:0.0} MB)";
                    }
                    else
                    {
                        ProgressDetailLabel.Text = $"{mbReceived:0.0} MB";
                    }

                    if (speedMb > 0.05)
                    {
                        StatusLabel.Text = $"Downloading update ({speedMb:0.0} MB/s)...";
                    }
                    else
                    {
                        StatusLabel.Text = "Downloading update package...";
                    }
                });
            });

            await _updateService.DownloadUpdateAsync(_updateInfo.DownloadUrl, tempInstaller, progress, _cts.Token);

            StatusLabel.Text = "Update downloaded. Restarting Docx2PDF...";
            ProgressDetailLabel.Text = "100%";
            DownloadProgressBar.Value = 100;

            await Task.Delay(500);

            // Execute the downloaded installer with --silent-update
            var psi = new ProcessStartInfo
            {
                FileName = tempInstaller,
                Arguments = "--silent-update",
                UseShellExecute = true,
                Verb = "runas"
            };

            Process.Start(psi);

            _isDownloading = false;
            DialogResult = true;
            Close();
        }
        catch (OperationCanceledException)
        {
            StatusLabel.Text = "Update canceled.";
            _isDownloading = false;
            BtnUpdate.IsEnabled = true;
            BtnUpdate.Content = "Retry Update";
        }
        catch (Exception ex)
        {
            StatusLabel.Text = $"Update error: {ex.Message}";
            _isDownloading = false;
            BtnUpdate.IsEnabled = true;
            BtnUpdate.Content = "Retry Update";
        }
    }
}
