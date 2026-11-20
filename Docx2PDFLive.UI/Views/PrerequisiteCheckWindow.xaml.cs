using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Windows;
using Docx2PDFLive.Core.Services;

namespace Docx2PDFLive.UI.Views;

public partial class PrerequisiteCheckWindow : Window
{
    private const string LibreOfficeDownloadUrl = "https://download.documentfoundation.org/libreoffice/stable/24.8.3/win/x86_64/LibreOffice_24.8.3_Win_x86-64.msi";
    private const string LibreOfficeWebPage = "https://www.libreoffice.org/download/download-libreoffice/";

    public PrerequisiteCheckWindow()
    {
        InitializeComponent();
    }

    private async void InstallButton_Click(object sender, RoutedEventArgs e)
    {
        InstallButton.IsEnabled = false;
        CheckAgainButton.IsEnabled = false;
        ExitButton.IsEnabled = false;

        DownloadProgressBar.Visibility = Visibility.Visible;
        ProgressDetailsText.Visibility = Visibility.Visible;
        StatusTitleText.Text = "Downloading LibreOffice...";
        StatusSubText.Text = "Downloading official installer from LibreOffice mirror. Please do not close this window.";
        ProgressDetailsText.Text = "Connecting to mirror...";

        var tempMsi = Path.Combine(Path.GetTempPath(), "LibreOffice_Setup.msi");

        try
        {
            await Task.Run(async () =>
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromMinutes(15);
                using var response = await client.GetAsync(LibreOfficeDownloadUrl, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                var totalBytes = response.Content.Headers.ContentLength ?? -1L;
                using var stream = await response.Content.ReadAsStreamAsync();
                using var fs = new FileStream(tempMsi, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);

                var buffer = new byte[81920];
                long totalRead = 0;
                int bytesRead;

                while ((bytesRead = await stream.ReadAsync(buffer)) > 0)
                {
                    await fs.WriteAsync(buffer.AsMemory(0, bytesRead));
                    totalRead += bytesRead;

                    if (totalBytes > 0)
                    {
                        var progressPercent = (int)((totalRead * 100) / totalBytes);
                        var mbRead = totalRead / (1024 * 1024);
                        var totalMb = totalBytes / (1024 * 1024);

                        Dispatcher.Invoke(() =>
                        {
                            ProgressDetailsText.Text = $"Downloading: {mbRead} MB / {totalMb} MB ({progressPercent}%)";
                        });
                    }
                }
            });

            StatusTitleText.Text = "Running LibreOffice Setup...";
            StatusSubText.Text = "Please complete the installation wizard on your screen.";
            ProgressDetailsText.Text = "Waiting for installation to finish...";

            await Task.Run(() =>
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{tempMsi}\"",
                    UseShellExecute = true
                };
                using var process = Process.Start(psi);
                process?.WaitForExit();
            });

            // Re-check engines
            OfficeEngineDetector.ResetCache();
            if (OfficeEngineDetector.GetAvailableEngines() != OfficeEngineType.None)
            {
                DialogResult = true;
                Close();
                return;
            }

            StatusTitleText.Text = "Setup Finished";
            StatusSubText.Text = "LibreOffice setup completed. Click 'Check Again' to proceed.";
            ProgressDetailsText.Text = "Re-checking installed engines...";
        }
        catch (Exception ex)
        {
            StatusTitleText.Text = "Direct Download Unavailable";
            StatusSubText.Text = "Could not download installer automatically. Opening the official LibreOffice download website in your browser.";
            ProgressDetailsText.Text = $"Notice: {ex.Message}";

            try
            {
                Process.Start(new ProcessStartInfo { FileName = LibreOfficeWebPage, UseShellExecute = true });
            }
            catch { }
        }
        finally
        {
            InstallButton.IsEnabled = true;
            CheckAgainButton.IsEnabled = true;
            ExitButton.IsEnabled = true;
            DownloadProgressBar.Visibility = Visibility.Collapsed;
        }
    }

    private void CheckAgainButton_Click(object sender, RoutedEventArgs e)
    {
        OfficeEngineDetector.ResetCache();
        var engines = OfficeEngineDetector.GetAvailableEngines();

        if (engines != OfficeEngineType.None)
        {
            DialogResult = true;
            Close();
            return;
        }

        ProgressDetailsText.Visibility = Visibility.Visible;
        ProgressDetailsText.Text = "Neither Microsoft Word nor LibreOffice was detected. Please install one to continue.";
    }

    private void ExitButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
