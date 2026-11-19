using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Docx2PDFLive.Setup;

static class Program
{
    private const string LibreOfficeDownloadUrl = "https://download.documentfoundation.org/libreoffice/stable/24.8.3/win/x86_64/LibreOffice_24.8.3_Win_x86-64.msi";
    private const string LibreOfficeWebPage = "https://www.libreoffice.org/download/download-libreoffice/";
    private const string InstallerMutexName = @"Local\Docx2PDF_Installer_SingleInstance_Mutex_9F842FDEE031";

    [STAThread]
    static void Main()
    {
        using var mutex = new Mutex(true, InstallerMutexName, out bool isNewInstance);
        if (!isNewInstance)
        {
            return;
        }

        ApplicationConfiguration.Initialize();

        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var msiCandidates = new[]
        {
            Path.Combine(baseDir, "Docx2PDF_Setup.msi"),
            Path.Combine(baseDir, "Docx2PDFLive_Setup.msi"),
            Path.Combine(baseDir, "Installer", "Docx2PDF_Setup.msi"),
            Path.Combine(baseDir, "Installer", "Docx2PDFLive_Setup.msi"),
            Path.Combine(baseDir, "..", "Installer", "Docx2PDF_Setup.msi"),
            Path.Combine(baseDir, "..", "Installer", "Docx2PDFLive_Setup.msi")
        };

        var msiPath = msiCandidates.FirstOrDefault(File.Exists) ?? Path.Combine(baseDir, "Installer", "Docx2PDF_Setup.msi");

        Application.Run(new SetupForm(msiPath));
    }

    public static bool CheckLibreOfficeInstalled()
    {
        string[] standardPaths =
        [
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe"),
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ];

        foreach (var p in standardPaths)
        {
            if (File.Exists(p)) return true;
        }

        try
        {
            using var k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\LibreOffice");
            var val = k?.GetValue("Path") as string;
            if (!string.IsNullOrEmpty(val) && File.Exists(Path.Combine(val, "program", "soffice.exe")))
            {
                return true;
            }
        }
        catch { }

        return false;
    }

    public static void DownloadAndInstallLibreOffice(IWin32Window parent)
    {
        try
        {
            var tempMsi = Path.Combine(Path.GetTempPath(), "LibreOffice_Setup.msi");

            using var form = new Form
            {
                Text = "Downloading LibreOffice...",
                Width = 420,
                Height = 150,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White
            };

            var label = new Label
            {
                Text = "Downloading official LibreOffice package from mirror...\nPlease wait.",
                Left = 24,
                Top = 20,
                Width = 360,
                Height = 40,
                Font = new Font("Segoe UI", 9.5f)
            };
            var pbar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                Left = 24,
                Top = 68,
                Width = 356,
                Height = 22
            };
            form.Controls.Add(label);
            form.Controls.Add(pbar);

            form.Shown += async (s, e) =>
            {
                try
                {
                    using var client = new HttpClient();
                    client.Timeout = TimeSpan.FromMinutes(10);
                    using var stream = await client.GetStreamAsync(LibreOfficeDownloadUrl);
                    using var fs = new FileStream(tempMsi, FileMode.Create, FileAccess.Write, FileShare.None);
                    await stream.CopyToAsync(fs);

                    form.Close();

                    var psi = new ProcessStartInfo
                    {
                        FileName = "msiexec.exe",
                        Arguments = $"/i \"{tempMsi}\"",
                        UseShellExecute = true
                    };
                    using var p = Process.Start(psi);
                    p?.WaitForExit();
                }
                catch (Exception ex)
                {
                    form.Close();
                    MessageBox.Show(
                        $"Could not download installer directly ({ex.Message}).\nOpening the official LibreOffice download page in your browser.",
                        "Download LibreOffice",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    Process.Start(new ProcessStartInfo { FileName = LibreOfficeWebPage, UseShellExecute = true });
                }
            };

            form.ShowDialog(parent);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error launching LibreOffice setup: {ex.Message}", "Setup Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

public class SetupForm : Form
{
    private readonly string _msiPath;
    private readonly bool _isLoInstalled;
    private Button _btnInstall = null!;
    private Button _btnCancel = null!;
    private ProgressBar _progressBar = null!;
    private Label _statusLabel = null!;
    private bool _isInstalled = false;

    public SetupForm(string msiPath)
    {
        _msiPath = msiPath;
        _isLoInstalled = Program.CheckLibreOfficeInstalled();

        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Docx2PDF Setup";
        Width = 520;
        Height = 420;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = true;
        BackColor = Color.White;

        var assembly = Assembly.GetExecutingAssembly();
        using var iconStream = assembly.GetManifestResourceStream("Docx2PDFLive.Setup.Resources.app.ico")
            ?? assembly.GetManifestResourceStream(assembly.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("app.ico", StringComparison.OrdinalIgnoreCase)) ?? "");
        if (iconStream != null)
        {
            try { Icon = new Icon(iconStream); } catch { }
        }

        Image? appImage = null;
        using var imgStream = assembly.GetManifestResourceStream("Docx2PDFLive.Setup.Resources.app_icon.png")
            ?? assembly.GetManifestResourceStream(assembly.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("app_icon.png", StringComparison.OrdinalIgnoreCase)) ?? "");
        if (imgStream != null)
        {
            try { appImage = Image.FromStream(imgStream); } catch { }
        }

        var headerPanel = new Panel
        {
            Left = 0,
            Top = 0,
            Width = 520,
            Height = 110,
            BackColor = Color.FromArgb(248, 250, 252)
        };

        var logoBox = new PictureBox
        {
            Left = 28,
            Top = 18,
            Width = 72,
            Height = 72,
            SizeMode = PictureBoxSizeMode.Zoom,
            Image = appImage
        };
        headerPanel.Controls.Add(logoBox);

        var titleLabel = new Label
        {
            Text = "Docx2PDF",
            Left = 112,
            Top = 20,
            Width = 300,
            Height = 34,
            Font = new Font("Segoe UI", 18f, FontStyle.Bold),
            ForeColor = Color.FromArgb(15, 23, 42)
        };
        headerPanel.Controls.Add(titleLabel);

        var subtitleLabel = new Label
        {
            Text = "Real-time DOCX to PDF converter",
            Left = 114,
            Top = 56,
            Width = 350,
            Height = 22,
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.FromArgb(100, 116, 139)
        };
        headerPanel.Controls.Add(subtitleLabel);

        var creditLabel = new Label
        {
            Text = "Created by Akib",
            Left = 114,
            Top = 78,
            Width = 200,
            Height = 20,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(148, 163, 184)
        };
        headerPanel.Controls.Add(creditLabel);

        Controls.Add(headerPanel);

        var cardPanel = new Panel
        {
            Left = 28,
            Top = 130,
            Width = 448,
            Height = 120,
            BackColor = Color.FromArgb(241, 245, 249)
        };

        var engineTitle = new Label
        {
            Left = 18,
            Top = 16,
            Width = 410,
            Height = 24,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold)
        };

        var engineDesc = new Label
        {
            Left = 18,
            Top = 44,
            Width = 410,
            Height = 40,
            Font = new Font("Segoe UI", 9f),
            ForeColor = Color.FromArgb(71, 85, 105)
        };

        if (_isLoInstalled)
        {
            engineTitle.Text = "High-Fidelity LibreOffice Engine Detected";
            engineTitle.ForeColor = Color.FromArgb(22, 163, 74);
            engineDesc.Text = "Conversions will match Microsoft Word formatting and layouts with 100% accuracy in the background.";
        }
        else
        {
            engineTitle.Text = "Built-in OpenXML + PDFSharp Engine Active";
            engineTitle.ForeColor = Color.FromArgb(37, 99, 235);
            engineDesc.Text = "Fast, lightweight built-in conversion without third-party dependencies. For 100% Word matching, you can optionally install LibreOffice.";

            var btnGetLo = new Button
            {
                Text = "Get LibreOffice",
                Left = 18,
                Top = 86,
                Width = 120,
                Height = 26,
                BackColor = Color.White,
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI", 8.5f)
            };
            btnGetLo.Click += (s, e) => Program.DownloadAndInstallLibreOffice(this);
            cardPanel.Controls.Add(btnGetLo);
        }

        cardPanel.Controls.Add(engineTitle);
        cardPanel.Controls.Add(engineDesc);
        Controls.Add(cardPanel);

        _progressBar = new ProgressBar
        {
            Left = 28,
            Top = 270,
            Width = 448,
            Height = 14,
            Visible = false,
            Style = ProgressBarStyle.Marquee
        };
        Controls.Add(_progressBar);

        _statusLabel = new Label
        {
            Text = "Ready to install Docx2PDF to Program Files.",
            Left = 28,
            Top = 290,
            Width = 448,
            Height = 20,
            Font = new Font("Segoe UI", 9f),
            ForeColor = Color.FromArgb(100, 116, 139)
        };
        Controls.Add(_statusLabel);

        _btnCancel = new Button
        {
            Text = "Cancel",
            Left = 250,
            Top = 325,
            Width = 100,
            Height = 36,
            Font = new Font("Segoe UI", 9.5f),
            BackColor = Color.FromArgb(241, 245, 249),
            FlatStyle = FlatStyle.Flat
        };
        _btnCancel.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        _btnCancel.Click += (s, e) => Close();
        Controls.Add(_btnCancel);

        _btnInstall = new Button
        {
            Text = "Install Docx2PDF",
            Left = 360,
            Top = 325,
            Width = 116,
            Height = 36,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            BackColor = Color.FromArgb(37, 99, 235),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _btnInstall.FlatAppearance.BorderSize = 0;
        _btnInstall.Click += OnInstallClicked;
        Controls.Add(_btnInstall);
    }

    private async void OnInstallClicked(object? sender, EventArgs e)
    {
        if (_isInstalled)
        {
            var installedExe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Docx2PDF", "Docx2PDF.exe");
            if (!File.Exists(installedExe))
            {
                installedExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".dev", "Docx2PDF.exe");
            }

            if (File.Exists(installedExe))
            {
                Process.Start(new ProcessStartInfo { FileName = installedExe, UseShellExecute = true });
            }

            Close();
            return;
        }

        if (!File.Exists(_msiPath))
        {
            MessageBox.Show(
                $"Could not locate installer package:\n{_msiPath}\n\nPlease ensure Docx2PDF_Setup.msi is present in the Installer directory.",
                "Docx2PDF Setup",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _btnInstall.Enabled = false;
        _btnCancel.Enabled = false;
        _progressBar.Visible = true;
        _statusLabel.Text = "Installing Docx2PDF... Please wait.";

        try
        {
            await Task.Run(() =>
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{_msiPath}\"",
                    UseShellExecute = true
                };
                using var p = Process.Start(psi);
                p?.WaitForExit();
            });

            _progressBar.Visible = false;
            _statusLabel.Text = "Installation completed successfully!";
            _statusLabel.ForeColor = Color.FromArgb(22, 163, 74);

            _isInstalled = true;
            _btnInstall.Text = "Launch";
            _btnInstall.Enabled = true;
            _btnCancel.Text = "Close";
            _btnCancel.Enabled = true;
        }
        catch (Exception ex)
        {
            _progressBar.Visible = false;
            _statusLabel.Text = $"Error: {ex.Message}";
            _statusLabel.ForeColor = Color.Red;
            _btnInstall.Enabled = true;
            _btnCancel.Enabled = true;
        }
    }
}
