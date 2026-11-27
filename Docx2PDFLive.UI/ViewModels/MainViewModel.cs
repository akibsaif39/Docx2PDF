using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;
using Docx2PDFLive.UI.Views;

namespace Docx2PDFLive.UI.ViewModels;

public class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly ConversionQueue _queue;
    private readonly FolderWatcherService _watcherService;
    private readonly AppSettingsService _settingsService;
    private string? _watchedPath;
    private bool _rememberLastFolder;
    private bool _isPaused;
    private string _statusSummary = "Idle - Select a folder to begin";
    private bool _disposed;

    public ObservableCollection<ConversionItem> ConversionItems { get; } = new();

    public string WatchedPath => string.IsNullOrEmpty(_watchedPath)
        ? "No folder selected (Idle)"
        : _watchedPath;

    public bool HasWatchedFolder => !string.IsNullOrEmpty(_watchedPath);

    public bool RememberLastFolder
    {
        get => _rememberLastFolder;
        set
        {
            if (_rememberLastFolder != value)
            {
                _rememberLastFolder = value;
                OnPropertyChanged();
                PersistSettings();
            }
        }
    }

    public bool IsPaused
    {
        get => _isPaused;
        private set
        {
            if (_isPaused != value)
            {
                _isPaused = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PauseResumeText));
                OnPropertyChanged(nameof(PauseResumeIcon));
            }
        }
    }

    public string PauseResumeText => IsPaused ? "Resume" : "Pause";
    public string PauseResumeIcon => IsPaused ? "▶" : "⏸";

    public string AppVersionText
    {
        get
        {
            var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            return ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}" : "v1.0.0";
        }
    }

    public string StatusSummary
    {
        get => _statusSummary;
        private set
        {
            if (_statusSummary != value)
            {
                _statusSummary = value;
                OnPropertyChanged();
            }
        }
    }

    public string ActiveEngineDescription { get; }
    public bool IsLibreOfficeMissing { get; }
    public bool IsQueueBusy => _queue.IsBusy;

    public ICommand SelectFolderCommand { get; }
    public ICommand PauseResumeCommand { get; }
    public ICommand ResyncAllCommand { get; }
    public ICommand OpenAboutCommand { get; }
    public ICommand OpenFolderInExplorerCommand { get; }
    public ICommand GetLibreOfficeCommand { get; }

    public MainViewModel()
    {
        _settingsService = new AppSettingsService();
        _rememberLastFolder = _settingsService.Load().RememberLastFolder;

        var fastest = OfficeEngineDetector.GetFastestAvailableEngine();
        ActiveEngineDescription = fastest == OfficeEngineType.MicrosoftWord
            ? "Microsoft Word (Native High-Speed)"
            : (fastest == OfficeEngineType.LibreOffice ? "LibreOffice Headless" : "High-Fidelity Engine");

        IsLibreOfficeMissing = !OfficeEngineDetector.IsLibreOfficeInstalled(out _);

        _queue = new ConversionQueue();
        _queue.ItemStatusChanged += OnQueueItemStatusChanged;
        _queue.PauseStateChanged += OnQueuePauseStateChanged;

        _watcherService = new FolderWatcherService(_queue);
        _watcherService.FileDiscovered += OnFileDiscovered;

        SelectFolderCommand = new RelayCommand(PromptSelectFolder);
        PauseResumeCommand = new RelayCommand(() => _queue.TogglePause(), () => HasWatchedFolder);
        ResyncAllCommand = new RelayCommand(() => _watcherService.ResyncAll(forceReconvert: true), () => HasWatchedFolder);
        OpenAboutCommand = new RelayCommand(ShowAboutDialog);
        OpenFolderInExplorerCommand = new RelayCommand(OpenWatchedFolderInExplorer, () => HasWatchedFolder);
        GetLibreOfficeCommand = new RelayCommand(LibreOfficeDetector.OpenDownloadPage);
    }

    public void InitializeFolderOnStartup()
    {
        var saved = _settingsService.Load();
        _rememberLastFolder = saved.RememberLastFolder;
        OnPropertyChanged(nameof(RememberLastFolder));

        var resolution = AppSettingsService.ResolveStartupFolder(saved);
        switch (resolution.Action)
        {
            case StartupFolderAction.AutoResumeFolder:
                StartWatchingFolder(resolution.FolderPath!);
                return;

            case StartupFolderAction.MissingRememberedFolderPrompt:
                MessageBox.Show(
                    $"Previously watched folder was not found:\n{resolution.FolderPath}\n\nPlease select a folder to continue watching.",
                    "Docx2PDF",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Clear the missing folder so canceling does not repeat the same warning on next launch
                _settingsService.Save(new AppSettings
                {
                    RememberLastFolder = _rememberLastFolder,
                    LastWatchedFolder = null
                });
                break;
        }

        PromptSelectFolder();
    }

    public void PromptSelectFolder()
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Select Root Folder to Watch for DOCX Files",
                Multiselect = false
            };

            if (HasWatchedFolder && Directory.Exists(_watchedPath))
            {
                dialog.InitialDirectory = _watchedPath;
            }

            if (dialog.ShowDialog() == true)
            {
                StartWatchingFolder(dialog.FolderName);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Unable to open folder dialog: {ex.Message}", "Docx2PDF Live", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void StartWatchingFolder(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            return;
        }

        // Mid-conversion folder switch: IMMEDIATELY abort any active conversion in progress
        // and drain previous folder queue items
        _queue.CancelAndClear();

        _watchedPath = folderPath;
        OnPropertyChanged(nameof(WatchedPath));
        OnPropertyChanged(nameof(HasWatchedFolder));

        if (_rememberLastFolder)
        {
            PersistSettings();
        }

        // Clear UI items from previous folder
        SafeInvoke(() =>
        {
            ConversionItems.Clear();
            UpdateSummary();
        });

        // Immediately start watching and converting documents in the new folder
        _watcherService.StartWatching(folderPath);
        UpdateSummary();
    }

    private void OnFileDiscovered(ConversionItem item)
    {
        SafeInvoke(() =>
        {
            var existing = ConversionItems.FirstOrDefault(x => string.Equals(x.FilePath, item.FilePath, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                ConversionItems.Insert(0, item);
            }
            else
            {
                existing.Status = item.Status;
                existing.Timestamp = item.Timestamp;
                existing.Engine = item.Engine;
                existing.ErrorMessage = item.ErrorMessage;
            }
            UpdateSummary();
        });
    }

    private void OnQueueItemStatusChanged(ConversionItem item)
    {
        SafeInvoke(() =>
        {
            var existing = ConversionItems.FirstOrDefault(x => string.Equals(x.FilePath, item.FilePath, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Status = item.Status;
                existing.Timestamp = item.Timestamp;
                existing.Engine = item.Engine;
                existing.ErrorMessage = item.ErrorMessage;
            }
            else
            {
                ConversionItems.Insert(0, item);
            }
            UpdateSummary();
        });
    }

    private void OnQueuePauseStateChanged(bool isPaused)
    {
        SafeInvoke(() =>
        {
            IsPaused = isPaused;
            UpdateSummary();
        });
    }

    private void UpdateSummary()
    {
        if (!HasWatchedFolder)
        {
            StatusSummary = "Idle - Click 'Select Folder' to start watching";
            return;
        }

        var total = ConversionItems.Count;
        var synced = ConversionItems.Count(x => x.Status == ConversionStatus.Synced);
        var queued = ConversionItems.Count(x => x.Status == ConversionStatus.Queued);
        var converting = ConversionItems.Count(x => x.Status == ConversionStatus.Converting);
        var errors = ConversionItems.Count(x => x.Status == ConversionStatus.Error);

        var pauseState = IsPaused ? " [PAUSED]" : "";
        StatusSummary = $"Watching: {total} files | Synced: {synced} | Queued: {queued} | In Progress: {converting} | Errors: {errors}{pauseState}";
    }

    public void ShowAboutDialog()
    {
        var aboutWin = new FirstRunWelcomeWindow
        {
            Owner = Application.Current.MainWindow,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };
        aboutWin.ShowDialog();
    }

    private void OpenWatchedFolderInExplorer()
    {
        if (HasWatchedFolder && Directory.Exists(_watchedPath))
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _watchedPath,
                    UseShellExecute = true
                });
            }
            catch { }
        }
    }

    private void PersistSettings()
    {
        _settingsService.Save(new AppSettings
        {
            RememberLastFolder = _rememberLastFolder,
            LastWatchedFolder = _rememberLastFolder && !string.IsNullOrEmpty(_watchedPath) && Directory.Exists(_watchedPath)
                ? _watchedPath
                : null
        });
    }

    private static void SafeInvoke(Action action)
    {
        if (Application.Current?.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess())
        {
            Application.Current.Dispatcher.InvokeAsync(action);
        }
        else
        {
            action();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _queue.CancelAndClear();
        _watcherService.Dispose();
        _queue.Dispose();
    }
}
