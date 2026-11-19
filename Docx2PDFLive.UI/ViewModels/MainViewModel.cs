using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;
using MessageBox = System.Windows.MessageBox;

namespace Docx2PDFLive.UI.ViewModels;

public class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly FolderWatcherService _watcher = new();
    private readonly ConversionQueue _queue;
    private string? _selectedFolder;
    private string _engineStatus = "Detecting...";
    private string _statusSummary = "Ready. Select a folder to begin watching.";
    private bool _isLibreOfficeMissing = false;

    public ObservableCollection<ConversionItem> Items { get; } = new();

    public string SelectedFolderDisplay =>
        string.IsNullOrEmpty(_selectedFolder) ? "No folder selected — Click 'Select Folder' to begin" : _selectedFolder;

    public string EngineStatus
    {
        get => _engineStatus;
        private set { _engineStatus = value; OnPropertyChanged(); }
    }

    public string StatusSummary
    {
        get => _statusSummary;
        private set { _statusSummary = value; OnPropertyChanged(); }
    }

    public bool IsWatching => _watcher.IsWatching;
    public string WatchingStatusText => _watcher.IsWatching ? "Watching" : "Idle";
    public bool IsLibreOfficeMissing
    {
        get => _isLibreOfficeMissing;
        private set { _isLibreOfficeMissing = value; OnPropertyChanged(); }
    }

    public string PauseResumeButtonText => _queue.IsPaused ? "Resume" : "Pause";
    public string PauseResumeIcon => _queue.IsPaused ? "▶" : "⏸";

    public ICommand SelectFolderCommand { get; }
    public ICommand PauseResumeCommand { get; }
    public ICommand ForceReconvertAllCommand { get; }
    public ICommand GetLibreOfficeCommand { get; }

    public MainViewModel()
    {
        _queue = new ConversionQueue(maxDegreeOfParallelism: 2);
        _queue.ItemStatusChanged += OnItemStatusChanged;

        _watcher.FileDetected += OnFileDetected;
        _watcher.WatcherError += OnWatcherError;

        SelectFolderCommand = new RelayCommand(_ => SelectFolder());
        PauseResumeCommand = new RelayCommand(_ => TogglePause());
        ForceReconvertAllCommand = new RelayCommand(_ => ReconvertAll(), _ => !string.IsNullOrEmpty(_selectedFolder));
        GetLibreOfficeCommand = new RelayCommand(_ => OpenLibreOfficeDownload());

        DetectEngine();
    }

    private void DetectEngine()
    {
        if (LibreOfficeDetector.IsAvailable(out _))
        {
            EngineStatus = "LibreOffice";
            IsLibreOfficeMissing = false;
        }
        else
        {
            EngineStatus = "OpenXML (LibreOffice Missing)";
            IsLibreOfficeMissing = true;
        }
    }

    private void SelectFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Select a folder to watch and auto-convert DOCX files to PDF",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
        {
            _selectedFolder = dialog.SelectedPath;
            OnPropertyChanged(nameof(SelectedFolderDisplay));

            _queue.CancelAndClear();
            Items.Clear();

            _watcher.Start(_selectedFolder);
            OnPropertyChanged(nameof(IsWatching));
            OnPropertyChanged(nameof(WatchingStatusText));

            InitialScanFolder(_selectedFolder);

            StatusSummary = $"Watching: {_selectedFolder}";
        }
    }

    private void InitialScanFolder(string rootPath)
    {
        try
        {
            var files = Directory.GetFiles(rootPath, "*.docx", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var fileName = Path.GetFileName(file);
                if (fileName.StartsWith("~$")) continue;

                var rel = Path.GetRelativePath(rootPath, file);
                var item = new ConversionItem { FilePath = file, RelativePath = rel };
                Items.Add(item);
                _queue.Enqueue(item, forceReconvert: false);
            }
        }
        catch (Exception ex)
        {
            StatusSummary = $"Scan error: {ex.Message}";
        }
    }

    private void OnFileDetected(string fullPath)
    {
        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
        {
            if (string.IsNullOrEmpty(_selectedFolder)) return;

            var existing = Items.FirstOrDefault(i => string.Equals(i.FilePath, fullPath, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                var rel = Path.GetRelativePath(_selectedFolder, fullPath);
                existing = new ConversionItem { FilePath = fullPath, RelativePath = rel };
                Items.Insert(0, existing);
            }

            _queue.Enqueue(existing, forceReconvert: false);
        });
    }

    private void OnItemStatusChanged(ConversionItem item)
    {
        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
        {
            var total = Items.Count;
            var synced = Items.Count(i => i.Status == ConversionStatus.Synced);
            var converting = Items.Count(i => i.Status == ConversionStatus.Converting);
            var errors = Items.Count(i => i.Status == ConversionStatus.Error);

            StatusSummary = $"Files: {total} | Synced: {synced} | Converting: {converting} | Errors: {errors}";
        });
    }

    private void OnWatcherError(string message)
    {
        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
        {
            StatusSummary = $"Watcher error: {message}";
        });
    }

    private void TogglePause()
    {
        _queue.TogglePause();
        OnPropertyChanged(nameof(PauseResumeButtonText));
        OnPropertyChanged(nameof(PauseResumeIcon));
        StatusSummary = _queue.IsPaused ? "Queue paused." : "Queue resumed.";
    }

    private void ReconvertAll()
    {
        if (string.IsNullOrEmpty(_selectedFolder)) return;

        foreach (var item in Items)
        {
            _queue.Enqueue(item, forceReconvert: true);
        }
        StatusSummary = $"Re-queued {Items.Count} files for re-conversion.";
    }

    private void OpenLibreOfficeDownload()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://www.libreoffice.org/download/download-libreoffice/",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open browser: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    public void ShowAboutDialog()
    {
        MessageBox.Show(
            "Docx2PDF\n" +
            "Real-time DOCX to PDF converter.\n\n" +
            "Created by Akib",
            "About Docx2PDF",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    public void Dispose()
    {
        _watcher.Dispose();
        _queue.Dispose();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
