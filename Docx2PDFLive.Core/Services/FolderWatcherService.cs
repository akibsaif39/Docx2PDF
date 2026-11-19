using System.IO;

namespace Docx2PDFLive.Core.Services;

public class FolderWatcherService : IDisposable
{
    private FileSystemWatcher? _watcher;
    public string? CurrentFolderPath { get; private set; }
    public bool IsWatching => _watcher != null && _watcher.EnableRaisingEvents;

    public event Action<string>? FileDetected;
    public event Action<string>? WatcherError;

    public void Start(string folderPath)
    {
        Stop();

        if (!Directory.Exists(folderPath))
        {
            throw new DirectoryNotFoundException($"Folder not found: {folderPath}");
        }

        CurrentFolderPath = folderPath;

        _watcher = new FileSystemWatcher(folderPath, "*.docx")
        {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime
        };

        _watcher.Created += OnFileEvent;
        _watcher.Changed += OnFileEvent;
        _watcher.Renamed += (s, e) =>
        {
            if (e.FullPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                OnFileEvent(s, e);
            }
        };

        _watcher.Error += (s, e) =>
        {
            WatcherError?.Invoke(e.GetException()?.Message ?? "FileSystemWatcher encountered an error.");
        };

        _watcher.EnableRaisingEvents = true;
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e)
    {
        var fileName = Path.GetFileName(e.FullPath);
        if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase)) return;

        FileDetected?.Invoke(e.FullPath);
    }

    public void Stop()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Dispose();
            _watcher = null;
        }
        CurrentFolderPath = null;
    }

    public void Dispose() => Stop();
}
