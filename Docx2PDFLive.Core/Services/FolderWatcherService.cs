using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using Docx2PDFLive.Core.Models;

namespace Docx2PDFLive.Core.Services;

public class FolderWatcherService : IDisposable
{
    private readonly ConversionQueue _queue;
    private readonly ConcurrentDictionary<string, CancellationTokenSource> _debounceMap = new(StringComparer.OrdinalIgnoreCase);
    private FileSystemWatcher? _watcher;
    private string? _watchedRootPath;
    private bool _disposed;

    public string? WatchedRootPath => _watchedRootPath;
    public bool IsWatching => _watcher != null && _watcher.EnableRaisingEvents;

    public event Action<ConversionItem>? FileDiscovered;
    public event Action<string>? WatchedFolderChanged;

    public FolderWatcherService(ConversionQueue queue)
    {
        _queue = queue;
    }

    public void StartWatching(string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
        {
            throw new DirectoryNotFoundException($"Directory not found: {rootPath}");
        }

        StopWatching();

        _watchedRootPath = Path.GetFullPath(rootPath);

        _watcher = new FileSystemWatcher(_watchedRootPath)
        {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
            Filter = "*.*" // Listen to all and filter in handler for robustness
        };

        _watcher.Created += OnFileSystemEvent;
        _watcher.Changed += OnFileSystemEvent;
        _watcher.Renamed += OnFileSystemRenamed;
        _watcher.Error += OnWatcherError;

        _watcher.EnableRaisingEvents = true;
        WatchedFolderChanged?.Invoke(_watchedRootPath);

        // Perform initial sync of all existing files (respecting timestamp freshness)
        ResyncAll(forceReconvert: false);
    }

    public void StopWatching()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Created -= OnFileSystemEvent;
            _watcher.Changed -= OnFileSystemEvent;
            _watcher.Renamed -= OnFileSystemRenamed;
            _watcher.Error -= OnWatcherError;
            _watcher.Dispose();
            _watcher = null;
        }

        // Cancel all pending debounce tasks immediately
        foreach (var key in _debounceMap.Keys)
        {
            if (_debounceMap.TryRemove(key, out var cts))
            {
                try
                {
                    cts.Cancel();
                    cts.Dispose();
                }
                catch { }
            }
        }
    }

    public void ResyncAll(bool forceReconvert = false)
    {
        if (string.IsNullOrEmpty(_watchedRootPath) || !Directory.Exists(_watchedRootPath))
        {
            return;
        }

        try
        {
            var docxFiles = Directory.GetFiles(_watchedRootPath, "*.docx", SearchOption.AllDirectories);
            foreach (var file in docxFiles)
            {
                if (IsValidDocx(file))
                {
                    EnqueueFile(file, forceReconvert);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error scanning folder for resync: {ex.Message}");
        }
    }

    private void OnFileSystemEvent(object sender, FileSystemEventArgs e)
    {
        if (IsValidDocx(e.FullPath))
        {
            // Changed or created during live monitoring: force reconvert so PDF updates immediately!
            DebounceAndEnqueue(e.FullPath, forceReconvert: true);
        }
    }

    private void OnFileSystemRenamed(object sender, RenamedEventArgs e)
    {
        if (IsValidDocx(e.FullPath))
        {
            DebounceAndEnqueue(e.FullPath, forceReconvert: true);
        }
    }

    private void OnWatcherError(object sender, ErrorEventArgs e)
    {
        Debug.WriteLine($"FileSystemWatcher internal buffer error: {e.GetException()?.Message}");
    }

    private static bool IsValidDocx(string fullPath)
    {
        try
        {
            var fileName = Path.GetFileName(fullPath);
            if (string.IsNullOrEmpty(fileName)) return false;

            // Ignore temporary lock files starting with ~$
            if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase)) return false;

            // Ignore PDF files
            if (fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) return false;

            // Only process .docx
            return fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private void DebounceAndEnqueue(string fullPath, bool forceReconvert)
    {
        var normalizedPath = Path.GetFullPath(fullPath);

        // Cancel previous timer for this file if rapid saves occur
        if (_debounceMap.TryGetValue(normalizedPath, out var existingCts))
        {
            try
            {
                existingCts.Cancel();
                existingCts.Dispose();
            }
            catch { }
        }

        var cts = new CancellationTokenSource();
        _debounceMap[normalizedPath] = cts;

        Task.Run(async () =>
        {
            try
            {
                // Debounce 400 ms
                await Task.Delay(400, cts.Token);

                if (!cts.Token.IsCancellationRequested)
                {
                    _debounceMap.TryRemove(normalizedPath, out _);
                    EnqueueFile(normalizedPath, forceReconvert);
                }
            }
            catch (OperationCanceledException)
            {
                // Expected when debounced
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Debounce error: {ex.Message}");
            }
            finally
            {
                cts.Dispose();
            }
        });
    }

    private void EnqueueFile(string fullPath, bool forceReconvert)
    {
        var relativePath = GetRelativePath(_watchedRootPath, fullPath);

        var item = new ConversionItem
        {
            FilePath = fullPath,
            RelativePath = relativePath,
            Status = ConversionStatus.Queued,
            Timestamp = DateTime.Now.ToString("HH:mm:ss")
        };

        FileDiscovered?.Invoke(item);
        _queue.Enqueue(item, forceReconvert);
    }

    private static string GetRelativePath(string? rootPath, string fullPath)
    {
        if (string.IsNullOrEmpty(rootPath)) return Path.GetFileName(fullPath);

        try
        {
            return Path.GetRelativePath(rootPath, fullPath);
        }
        catch
        {
            return Path.GetFileName(fullPath);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        StopWatching();
    }
}
