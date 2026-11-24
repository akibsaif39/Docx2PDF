using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using Docx2PDFLive.Core.Models;

namespace Docx2PDFLive.Core.Services;

public class ConversionQueue : IDisposable
{
    private readonly IConverter _converter;
    private readonly BlockingCollection<ConversionItem> _queue = new();
    private readonly ManualResetEventSlim _pauseEvent = new(true);
    private CancellationTokenSource _runCts = new();
    private readonly CancellationTokenSource _disposeCts = new();
    private readonly SemaphoreSlim _concurrencySemaphore;
    private readonly List<Task> _workerTasks = new();
    private int _currentRunId = 1;
    private int _activeConversions = 0;
    private bool _isPaused;
    private bool _disposed;

    public event Action<ConversionItem>? ItemStatusChanged;
    public event Action<bool>? PauseStateChanged;

    public bool IsBusy => _queue.Count > 0 || Volatile.Read(ref _activeConversions) > 0;

    public bool IsPaused
    {
        get => _isPaused;
        private set
        {
            if (_isPaused != value)
            {
                _isPaused = value;
                if (_isPaused)
                {
                    _pauseEvent.Reset();
                }
                else
                {
                    _pauseEvent.Set();
                }
                PauseStateChanged?.Invoke(_isPaused);
            }
        }
    }

    public ConversionQueue(IConverter? converter = null, int maxDegreeOfParallelism = 2)
    {
        _converter = converter ?? new CompositeConverter();
        
        // Bounded parallel workers to balance high throughput with minimal RAM footprint (typically <150MB)
        var maxWorkers = Math.Clamp(maxDegreeOfParallelism, 1, 4);
        _concurrencySemaphore = new SemaphoreSlim(maxWorkers, maxWorkers);

        for (int i = 0; i < maxWorkers; i++)
        {
            _workerTasks.Add(Task.Run(ProcessQueueLoopAsync));
        }
    }

    public void Enqueue(ConversionItem item, bool forceReconvert = false)
    {
        if (_disposed) return;

        // Fast-path: Check if the PDF already exists and is fresher than or equal to the source DOCX
        if (!forceReconvert && File.Exists(item.FilePath))
        {
            var dir = Path.GetDirectoryName(item.FilePath) ?? Directory.GetCurrentDirectory();
            var pdfFileName = Path.GetFileNameWithoutExtension(item.FilePath) + ".pdf";
            var targetPdfPath = Path.Combine(dir, pdfFileName);

            if (File.Exists(targetPdfPath))
            {
                try
                {
                    var docxTime = File.GetLastWriteTimeUtc(item.FilePath);
                    var pdfTime = File.GetLastWriteTimeUtc(targetPdfPath);

                    // If PDF is up-to-date, mark Synced immediately without queue backlog
                    if (pdfTime >= docxTime)
                    {
                        item.Status = ConversionStatus.Synced;
                        item.Engine = "Up to date";
                        item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
                        item.ErrorMessage = null;
                        ItemStatusChanged?.Invoke(item);
                        return;
                    }
                }
                catch { }
            }
        }

        item.RunId = _currentRunId;
        item.Status = ConversionStatus.Queued;
        item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
        _queue.Add(item);
        ItemStatusChanged?.Invoke(item);
    }

    public void CancelAndClear()
    {
        // Increment run generation so pending dequeued items from old folder are instantly dropped
        Interlocked.Increment(ref _currentRunId);

        // 1. Cancel the active conversion run immediately (terminates any running soffice.exe / file streams)
        var oldCts = _runCts;
        _runCts = new CancellationTokenSource();
        try
        {
            oldCts.Cancel();
            oldCts.Dispose();
        }
        catch { }

        // 2. Drain all pending items from the queue
        while (_queue.TryTake(out _)) { }

        // Ensure pause state does not block new tasks
        if (!IsPaused)
        {
            _pauseEvent.Set();
        }
    }

    public void Pause()
    {
        IsPaused = true;
    }

    public void Resume()
    {
        IsPaused = false;
    }

    public void TogglePause()
    {
        if (IsPaused) Resume();
        else Pause();
    }

    private async Task ProcessQueueLoopAsync()
    {
        var disposeToken = _disposeCts.Token;

        while (!disposeToken.IsCancellationRequested && !_queue.IsCompleted)
        {
            try
            {
                _pauseEvent.Wait(disposeToken);

                ConversionItem? item = null;
                try
                {
                    item = _queue.Take(disposeToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }

                if (item == null) continue;

                // Discard stale items belonging to a cancelled previous folder run
                if (item.RunId != _currentRunId) continue;

                _pauseEvent.Wait(disposeToken);

                // Acquire bounded concurrency slot
                await _concurrencySemaphore.WaitAsync(disposeToken);

                _ = Task.Run(async () =>
                {
                    try
                    {
                        var runToken = _runCts.Token;
                        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(runToken, disposeToken);
                        await ProcessSingleItemAsync(item, linkedCts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        // Clean abort when folder was changed or cancelled
                    }
                    finally
                    {
                        _concurrencySemaphore.Release();
                    }
                }, disposeToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Queue loop exception: {ex.Message}");
            }
        }
    }

    private async Task ProcessSingleItemAsync(ConversionItem item, CancellationToken token)
    {
        Interlocked.Increment(ref _activeConversions);
        try
        {
            token.ThrowIfCancellationRequested();

            item.Status = ConversionStatus.Converting;
            item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
            ItemStatusChanged?.Invoke(item);

            // Verify file exists and is unlocked (e.g. Word saving)
            var isReady = await WaitForFileReadyAsync(item.FilePath, maxRetries: 4, delayMs: 400, token);
            if (!isReady)
            {
                if (!File.Exists(item.FilePath))
                {
                    item.Status = ConversionStatus.Error;
                    item.ErrorMessage = "Source file does not exist or was deleted.";
                    item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
                    ItemStatusChanged?.Invoke(item);
                    return;
                }

                item.Status = ConversionStatus.Error;
                item.ErrorMessage = "File remained locked by another application (e.g. Word).";
                item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
                ItemStatusChanged?.Invoke(item);
                return;
            }

            token.ThrowIfCancellationRequested();

            var dir = Path.GetDirectoryName(item.FilePath);
            if (string.IsNullOrEmpty(dir)) dir = Directory.GetCurrentDirectory();

            var pdfFileName = Path.GetFileNameWithoutExtension(item.FilePath) + ".pdf";
            var targetPdfPath = Path.Combine(dir, pdfFileName);

            var result = await _converter.ConvertAsync(item.FilePath, targetPdfPath, token);

            item.Engine = result.EngineUsed;
            item.Timestamp = DateTime.Now.ToString("HH:mm:ss");

            if (result.Success && File.Exists(targetPdfPath))
            {
                item.Status = ConversionStatus.Synced;
                item.ErrorMessage = null;
            }
            else
            {
                item.Status = ConversionStatus.Error;
                item.ErrorMessage = result.ErrorMessage ?? "Conversion failed without error output.";
            }

            ItemStatusChanged?.Invoke(item);
        }
        catch (OperationCanceledException)
        {
            if (item.RunId == _currentRunId)
            {
                item.Status = ConversionStatus.Error;
                item.ErrorMessage = "Cancelled (folder changed or stopped)";
                ItemStatusChanged?.Invoke(item);
            }
        }
        catch (Exception ex)
        {
            item.Status = ConversionStatus.Error;
            item.ErrorMessage = ex.Message;
            item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
            ItemStatusChanged?.Invoke(item);
        }
        finally
        {
            Interlocked.Decrement(ref _activeConversions);
        }
    }

    private static async Task<bool> WaitForFileReadyAsync(string path, int maxRetries = 4, int delayMs = 400, CancellationToken token = default)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            token.ThrowIfCancellationRequested();

            if (!File.Exists(path)) return false;

            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (fs.Length > 0)
                {
                    return true;
                }
            }
            catch (IOException)
            {
                // File locked by saving process
            }

            await Task.Delay(delayMs, token);
        }

        return false;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        CancelAndClear();
        _disposeCts.Cancel();
        _pauseEvent.Set();
        _queue.CompleteAdding();

        try
        {
            Task.WaitAll(_workerTasks.ToArray(), TimeSpan.FromSeconds(2));
        }
        catch { }

        _concurrencySemaphore.Dispose();
        _pauseEvent.Dispose();
        _runCts.Dispose();
        _disposeCts.Dispose();
        _queue.Dispose();

        if (_converter is IDisposable disposableConverter)
        {
            try { disposableConverter.Dispose(); } catch { }
        }
    }
}
