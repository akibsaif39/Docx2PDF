using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using Docx2PDFLive.Core.Models;

namespace Docx2PDFLive.Core.Services;

public class ConversionQueue : IDisposable
{
    private readonly BlockingCollection<ConversionItem> _queue = new();
    private readonly IConverter _converter;
    private readonly CancellationTokenSource _disposeCts = new();
    private readonly ManualResetEventSlim _pauseEvent = new(true);
    private readonly SemaphoreSlim _concurrencySemaphore;
    private readonly Task _workerTask;
    private CancellationTokenSource _runCts = new();
    private int _currentRunId = 0;

    public event Action<ConversionItem>? ItemStatusChanged;
    public bool IsPaused { get; private set; }

    public ConversionQueue(IConverter? converter = null, int maxDegreeOfParallelism = 2)
    {
        _converter = converter ?? new CompositeConverter();
        _concurrencySemaphore = new SemaphoreSlim(Math.Max(1, maxDegreeOfParallelism));
        _workerTask = Task.Run(ProcessQueueLoopAsync);
    }

    public void Enqueue(ConversionItem item, bool forceReconvert = false)
    {
        if (!File.Exists(item.FilePath)) return;

        var dir = Path.GetDirectoryName(item.FilePath) ?? string.Empty;
        var pdfPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(item.FilePath) + ".pdf");

        if (!forceReconvert && File.Exists(pdfPath))
        {
            var docxTime = File.GetLastWriteTimeUtc(item.FilePath);
            var pdfTime = File.GetLastWriteTimeUtc(pdfPath);
            if (pdfTime >= docxTime)
            {
                item.Status = ConversionStatus.Synced;
                item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
                ItemStatusChanged?.Invoke(item);
                return;
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
        Interlocked.Increment(ref _currentRunId);

        var oldCts = _runCts;
        _runCts = new CancellationTokenSource();
        try
        {
            oldCts.Cancel();
            oldCts.Dispose();
        }
        catch { }

        while (_queue.TryTake(out _)) { }

        if (!IsPaused)
        {
            _pauseEvent.Set();
        }
    }

    public void Pause() => IsPaused = true;
    public void Resume() => IsPaused = false;
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

                if (item == null || item.RunId != _currentRunId) continue;

                _pauseEvent.Wait(disposeToken);

                await _concurrencySemaphore.WaitAsync(disposeToken);

                _ = Task.Run(async () =>
                {
                    try
                    {
                        var runToken = _runCts.Token;
                        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(runToken, disposeToken);
                        await ProcessSingleItemAsync(item, linkedCts.Token);
                    }
                    catch (OperationCanceledException) { }
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
        token.ThrowIfCancellationRequested();

        item.Status = ConversionStatus.Converting;
        item.Timestamp = DateTime.Now.ToString("HH:mm:ss");
        ItemStatusChanged?.Invoke(item);

        try
        {
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
                item.ErrorMessage = "File remained locked by another application.";
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
    }

    private static async Task<bool> WaitForFileReadyAsync(string path, int maxRetries = 4, int delayMs = 400, CancellationToken token = default)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (stream.Length > 0) return true;
            }
            catch (IOException)
            {
                await Task.Delay(delayMs, token);
            }
        }
        return false;
    }

    public void Dispose()
    {
        _disposeCts.Cancel();
        _queue.CompleteAdding();
        _pauseEvent.Set();
        try { _workerTask.Wait(1000); } catch { }
        _disposeCts.Dispose();
        _pauseEvent.Dispose();
        _concurrencySemaphore.Dispose();
        _runCts.Dispose();
    }
}
