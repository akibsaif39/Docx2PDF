using System.IO;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class FolderWatcherAndQueueTests : IDisposable
{
    private readonly string _tempDir;

    public FolderWatcherAndQueueTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_QueueTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task ConversionQueue_ProcessesItemSequentially()
    {
        var docxPath = Path.Combine(_tempDir, "queue_test.docx");
        TestDocumentHelper.CreateSampleDocx(docxPath, "Queue Doc", "Testing queue conversion.");

        using var queue = new ConversionQueue();
        var item = new ConversionItem
        {
            FilePath = docxPath,
            RelativePath = "queue_test.docx"
        };

        var completedTcs = new TaskCompletionSource<bool>();
        queue.ItemStatusChanged += updatedItem =>
        {
            if (updatedItem.FilePath == docxPath && updatedItem.Status == ConversionStatus.Synced)
            {
                completedTcs.TrySetResult(true);
            }
            else if (updatedItem.FilePath == docxPath && updatedItem.Status == ConversionStatus.Error)
            {
                completedTcs.TrySetException(new Exception(updatedItem.ErrorMessage));
            }
        };

        queue.Enqueue(item);

        var completed = await Task.WhenAny(completedTcs.Task, Task.Delay(15000));
        Assert.True(completed == completedTcs.Task, "Queue failed to process item within timeout.");
        Assert.True(await completedTcs.Task);

        var expectedPdf = Path.Combine(_tempDir, "queue_test.pdf");
        Assert.True(File.Exists(expectedPdf), "Target PDF must exist beside DOCX.");
    }

    [Fact]
    public void ConversionQueue_PauseAndResume_TogglesCorrectly()
    {
        using var queue = new ConversionQueue();
        Assert.False(queue.IsPaused);

        queue.Pause();
        Assert.True(queue.IsPaused);

        queue.Resume();
        Assert.False(queue.IsPaused);

        queue.TogglePause();
        Assert.True(queue.IsPaused);

        queue.TogglePause();
        Assert.False(queue.IsPaused);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch { }
    }
}
