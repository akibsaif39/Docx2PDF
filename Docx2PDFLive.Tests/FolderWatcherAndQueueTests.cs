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
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_Tests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task ConversionQueue_ProcessesItemSequentially()
    {
        var docxPath = Path.Combine(_tempDir, "test1.docx");
        var pdfPath = Path.Combine(_tempDir, "test1.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "Queue Test", "Queue body text");

        using var queue = new ConversionQueue();
        var tcs = new TaskCompletionSource<bool>();

        queue.ItemStatusChanged += item =>
        {
            if (item.FilePath == docxPath && item.Status == ConversionStatus.Synced)
            {
                tcs.TrySetResult(true);
            }
        };

        var item = new ConversionItem { FilePath = docxPath, RelativePath = "test1.docx" };
        queue.Enqueue(item);

        var completed = await Task.WhenAny(tcs.Task, Task.Delay(15000));
        Assert.True(completed == tcs.Task, "Item did not reach Synced status within timeout.");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
    }

    [Fact]
    public void ConversionQueue_PauseAndResume_TogglesCorrectly()
    {
        using var queue = new ConversionQueue();

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
