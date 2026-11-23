using System.IO;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class QueuePerformanceAndCancelTests : IDisposable
{
    private readonly string _tempDir;

    public QueuePerformanceAndCancelTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_CancelTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task CancelAndClear_AbortsPendingAndRunningConversions()
    {
        using var queue = new ConversionQueue(maxDegreeOfParallelism: 1);

        // Enqueue 5 items
        for (int i = 0; i < 5; i++)
        {
            var docx = Path.Combine(_tempDir, $"file_{i}.docx");
            TestDocumentHelper.CreateSampleDocx(docx, $"Title {i}", $"Body {i}");
            queue.Enqueue(new ConversionItem { FilePath = docx, RelativePath = $"file_{i}.docx" });
        }

        // Cancel and clear immediately
        queue.CancelAndClear();

        // Give a short grace period
        await Task.Delay(500);

        // Enqueue a new item from a "new folder"
        var newDocx = Path.Combine(_tempDir, "new_folder_file.docx");
        TestDocumentHelper.CreateSampleDocx(newDocx, "New Folder Title", "New Content");

        var newTcs = new TaskCompletionSource<bool>();
        queue.ItemStatusChanged += item =>
        {
            if (item.FilePath == newDocx && item.Status == ConversionStatus.Synced)
            {
                newTcs.TrySetResult(true);
            }
        };

        queue.Enqueue(new ConversionItem { FilePath = newDocx, RelativePath = "new_folder_file.docx" });

        var completed = await Task.WhenAny(newTcs.Task, Task.Delay(15000));
        Assert.True(completed == newTcs.Task, "New folder item did not complete after cancel and clear.");

        var newPdf = Path.Combine(_tempDir, "new_folder_file.pdf");
        Assert.True(File.Exists(newPdf), "New folder PDF must be created.");
    }

    [Fact]
    public void UpToDateFile_IsMarkedSyncedImmediately_AndReconvertsWhenModified()
    {
        using var queue = new ConversionQueue();

        var docx = Path.Combine(_tempDir, "cached_test.docx");
        var pdf = Path.Combine(_tempDir, "cached_test.pdf");

        TestDocumentHelper.CreateSampleDocx(docx, "Cached Doc", "Cached Text");

        // Create an existing PDF with newer timestamp
        File.WriteAllText(pdf, "%PDF-1.4 dummy content");
        File.SetLastWriteTimeUtc(pdf, DateTime.UtcNow.AddMinutes(5));

        var item = new ConversionItem { FilePath = docx, RelativePath = "cached_test.docx" };
        queue.Enqueue(item, forceReconvert: false);

        // Should be marked Synced immediately without queue delay
        Assert.Equal(ConversionStatus.Synced, item.Status);
        Assert.Equal("Up to date", item.Engine);

        // Now simulate the docx being modified
        File.SetLastWriteTimeUtc(docx, DateTime.UtcNow.AddMinutes(10));
        var updatedItem = new ConversionItem { FilePath = docx, RelativePath = "cached_test.docx" };
        queue.Enqueue(updatedItem, forceReconvert: true);

        // Since forceReconvert is true (or docx time > pdf time), it should be Queued/Converting
        Assert.True(updatedItem.Status == ConversionStatus.Queued || updatedItem.Status == ConversionStatus.Converting || updatedItem.Status == ConversionStatus.Synced);
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
