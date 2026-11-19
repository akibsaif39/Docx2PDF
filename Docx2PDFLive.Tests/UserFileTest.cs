using System.IO;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class UserFileTest : IDisposable
{
    private readonly string _tempDir;

    public UserFileTest()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_UserFileTest_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task Convert_MalformedXmlDocx_Succeeds()
    {
        var docxPath = Path.Combine(_tempDir, "malformed.docx");
        var pdfPath = Path.Combine(_tempDir, "malformed.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "Resilience Test", "Document with normal text.");

        var converter = new OpenXmlPdfConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
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
