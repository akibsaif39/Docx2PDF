using System.IO;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class ConversionTests : IDisposable
{
    private readonly string _tempDir;

    public ConversionTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_Tests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task OpenXmlPdfConverter_ConvertsDocxToPdf_Successfully()
    {
        var docxPath = Path.Combine(_tempDir, "sample_openxml.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_openxml.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "OpenXML Test Doc", "Testing OpenXml and PdfSharp converter.");

        var converter = new OpenXmlPdfConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
        var fileInfo = new FileInfo(pdfPath);
        Assert.True(fileInfo.Length > 500, "PDF file size should be > 500 bytes.");
    }

    [Fact]
    public async Task LibreOfficeConverter_ConvertsDocxToPdf_Successfully()
    {
        var docxPath = Path.Combine(_tempDir, "sample_lo.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_lo.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "LibreOffice Test Doc", "Testing Headless LibreOffice conversion.");

        var converter = new LibreOfficeConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created by LibreOffice.");
        var fileInfo = new FileInfo(pdfPath);
        Assert.True(fileInfo.Length > 1000, "PDF file size should be > 1000 bytes.");
    }

    [Fact]
    public async Task WarmWordConverter_ConvertsDocxToPdf_Successfully()
    {
        if (!OfficeEngineDetector.IsWordInstalled(out _))
        {
            return; // Skip on systems without Microsoft Word
        }

        var docxPath = Path.Combine(_tempDir, "sample_word.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_word.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "Word Test Doc", "Testing high-speed Microsoft Word warm conversion.");

        using var converter = new WarmWordConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created by Word.");
        var fileInfo = new FileInfo(pdfPath);
        Assert.True(fileInfo.Length > 1000, "PDF file size should be > 1000 bytes.");
    }

    [Fact]
    public async Task CompositeConverter_ConvertsDocxToPdf_Successfully()
    {
        var docxPath = Path.Combine(_tempDir, "sample_composite.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_composite.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "Composite Test Doc", "Testing high-fidelity composite converter.");

        using var converter = new CompositeConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
        Assert.True(result.EngineUsed.Contains("Word") || result.EngineUsed.Contains("LibreOffice"));
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
