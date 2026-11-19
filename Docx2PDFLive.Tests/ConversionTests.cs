using System.IO;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class ConversionTests : IDisposable
{
    private readonly string _tempDir;

    public ConversionTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDFLive_ConvTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task OpenXmlPdfConverter_ConvertsDocxToPdf_Successfully()
    {
        var docxPath = Path.Combine(_tempDir, "sample_openxml.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_openxml.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "OpenXML Test Title", "This is converted by OpenXml+PdfSharp.");

        var converter = new OpenXmlPdfConverter();
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Conversion failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
        Assert.True(new FileInfo(pdfPath).Length > 0, "PDF file is empty.");
    }

    [Fact]
    public async Task LibreOfficeConverter_ConvertsDocxToPdf_Successfully()
    {
        if (!LibreOfficeDetector.IsAvailable(out var loPath))
        {
            return;
        }

        var docxPath = Path.Combine(_tempDir, "sample_lo.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_lo.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "LibreOffice Test Title", "This is converted by LibreOffice Headless.");

        var converter = new LibreOfficeConverter(loPath);
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"LibreOffice failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF output missing.");
    }

    [Fact]
    public async Task CompositeConverter_ConvertsDocxToPdf_Successfully()
    {
        var docxPath = Path.Combine(_tempDir, "sample_composite.docx");
        var pdfPath = Path.Combine(_tempDir, "sample_composite.pdf");

        TestDocumentHelper.CreateSampleDocx(docxPath, "Composite Test Title", "This is converted by CompositeConverter.");

        var converter = new CompositeConverter(preferLibreOffice: true);
        var result = await converter.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Composite converter failed: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "Composite PDF output missing.");
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
