using System.IO;
using System.IO.Compression;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class UserFileTest
{
    [Fact]
    public async Task Convert_MalformedXmlDocx_Succeeds()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDF_Malformed_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var docxPath = Path.Combine(tempDir, "malformed_test.docx");
        var pdfPath = Path.Combine(tempDir, "malformed_test.pdf");

        // Construct a docx package with unquoted attributes and backslash entry name
        using (var fs = new FileStream(docxPath, FileMode.Create))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // [Content_Types].xml with unquoted attributes
            var ct = zip.CreateEntry("[Content_Types].xml");
            using (var writer = new StreamWriter(ct.Open()))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default ContentType=application/vnd.openxmlformats-package.relationships+xml Extension=rels/><Default ContentType=application/xml Extension=xml/><Override ContentType=application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml PartName=/word/document.xml/></Types>");
            }

            // word\document.xml with backslash
            var doc = zip.CreateEntry("word\\document.xml");
            using (var writer = new StreamWriter(doc.Open()))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Testing resilient conversion of malformed package.</w:t></w:r></w:p></w:body></w:document>");
            }
        }

        using var composite = new CompositeConverter();
        var result = await composite.ConvertAsync(docxPath, pdfPath);

        Assert.True(result.Success, $"Error: {result.ErrorMessage}");
        Assert.True(File.Exists(pdfPath), "PDF file was not created.");
        Assert.True(new FileInfo(pdfPath).Length > 500, "PDF size must be > 500 bytes.");

        try { Directory.Delete(tempDir, true); } catch { }
    }
}
