using System.IO;
using System.IO.Compression;
using System.Text;

namespace Docx2PDFLive.Tests;

public static class TestDocumentHelper
{
    public static void CreateSampleDocx(string filePath, string headingText, string bodyText)
    {
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

        var contentTypes = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
    <Default Extension=""xml"" ContentType=""application/xml""/>
    <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
</Types>";
        AddZipEntry(zip, "[Content_Types].xml", contentTypes);

        var docXml = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:body>
        <w:p>
            <w:pPr><w:pStyle w:val=""Heading1""/></w:pPr>
            <w:r><w:t>{headingText}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>{bodyText}</w:t></w:r>
        </w:p>
    </w:body>
</w:document>";
        AddZipEntry(zip, "word/document.xml", docXml);
    }

    private static void AddZipEntry(ZipArchive zip, string entryName, string content)
    {
        var entry = zip.CreateEntry(entryName, CompressionLevel.Fastest);
        using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
        writer.Write(content);
    }
}
