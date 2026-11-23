using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Docx2PDFLive.Tests;

public static class TestDocumentHelper
{
    public static string CreateSampleDocx(string filePath, string heading = "Test Heading", string bodyText = "This is a test paragraph for Docx2PDF Live conversion.")
    {
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = mainPart.Document.AppendChild(new Body());

        // Heading paragraph
        var headingPara = body.AppendChild(new Paragraph());
        var headingRun = headingPara.AppendChild(new Run());
        headingRun.AppendChild(new Text(heading));

        // Body paragraph
        var bodyPara = body.AppendChild(new Paragraph());
        var bodyRun = bodyPara.AppendChild(new Run());
        bodyRun.AppendChild(new Text(bodyText));

        // Add a table
        var table = body.AppendChild(new Table());
        for (int r = 0; r < 2; r++)
        {
            var row = table.AppendChild(new TableRow());
            for (int c = 0; c < 2; c++)
            {
                var cell = row.AppendChild(new TableCell());
                var cellPara = cell.AppendChild(new Paragraph());
                var cellRun = cellPara.AppendChild(new Run());
                cellRun.AppendChild(new Text($"Cell {r},{c}"));
            }
        }

        mainPart.Document.Save();
        return filePath;
    }
}
