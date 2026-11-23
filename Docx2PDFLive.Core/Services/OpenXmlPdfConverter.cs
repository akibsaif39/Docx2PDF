using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Docx2PDFLive.Core.Services;

public class OpenXmlPdfConverter : IConverter
{
    static OpenXmlPdfConverter()
    {
        try
        {
            PdfSharp.PdfSharpLicense.SetAgplLicense();
        }
        catch { }
    }

    public string Name => "OpenXml+PdfSharp";

    public Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();

        try
        {
            if (!File.Exists(docxPath))
            {
                return Task.FromResult(ConversionResult.Fail(Name, $"File not found: {docxPath}"));
            }

            var outDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            using var pdfDoc = new PdfDocument();
            pdfDoc.Info.Title = Path.GetFileNameWithoutExtension(docxPath);
            pdfDoc.Info.Creator = "Docx2PDF - Created by Akib";

            // Standard letter/A4 page size setup (A4 is 595 x 842 pt)
            const double pageWidth = 595;
            const double pageHeight = 842;
            const double marginLeft = 50;
            const double marginRight = 50;
            const double marginTop = 50;
            const double marginBottom = 50;
            var contentWidth = pageWidth - marginLeft - marginRight;

            PdfPage? currentPage = null;
            XGraphics? gfx = null;
            double currentY = marginTop;

            void EnsurePage()
            {
                if (currentPage == null || currentY > pageHeight - marginBottom)
                {
                    gfx?.Dispose();
                    currentPage = pdfDoc.AddPage();
                    currentPage.Width = XUnit.FromPoint(pageWidth);
                    currentPage.Height = XUnit.FromPoint(pageHeight);
                    gfx = XGraphics.FromPdfPage(currentPage);
                    currentY = marginTop;
                }
            }

            try
            {
                // Attempt standard OpenXml SDK package load
                using var fileStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var wordDoc = WordprocessingDocument.Open(fileStream, false);

                var body = wordDoc.MainDocumentPart?.Document.Body;
                if (body == null)
                {
                    throw new InvalidOperationException("DOCX body is empty or null.");
                }

                // Traverse elements (Paragraphs, Tables, etc.)
                foreach (var element in body.Elements())
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (element is Paragraph para)
                    {
                        RenderParagraph(para, ref currentY, contentWidth, marginLeft, pageHeight, marginBottom, EnsurePage, () => gfx!);
                    }
                    else if (element is Table table)
                    {
                        RenderTable(table, ref currentY, contentWidth, marginLeft, pageHeight, marginBottom, EnsurePage, () => gfx!);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Standard OpenXml read failed ({ex.Message}). Falling back to direct ZIP XML extraction.");

                // Resilient fallback: Direct ZIP XML parser for malformed DOCX files
                // (e.g. unquoted XML attributes in [Content_Types].xml, or backslash paths in ZIP entries like word\document.xml)
                ConvertDirectFromZip(docxPath, pdfDoc, ref currentY, contentWidth, marginLeft, pageHeight, marginBottom, EnsurePage, () => gfx!, cancellationToken);
            }

            gfx?.Dispose();

            // If document was completely empty, at least add one empty page
            if (pdfDoc.PageCount == 0)
            {
                var emptyPage = pdfDoc.AddPage();
                emptyPage.Width = XUnit.FromPoint(pageWidth);
                emptyPage.Height = XUnit.FromPoint(pageHeight);
            }

            // Save PDF
            pdfDoc.Save(pdfPath);
            sw.Stop();

            return Task.FromResult(ConversionResult.Ok(Name, sw.ElapsedMilliseconds));
        }
        catch (Exception ex)
        {
            sw.Stop();
            return Task.FromResult(ConversionResult.Fail(Name, $"OpenXml conversion error: {ex.Message}", sw.ElapsedMilliseconds));
        }
    }

    private static void ConvertDirectFromZip(
        string docxPath,
        PdfDocument pdfDoc,
        ref double currentY,
        double contentWidth,
        double marginLeft,
        double pageHeight,
        double marginBottom,
        Action ensurePage,
        Func<XGraphics> getGfx,
        CancellationToken cancellationToken)
    {
        using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);

        // Find main document entry (handles word/document.xml, word\document.xml, etc.)
        var docEntry = zip.Entries.FirstOrDefault(e =>
            e.FullName.EndsWith("document.xml", StringComparison.OrdinalIgnoreCase));

        if (docEntry == null)
        {
            throw new FileNotFoundException("Could not find document.xml entry inside the DOCX package.");
        }

        using var entryStream = docEntry.Open();
        var xdoc = XDocument.Load(entryStream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        var body = xdoc.Root?.Element(w + "body");
        if (body == null) return;

        foreach (var el in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (el.Name == w + "p")
            {
                var text = string.Concat(el.Descendants(w + "t").Select(t => t.Value));
                RenderPlainText(text, ref currentY, contentWidth, marginLeft, pageHeight, marginBottom, ensurePage, getGfx);
            }
            else if (el.Name == w + "tbl")
            {
                RenderXmlTable(el, w, ref currentY, contentWidth, marginLeft, pageHeight, marginBottom, ensurePage, getGfx);
            }
        }
    }

    private static void RenderPlainText(
        string text,
        ref double currentY,
        double contentWidth,
        double marginLeft,
        double pageHeight,
        double marginBottom,
        Action ensurePage,
        Func<XGraphics> getGfx)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            currentY += 15;
            return;
        }

        ensurePage();
        var gfx = getGfx();
        var font = new XFont("Arial", 11, XFontStyleEx.Regular);
        var brush = XBrushes.Black;

        var words = text.Split(' ');
        var currentLine = "";

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
            var size = gfx.MeasureString(testLine, font);

            if (size.Width > contentWidth && !string.IsNullOrEmpty(currentLine))
            {
                ensurePage();
                gfx = getGfx();
                gfx.DrawString(currentLine, font, brush, new XPoint(marginLeft, currentY), XStringFormats.TopLeft);
                currentY += 15;
                currentLine = word;
            }
            else
            {
                currentLine = testLine;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
        {
            ensurePage();
            gfx = getGfx();
            gfx.DrawString(currentLine, font, brush, new XPoint(marginLeft, currentY), XStringFormats.TopLeft);
            currentY += 17;
        }
    }

    private static void RenderXmlTable(
        XElement tbl,
        XNamespace w,
        ref double currentY,
        double contentWidth,
        double marginLeft,
        double pageHeight,
        double marginBottom,
        Action ensurePage,
        Func<XGraphics> getGfx)
    {
        var rows = tbl.Elements(w + "tr").ToList();
        if (rows.Count == 0) return;

        var font = new XFont("Arial", 10, XFontStyleEx.Regular);
        var pen = new XPen(XColors.Gray, 0.5);

        foreach (var row in rows)
        {
            var cells = row.Elements(w + "tc").ToList();
            if (cells.Count == 0) continue;

            var cellWidth = contentWidth / cells.Count;
            const double rowHeight = 22.0;

            ensurePage();
            var gfx = getGfx();

            for (int i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                var cellX = marginLeft + (i * cellWidth);
                var rect = new XRect(cellX, currentY, cellWidth, rowHeight);
                gfx.DrawRectangle(pen, rect);

                var cellText = string.Concat(cell.Descendants(w + "t").Select(t => t.Value)).Trim();
                if (!string.IsNullOrEmpty(cellText))
                {
                    var textRect = new XRect(cellX + 4, currentY + 3, cellWidth - 8, rowHeight - 6);
                    gfx.DrawString(cellText, font, XBrushes.Black, textRect, XStringFormats.TopLeft);
                }
            }

            currentY += rowHeight;
        }

        currentY += 8;
    }

    private static void RenderParagraph(
        Paragraph para,
        ref double currentY,
        double contentWidth,
        double marginLeft,
        double pageHeight,
        double marginBottom,
        Action ensurePage,
        Func<XGraphics> getGfx)
    {
        // Extract paragraph style / headings
        var isHeading1 = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.Contains("Heading1", StringComparison.OrdinalIgnoreCase) == true;
        var isHeading2 = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.Contains("Heading2", StringComparison.OrdinalIgnoreCase) == true;

        var baseFontSize = isHeading1 ? 18.0 : (isHeading2 ? 14.0 : 11.0);
        var baseBold = isHeading1 || isHeading2;

        // Check paragraph alignment
        var justification = para.ParagraphProperties?.Justification?.Val?.Value;
        var align = XStringFormats.TopLeft;
        if (justification == JustificationValues.Center) align = XStringFormats.TopCenter;
        else if (justification == JustificationValues.Right) align = XStringFormats.TopRight;

        var fullText = para.InnerText;

        if (string.IsNullOrWhiteSpace(fullText))
        {
            // Empty line
            currentY += baseFontSize + 4;
            return;
        }

        ensurePage();
        var gfx = getGfx();

        var fontStyle = baseBold ? XFontStyleEx.Bold : XFontStyleEx.Regular;
        var font = new XFont("Arial", baseFontSize, fontStyle);
        var brush = XBrushes.Black;

        // Word-wrap long lines
        var words = fullText.Split(' ');
        var currentLine = "";

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
            var size = gfx.MeasureString(testLine, font);

            if (size.Width > contentWidth && !string.IsNullOrEmpty(currentLine))
            {
                // Draw current line
                ensurePage();
                gfx = getGfx();
                var drawX = marginLeft;
                if (align == XStringFormats.TopCenter) drawX = marginLeft + (contentWidth / 2);
                else if (align == XStringFormats.TopRight) drawX = marginLeft + contentWidth;

                gfx.DrawString(currentLine, font, brush, new XPoint(drawX, currentY), align);
                currentY += baseFontSize + 4;
                currentLine = word;
            }
            else
            {
                currentLine = testLine;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
        {
            ensurePage();
            gfx = getGfx();
            var drawX = marginLeft;
            if (align == XStringFormats.TopCenter) drawX = marginLeft + (contentWidth / 2);
            else if (align == XStringFormats.TopRight) drawX = marginLeft + contentWidth;

            gfx.DrawString(currentLine, font, brush, new XPoint(drawX, currentY), align);
            currentY += baseFontSize + 6;
        }
    }

    private static void RenderTable(
        Table table,
        ref double currentY,
        double contentWidth,
        double marginLeft,
        double pageHeight,
        double marginBottom,
        Action ensurePage,
        Func<XGraphics> getGfx)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        var font = new XFont("Arial", 10, XFontStyleEx.Regular);
        var pen = new XPen(XColors.Gray, 0.5);

        foreach (var row in rows)
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count == 0) continue;

            var cellWidth = contentWidth / cells.Count;
            const double rowHeight = 22.0;

            ensurePage();
            var gfx = getGfx();

            for (int i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                var cellX = marginLeft + (i * cellWidth);
                var rect = new XRect(cellX, currentY, cellWidth, rowHeight);

                gfx.DrawRectangle(pen, rect);
                var text = cell.InnerText.Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    // Draw text inside cell with margin
                    var textRect = new XRect(cellX + 4, currentY + 3, cellWidth - 8, rowHeight - 6);
                    gfx.DrawString(text, font, XBrushes.Black, textRect, XStringFormats.TopLeft);
                }
            }

            currentY += rowHeight;
        }

        currentY += 8; // Space after table
    }
}
