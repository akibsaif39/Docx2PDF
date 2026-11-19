using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Docx2PDFLive.Core.Services;

public class OpenXmlPdfConverter : IConverter
{
    public string Name => "Built-in OpenXml Engine";

    public async Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var sw = Stopwatch.StartNew();
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                var paragraphs = ExtractDocumentParagraphs(docxPath);

                using var pdf = new PdfDocument();
                pdf.Info.Title = Path.GetFileNameWithoutExtension(docxPath);
                pdf.Info.Creator = "Docx2PDF";

                var page = pdf.AddPage();
                var gfx = XGraphics.FromPdfPage(page);
                var fontRegular = new XFont("Segoe UI", 11);
                var fontHeading = new XFont("Segoe UI", 14, XFontStyleEx.Bold);
                var brush = XBrushes.Black;

                double xMargin = 50;
                double yMargin = 50;
                double currentY = yMargin;
                double maxPageY = page.Height - yMargin;

                foreach (var (text, isHeading) in paragraphs)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var font = isHeading ? fontHeading : fontRegular;
                    var lineSpacing = isHeading ? 22 : 16;

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        currentY += 10;
                        continue;
                    }

                    var lines = WrapText(gfx, text, font, page.Width - (xMargin * 2));

                    foreach (var line in lines)
                    {
                        if (currentY + lineSpacing > maxPageY)
                        {
                            page = pdf.AddPage();
                            gfx.Dispose();
                            gfx = XGraphics.FromPdfPage(page);
                            currentY = yMargin;
                        }

                        gfx.DrawString(line, font, brush, new XPoint(xMargin, currentY));
                        currentY += lineSpacing;
                    }

                    currentY += isHeading ? 8 : 4;
                }

                gfx.Dispose();

                var outDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                pdf.Save(pdfPath);
                sw.Stop();
                return ConversionResult.Ok(Name, sw.ElapsedMilliseconds);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ConversionResult.Fail(Name, ex.Message, sw.ElapsedMilliseconds);
            }
        }, cancellationToken);
    }

    private static List<(string Text, bool IsHeading)> ExtractDocumentParagraphs(string docxPath)
    {
        var paragraphs = new List<(string, bool)>();

        using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);

        ZipArchiveEntry? entry = null;
        foreach (var e in zip.Entries)
        {
            var clean = e.FullName.Replace('\\', '/');
            if (clean.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
            {
                entry = e;
                break;
            }
        }

        if (entry == null) return paragraphs;

        using var reader = new StreamReader(entry.Open());
        var doc = XDocument.Load(reader);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        foreach (var p in doc.Descendants(w + "p"))
        {
            var pStyle = p.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val")?.Value ?? string.Empty;
            var isHeading = pStyle.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) || pStyle.Equals("Title", StringComparison.OrdinalIgnoreCase);

            var sb = new System.Text.StringBuilder();
            foreach (var t in p.Descendants(w + "t"))
            {
                sb.Append(t.Value);
            }

            var text = sb.ToString();
            paragraphs.Add((text, isHeading));
        }

        return paragraphs;
    }

    private static List<string> WrapText(XGraphics gfx, string text, XFont font, double maxWidth)
    {
        var lines = new List<string>();
        var words = text.Split(' ');
        var currentLine = string.Empty;

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
            var size = gfx.MeasureString(testLine, font);

            if (size.Width <= maxWidth)
            {
                currentLine = testLine;
            }
            else
            {
                if (!string.IsNullOrEmpty(currentLine))
                {
                    lines.Add(currentLine);
                }
                currentLine = word;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
        {
            lines.Add(currentLine);
        }

        return lines;
    }
}
