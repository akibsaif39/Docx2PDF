using System.IO;

namespace Docx2PDFLive.Core.Services;

public class CompositeConverter : IConverter
{
    private readonly LibreOfficeConverter? _libreOfficeConverter;
    private readonly OpenXmlPdfConverter _openXmlPdfConverter;
    private readonly bool _preferLibreOffice;

    public string Name => "Smart Composite Converter";

    public CompositeConverter(bool preferLibreOffice = true)
    {
        _preferLibreOffice = preferLibreOffice;
        _openXmlPdfConverter = new OpenXmlPdfConverter();

        if (LibreOfficeDetector.IsAvailable(out var loPath))
        {
            _libreOfficeConverter = new LibreOfficeConverter(loPath);
        }
    }

    public async Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var hasComplexElements = HasComplexFormatting(docxPath);
        var useLibreOffice = _libreOfficeConverter != null && (_preferLibreOffice || hasComplexElements);

        if (useLibreOffice && _libreOfficeConverter != null)
        {
            try
            {
                var loResult = await _libreOfficeConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
                if (loResult.Success && File.Exists(pdfPath))
                {
                    return loResult;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch { }
        }

        cancellationToken.ThrowIfCancellationRequested();

        return await _openXmlPdfConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
    }

    public static bool HasComplexFormatting(string docxPath)
    {
        try
        {
            if (!File.Exists(docxPath)) return false;
            using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read);
            var docEntry = archive.GetEntry("word/document.xml");
            if (docEntry == null) return false;

            using var reader = new StreamReader(docEntry.Open());
            var xml = reader.ReadToEnd();
            return xml.Contains("<w:tbl") || xml.Contains("<w:drawing") || xml.Contains("<w:pict");
        }
        catch
        {
            return false;
        }
    }
}
