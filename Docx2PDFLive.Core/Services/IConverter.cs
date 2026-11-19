namespace Docx2PDFLive.Core.Services;

public interface IConverter
{
    string Name { get; }
    Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default);
}
