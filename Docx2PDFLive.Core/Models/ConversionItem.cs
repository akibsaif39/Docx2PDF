namespace Docx2PDFLive.Core.Models;

public class ConversionItem
{
    public string FilePath { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public ConversionStatus Status { get; set; } = ConversionStatus.Pending;
    public string Timestamp { get; set; } = string.Empty;
    public string? Engine { get; set; }
    public string? ErrorMessage { get; set; }
    public int RunId { get; set; }
}
