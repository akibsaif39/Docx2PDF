namespace Docx2PDFLive.Core.Models;

public enum ConversionStatus
{
    Queued,
    Converting,
    Synced,
    Error
}

public enum ConversionEngineType
{
    None,
    LibreOffice,
    OpenXmlPdfSharp
}
