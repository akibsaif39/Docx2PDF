namespace Docx2PDFLive.Core.Models;

/// <summary>
/// Persisted user preferences for Docx2PDF startup and folder watching behavior.
/// </summary>
public class AppSettings
{
    public bool RememberLastFolder { get; set; } = true;
    public string? LastWatchedFolder { get; set; }
}
