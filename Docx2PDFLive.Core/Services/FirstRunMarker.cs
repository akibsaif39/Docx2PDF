using System.IO;

namespace Docx2PDFLive.Core.Services;

public static class FirstRunMarker
{
    private static readonly string MarkerDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Docx2PDF");

    private static readonly string MarkerFilePath = Path.Combine(MarkerDirectory, ".first_run_completed");

    public static bool HasBeenWelcomed()
    {
        try
        {
            return File.Exists(MarkerFilePath);
        }
        catch
        {
            return false;
        }
    }

    public static void MarkWelcomed()
    {
        try
        {
            if (!Directory.Exists(MarkerDirectory))
            {
                Directory.CreateDirectory(MarkerDirectory);
            }

            File.WriteAllText(MarkerFilePath, $"Welcomed at: {DateTime.UtcNow:O}");
        }
        catch { }
    }
}
