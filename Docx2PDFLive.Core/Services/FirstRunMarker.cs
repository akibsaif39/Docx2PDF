using System.IO;

namespace Docx2PDFLive.Core.Services;

public static class FirstRunMarker
{
    public static string DirectoryPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Docx2PDFLive");

    public static string FlagFilePath =>
        Path.Combine(DirectoryPath, "welcomed.flag");

    public static bool HasBeenWelcomed()
    {
        try
        {
            return File.Exists(FlagFilePath);
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
            if (!Directory.Exists(DirectoryPath))
            {
                Directory.CreateDirectory(DirectoryPath);
            }
            File.WriteAllText(FlagFilePath, $"Welcomed at {DateTime.UtcNow:O} - Docx2PDF created by Akib");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to write welcome flag: {ex.Message}");
        }
    }

    public static void ResetWelcome()
    {
        try
        {
            if (File.Exists(FlagFilePath))
            {
                File.Delete(FlagFilePath);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to delete welcome flag: {ex.Message}");
        }
    }
}
