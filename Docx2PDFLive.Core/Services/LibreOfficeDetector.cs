using System.IO;
using Microsoft.Win32;

namespace Docx2PDFLive.Core.Services;

public static class LibreOfficeDetector
{
    private static readonly string[] SearchLocations =
    [
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe"),
        @"C:\Program Files\LibreOffice\program\soffice.exe",
        @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ];

    public static bool IsAvailable(out string sofficePath)
    {
        foreach (var path in SearchLocations)
        {
            if (File.Exists(path))
            {
                sofficePath = path;
                return true;
            }
        }

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\LibreOffice");
            if (key != null)
            {
                var val = key.GetValue("Path") as string;
                if (!string.IsNullOrEmpty(val))
                {
                    var exe = Path.Combine(val, "program", "soffice.exe");
                    if (File.Exists(exe))
                    {
                        sofficePath = exe;
                        return true;
                    }
                }
            }
        }
        catch { }

        sofficePath = string.Empty;
        return false;
    }
}
