using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Docx2PDFLive.Core.Services;

public static class LibreOfficeDetector
{
    public const string OfficialDownloadPageUrl = "https://www.libreoffice.org/download/download-libreoffice/";
    public const string DirectInstallerUrl = "https://download.documentfoundation.org/libreoffice/stable/24.8.3/win/x86_64/LibreOffice_24.8.3_Win_x86-64.msi";

    private static string? _cachedPath;

    public static bool IsAvailable(out string sofficePath)
    {
        if (!string.IsNullOrEmpty(_cachedPath) && File.Exists(_cachedPath))
        {
            sofficePath = _cachedPath;
            return true;
        }

        var found = FindSofficeExecutable();
        if (!string.IsNullOrEmpty(found) && File.Exists(found))
        {
            _cachedPath = found;
            sofficePath = found;
            return true;
        }

        sofficePath = string.Empty;
        return false;
    }

    public static string? FindSofficeExecutable()
    {
        // 1. Check known default file locations
        string[] standardPaths =
        [
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "LibreOffice", "program", "soffice.exe"),
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ];

        foreach (var path in standardPaths)
        {
            if (File.Exists(path)) return path;
        }

        // 2. Check Registry: App Paths
        try
        {
            using var appPathKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe");
            var appPath = appPathKey?.GetValue("") as string;
            if (!string.IsNullOrEmpty(appPath) && File.Exists(appPath))
            {
                return appPath;
            }
        }
        catch { /* Ignore registry permission errors */ }

        // 3. Check Registry: LibreOffice installation key
        string[] regKeys =
        [
            @"SOFTWARE\LibreOffice\LibreOffice",
            @"SOFTWARE\WOW6432Node\LibreOffice\LibreOffice"
        ];

        foreach (var keyName in regKeys)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(keyName);
                var val = key?.GetValue("Path") as string;
                if (!string.IsNullOrEmpty(val))
                {
                    var full = Path.Combine(val, "program", "soffice.exe");
                    if (File.Exists(full)) return full;
                    if (File.Exists(val)) return val;
                }
            }
            catch { }
        }

        // 4. Check PATH environment variable
        var envPath = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(envPath))
        {
            var parts = envPath.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                try
                {
                    var candidate = Path.Combine(part.Trim('"', ' '), "soffice.exe");
                    if (File.Exists(candidate)) return candidate;
                }
                catch { }
            }
        }

        return null;
    }

    public static void ResetCache()
    {
        _cachedPath = null;
    }

    public static void OpenDownloadPage()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = OfficialDownloadPageUrl,
                UseShellExecute = true
            });
        }
        catch { }
    }
}
