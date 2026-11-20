using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Docx2PDFLive.Core.Services;

[Flags]
public enum OfficeEngineType
{
    None = 0,
    MicrosoftWord = 1,
    LibreOffice = 2,
    Both = MicrosoftWord | LibreOffice
}

/// <summary>
/// Detects installed high-fidelity office conversion engines (Microsoft Word and LibreOffice).
/// </summary>
public static class OfficeEngineDetector
{
    private static bool? _isWordAvailable;
    private static string? _wordPath;

    public static OfficeEngineType GetAvailableEngines()
    {
        var engines = OfficeEngineType.None;

        if (IsWordInstalled(out _))
        {
            engines |= OfficeEngineType.MicrosoftWord;
        }

        if (IsLibreOfficeInstalled(out _))
        {
            engines |= OfficeEngineType.LibreOffice;
        }

        return engines;
    }

    public static OfficeEngineType GetFastestAvailableEngine()
    {
        // Microsoft Word COM in warm mode converts documents in ~150ms-250ms with 100% native Word layout.
        // It is strictly prioritized whenever available.
        if (IsWordInstalled(out _))
        {
            return OfficeEngineType.MicrosoftWord;
        }

        if (IsLibreOfficeInstalled(out _))
        {
            return OfficeEngineType.LibreOffice;
        }

        return OfficeEngineType.None;
    }

    public static bool IsWordInstalled(out string? wordPath)
    {
        if (_isWordAvailable.HasValue)
        {
            wordPath = _wordPath;
            return _isWordAvailable.Value;
        }

        // 1. Check COM ProgID registration
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType != null)
            {
                _isWordAvailable = true;
                _wordPath = FindWordExecutable();
                wordPath = _wordPath;
                return true;
            }
        }
        catch { }

        // 2. Check executable on disk if COM ProgID check was inconclusive
        var path = FindWordExecutable();
        if (!string.IsNullOrEmpty(path) && File.Exists(path))
        {
            _isWordAvailable = true;
            _wordPath = path;
            wordPath = path;
            return true;
        }

        _isWordAvailable = false;
        _wordPath = null;
        wordPath = null;
        return false;
    }

    public static bool IsLibreOfficeInstalled(out string? sofficePath)
    {
        bool available = LibreOfficeDetector.IsAvailable(out var path);
        sofficePath = available ? path : null;
        return available;
    }

    public static void ResetCache()
    {
        _isWordAvailable = null;
        _wordPath = null;
        LibreOfficeDetector.ResetCache();
    }

    private static string? FindWordExecutable()
    {
        // 1. Check Registry App Paths
        string[] registryRoots = [
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
        ];

        foreach (var regKey in registryRoots)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(regKey);
                var val = key?.GetValue("") as string;
                if (!string.IsNullOrEmpty(val) && File.Exists(val))
                {
                    return val;
                }
            }
            catch { }

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(regKey);
                var val = key?.GetValue("") as string;
                if (!string.IsNullOrEmpty(val) && File.Exists(val))
                {
                    return val;
                }
            }
            catch { }
        }

        // 2. Check standard installation directories
        string[] standardPaths = [
            @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office14\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE"
        ];

        foreach (var p in standardPaths)
        {
            if (File.Exists(p)) return p;
        }

        return null;
    }
}
