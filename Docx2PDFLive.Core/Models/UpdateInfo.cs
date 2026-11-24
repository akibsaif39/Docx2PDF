using System;

namespace Docx2PDFLive.Core.Models;

public class UpdateInfo
{
    public bool IsUpdateAvailable { get; set; }
    public bool IsMandatory { get; set; } = true;
    public string LatestVersion { get; set; } = string.Empty;
    public string CurrentVersion { get; set; } = string.Empty;
    public string ReleaseNotes { get; set; } = string.Empty;
    public string DownloadUrl { get; set; } = string.Empty;
    public bool IsOffline { get; set; }
    public string? ErrorMessage { get; set; }

    public static UpdateInfo NoUpdate(string currentVersion) => new()
    {
        IsUpdateAvailable = false,
        CurrentVersion = currentVersion,
        LatestVersion = currentVersion
    };

    public static UpdateInfo Offline(string currentVersion) => new()
    {
        IsUpdateAvailable = false,
        IsOffline = true,
        CurrentVersion = currentVersion
    };
}

public class DownloadProgressReport
{
    public long BytesReceived { get; set; }
    public long TotalBytes { get; set; }
    public double ProgressPercentage { get; set; }
    public double SpeedBytesPerSecond { get; set; }
}
