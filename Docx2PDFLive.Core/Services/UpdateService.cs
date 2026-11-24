using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Docx2PDFLive.Core.Models;

namespace Docx2PDFLive.Core.Services;

public class UpdateService
{
    public const string DefaultManifestUrl = "https://raw.githubusercontent.com/akibsaif39/Docx2PDF/main/version.json";
    public const string DefaultGitHubReleasesUrl = "https://api.github.com/repos/akibsaif39/Docx2PDF/releases/latest";

    private readonly HttpClient _httpClient;

    public UpdateService(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        if (!_httpClient.DefaultRequestHeaders.Contains("User-Agent"))
        {
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Docx2PDF-AutoUpdater");
        }
    }

    /// <summary>
    /// Checks whether an online update is available.
    /// Gracefully returns an Offline / NoUpdate report if network is down or times out.
    /// </summary>
    public async Task<UpdateInfo> CheckForUpdateAsync(
        Version currentVersion,
        string manifestUrl = DefaultManifestUrl,
        int timeoutSeconds = 3,
        CancellationToken cancellationToken = default)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

        // 1. Try checking the fast, unthrottled version.json manifest
        try
        {
            using var response = await _httpClient.GetAsync(manifestUrl, cts.Token).ConfigureAwait(false);
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync(cts.Token).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var latestStr = root.TryGetProperty("latestVersion", out var lv) ? lv.GetString() : null;
                var mandatory = !root.TryGetProperty("mandatory", out var m) || m.GetBoolean();
                var notes = root.TryGetProperty("releaseNotes", out var n) ? n.GetString() : string.Empty;
                var downloadUrl = root.TryGetProperty("downloadUrl", out var du) ? du.GetString() : null;

                if (!string.IsNullOrEmpty(latestStr) && CleanVersion(latestStr) > currentVersion)
                {
                    return new UpdateInfo
                    {
                        IsUpdateAvailable = true,
                        IsMandatory = mandatory,
                        LatestVersion = latestStr,
                        CurrentVersion = currentVersion.ToString(),
                        ReleaseNotes = notes ?? string.Empty,
                        DownloadUrl = downloadUrl ?? string.Empty
                    };
                }

                return UpdateInfo.NoUpdate(currentVersion.ToString());
            }
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // Timeout reached: gracefully treat as offline mode
            return UpdateInfo.Offline(currentVersion.ToString());
        }
        catch (HttpRequestException)
        {
            // Network unreachable / DNS failure: gracefully treat as offline mode
            return UpdateInfo.Offline(currentVersion.ToString());
        }
        catch
        {
            // Manifest parsing issue: fall through to GitHub Releases API fallback below
        }

        // 2. Fallback check: Query GitHub Releases API directly
        try
        {
            using var ghCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            ghCts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

            using var ghResp = await _httpClient.GetAsync(DefaultGitHubReleasesUrl, ghCts.Token).ConfigureAwait(false);
            if (ghResp.IsSuccessStatusCode)
            {
                var ghJson = await ghResp.Content.ReadAsStringAsync(ghCts.Token).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(ghJson);
                var root = doc.RootElement;

                var tagName = root.TryGetProperty("tag_name", out var tn) ? tn.GetString() : null;
                var notes = root.TryGetProperty("body", out var b) ? b.GetString() : string.Empty;

                string? downloadUrl = null;
                if (root.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
                {
                    foreach (var asset in assets.EnumerateArray())
                    {
                        var name = asset.TryGetProperty("name", out var an) ? an.GetString() : string.Empty;
                        if (string.Equals(name, "Docx2PDF_Install.exe", StringComparison.OrdinalIgnoreCase))
                        {
                            downloadUrl = asset.TryGetProperty("browser_download_url", out var bdu) ? bdu.GetString() : null;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(tagName) && CleanVersion(tagName) > currentVersion && !string.IsNullOrEmpty(downloadUrl))
                {
                    return new UpdateInfo
                    {
                        IsUpdateAvailable = true,
                        IsMandatory = true,
                        LatestVersion = tagName.TrimStart('v', 'V'),
                        CurrentVersion = currentVersion.ToString(),
                        ReleaseNotes = notes ?? string.Empty,
                        DownloadUrl = downloadUrl
                    };
                }

                return UpdateInfo.NoUpdate(currentVersion.ToString());
            }
        }
        catch
        {
            // Network unavailable
            return UpdateInfo.Offline(currentVersion.ToString());
        }

        return UpdateInfo.NoUpdate(currentVersion.ToString());
    }

    /// <summary>
    /// Downloads the update installer with real-time progress streaming.
    /// </summary>
    public async Task DownloadUpdateAsync(
        string downloadUrl,
        string destinationFilePath,
        IProgress<DownloadProgressReport>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var tempPartPath = destinationFilePath + ".part";
        if (File.Exists(tempPartPath))
        {
            try { File.Delete(tempPartPath); } catch { }
        }

        var dir = Path.GetDirectoryName(destinationFilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength ?? -1L;

        using (var contentStream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false))
        using (var fileStream = new FileStream(tempPartPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
        {
            var buffer = new byte[81920];
            long totalReadBytes = 0;
            var sw = Stopwatch.StartNew();
            long lastSampleBytes = 0;
            var lastSampleTime = sw.Elapsed;

            int bytesRead;
            while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await fileStream.WriteAsync(buffer, 0, bytesRead, cancellationToken).ConfigureAwait(false);
                totalReadBytes += bytesRead;

                var elapsed = sw.Elapsed;
                var deltaSec = (elapsed - lastSampleTime).TotalSeconds;
                double speed = 0;
                if (deltaSec >= 0.25)
                {
                    speed = (totalReadBytes - lastSampleBytes) / deltaSec;
                    lastSampleBytes = totalReadBytes;
                    lastSampleTime = elapsed;
                }

                double percentage = totalBytes > 0 ? ((double)totalReadBytes / totalBytes) * 100.0 : 0.0;

                progress?.Report(new DownloadProgressReport
                {
                    BytesReceived = totalReadBytes,
                    TotalBytes = totalBytes,
                    ProgressPercentage = percentage,
                    SpeedBytesPerSecond = speed
                });
            }
        }

        if (File.Exists(destinationFilePath))
        {
            try { File.Delete(destinationFilePath); } catch { }
        }

        File.Move(tempPartPath, destinationFilePath);
    }

    private static Version CleanVersion(string vStr)
    {
        var clean = vStr.Trim().TrimStart('v', 'V');
        // Handle semantic suffixes like "-preview" if any
        var dashIdx = clean.IndexOf('-');
        if (dashIdx >= 0) clean = clean.Substring(0, dashIdx);

        if (Version.TryParse(clean, out var v))
        {
            return v;
        }

        // Try padding e.g. "1.0" -> "1.0.0"
        var parts = clean.Split('.');
        if (parts.Length == 1) clean += ".0.0";
        else if (parts.Length == 2) clean += ".0";

        if (Version.TryParse(clean, out v))
        {
            return v;
        }

        return new Version(1, 0, 0);
    }
}
