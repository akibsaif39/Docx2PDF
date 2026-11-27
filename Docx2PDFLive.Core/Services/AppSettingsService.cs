using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using Docx2PDFLive.Core.Models;

namespace Docx2PDFLive.Core.Services;

public enum StartupFolderAction
{
    PromptForFolder,
    AutoResumeFolder,
    MissingRememberedFolderPrompt
}

public readonly record struct StartupFolderResolution(StartupFolderAction Action, string? FolderPath);

/// <summary>
/// Loads and persists user settings in %LocalAppData%\Docx2PDFLive\settings.json.
/// Handles I/O and malformed JSON defensively so startup never fails.
/// </summary>
public class AppSettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _settingsFilePath;
    private readonly object _lock = new();

    public static string DefaultSettingsFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Docx2PDFLive",
        "settings.json");

    public string SettingsFilePath => _settingsFilePath;

    public AppSettingsService(string? settingsFilePath = null)
    {
        _settingsFilePath = string.IsNullOrWhiteSpace(settingsFilePath)
            ? DefaultSettingsFilePath
            : settingsFilePath;
    }

    public AppSettings Load()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(_settingsFilePath))
                {
                    return new AppSettings();
                }

                var json = File.ReadAllText(_settingsFilePath, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return new AppSettings();
                }

                var parsed = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                if (parsed == null)
                {
                    return new AppSettings();
                }

                if (!parsed.RememberLastFolder)
                {
                    parsed.LastWatchedFolder = null;
                }

                return parsed;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load settings from {_settingsFilePath}: {ex.Message}");
                return new AppSettings();
            }
        }
    }

    public void Save(AppSettings settings)
    {
        if (settings == null) throw new ArgumentNullException(nameof(settings));

        lock (_lock)
        {
            try
            {
                var dir = Path.GetDirectoryName(_settingsFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var normalized = new AppSettings
                {
                    RememberLastFolder = settings.RememberLastFolder,
                    LastWatchedFolder = settings.RememberLastFolder && !string.IsNullOrWhiteSpace(settings.LastWatchedFolder)
                        ? settings.LastWatchedFolder
                        : null
                };

                var json = JsonSerializer.Serialize(normalized, JsonOptions);
                File.WriteAllText(_settingsFilePath, json + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to save settings to {_settingsFilePath}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Determines startup folder behavior from persisted settings and disk state.
    /// </summary>
    public static StartupFolderResolution ResolveStartupFolder(AppSettings settings, Func<string, bool>? directoryExists = null)
    {
        if (settings == null || !settings.RememberLastFolder || string.IsNullOrWhiteSpace(settings.LastWatchedFolder))
        {
            return new StartupFolderResolution(StartupFolderAction.PromptForFolder, null);
        }

        var existsCheck = directoryExists ?? Directory.Exists;
        if (existsCheck(settings.LastWatchedFolder))
        {
            return new StartupFolderResolution(StartupFolderAction.AutoResumeFolder, settings.LastWatchedFolder);
        }

        return new StartupFolderResolution(StartupFolderAction.MissingRememberedFolderPrompt, settings.LastWatchedFolder);
    }
}
