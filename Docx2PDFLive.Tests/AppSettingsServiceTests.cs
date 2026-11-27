using System.IO;
using Docx2PDFLive.Core.Models;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class AppSettingsServiceTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _settingsFile;

    public AppSettingsServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "Docx2PDF_SettingsTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _settingsFile = Path.Combine(_tempDir, "settings.json");
    }

    [Fact]
    public void Load_WhenFileDoesNotExist_ReturnsDefaultRememberTrueAndNullFolder()
    {
        var service = new AppSettingsService(_settingsFile);
        var settings = service.Load();

        Assert.True(settings.RememberLastFolder);
        Assert.Null(settings.LastWatchedFolder);
    }

    [Fact]
    public void SaveAndLoad_WhenRememberEnabled_PersistsLastWatchedFolder()
    {
        var service = new AppSettingsService(_settingsFile);
        var folder = Path.Combine(_tempDir, "Contracts");
        Directory.CreateDirectory(folder);

        service.Save(new AppSettings
        {
            RememberLastFolder = true,
            LastWatchedFolder = folder
        });

        var reloaded = service.Load();
        Assert.True(reloaded.RememberLastFolder);
        Assert.Equal(folder, reloaded.LastWatchedFolder);

        var resolution = AppSettingsService.ResolveStartupFolder(reloaded);
        Assert.Equal(StartupFolderAction.AutoResumeFolder, resolution.Action);
        Assert.Equal(folder, resolution.FolderPath);
    }

    [Fact]
    public void Save_WhenRememberDisabled_ClearsLastWatchedFolder()
    {
        var service = new AppSettingsService(_settingsFile);
        service.Save(new AppSettings
        {
            RememberLastFolder = true,
            LastWatchedFolder = _tempDir
        });

        // User unchecks Remember folder on launch
        service.Save(new AppSettings
        {
            RememberLastFolder = false,
            LastWatchedFolder = _tempDir
        });

        var reloaded = service.Load();
        Assert.False(reloaded.RememberLastFolder);
        Assert.Null(reloaded.LastWatchedFolder);

        var resolution = AppSettingsService.ResolveStartupFolder(reloaded);
        Assert.Equal(StartupFolderAction.PromptForFolder, resolution.Action);
    }

    [Fact]
    public void Load_WhenJsonIsCorrupt_FallsBackGracefullyToDefaults()
    {
        File.WriteAllText(_settingsFile, "{ not valid json !!!");
        var service = new AppSettingsService(_settingsFile);

        var settings = service.Load();
        Assert.True(settings.RememberLastFolder);
        Assert.Null(settings.LastWatchedFolder);
    }

    [Fact]
    public void ResolveStartupFolder_WhenRememberedFolderMissing_ReturnsMissingFolderPrompt()
    {
        var missingPath = Path.Combine(_tempDir, "DeletedOrUnpluggedDrive");
        var settings = new AppSettings
        {
            RememberLastFolder = true,
            LastWatchedFolder = missingPath
        };

        var resolution = AppSettingsService.ResolveStartupFolder(settings);
        Assert.Equal(StartupFolderAction.MissingRememberedFolderPrompt, resolution.Action);
        Assert.Equal(missingPath, resolution.FolderPath);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch { }
    }
}
