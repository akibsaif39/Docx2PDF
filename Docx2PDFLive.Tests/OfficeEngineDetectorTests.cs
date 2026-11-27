using System.Diagnostics;
using Docx2PDFLive.Core.Interop;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class OfficeEngineDetectorTests
{
    [Fact]
    public void OfficeEngineDetector_DetectsAtLeastOneHighFidelityEngine()
    {
        var engines = OfficeEngineDetector.GetAvailableEngines();
        Assert.NotEqual(OfficeEngineType.None, engines);

        var fastest = OfficeEngineDetector.GetFastestAvailableEngine();
        Assert.NotEqual(OfficeEngineType.None, fastest);
    }

    [Fact]
    public void JobObjectManager_InitializesSuccessfully()
    {
        JobObjectManager.InitializeGlobalJobObject();
        Assert.True(JobObjectManager.IsInitialized);
    }

    [Fact]
    public void JobObjectManager_DoesNotAssignParentOrUnassignedChildren_OnlyAssignedWorkers()
    {
        JobObjectManager.InitializeGlobalJobObject();
        Assert.True(JobObjectManager.IsInitialized);

        using var current = Process.GetCurrentProcess();
        Assert.False(JobObjectManager.IsProcessAssigned(current));

        using var assignedWorker = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -Command Start-Sleep 10",
            UseShellExecute = false,
            CreateNoWindow = true
        });
        using var unassignedUpdater = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -Command Start-Sleep 10",
            UseShellExecute = false,
            CreateNoWindow = true
        });

        Assert.NotNull(assignedWorker);
        Assert.NotNull(unassignedUpdater);

        try
        {
            Assert.True(JobObjectManager.AssignProcess(assignedWorker!));
            Assert.True(JobObjectManager.IsProcessAssigned(assignedWorker!));
            Assert.False(JobObjectManager.IsProcessAssigned(unassignedUpdater!));
        }
        finally
        {
            try { assignedWorker?.Kill(true); } catch { }
            try { unassignedUpdater?.Kill(true); } catch { }
        }
    }
}
