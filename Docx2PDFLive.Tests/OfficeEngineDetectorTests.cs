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
}
