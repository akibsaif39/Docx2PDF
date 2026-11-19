using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class LibreOfficeDetectorTests
{
    [Fact]
    public void LibreOfficeDetector_FindsInstalledLibreOffice()
    {
        var found = LibreOfficeDetector.IsAvailable(out var path);
        if (found)
        {
            Assert.True(System.IO.File.Exists(path), "Path returned by detector must exist on disk.");
        }
    }
}
