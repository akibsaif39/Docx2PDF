using System.IO;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class LibreOfficeDetectorTests
{
    [Fact]
    public void LibreOfficeDetector_FindsInstalledLibreOffice()
    {
        var isAvailable = LibreOfficeDetector.IsAvailable(out var path);

        // On this Windows system, LibreOffice 26.2 is installed at C:\Program Files\LibreOffice\program\soffice.exe
        Assert.True(isAvailable, "LibreOffice should be detected on this system.");
        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.True(File.Exists(path), $"soffice.exe must exist at {path}");
    }
}
