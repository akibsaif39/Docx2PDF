using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class FirstRunMarkerTests
{
    [Fact]
    public void FirstRunMarker_CreatesAndDetectsFlagFile()
    {
        FirstRunMarker.MarkWelcomed();
        Assert.True(FirstRunMarker.HasBeenWelcomed(), "FirstRunMarker should return true after MarkWelcomed()");
    }
}
