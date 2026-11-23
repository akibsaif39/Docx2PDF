using System.IO;
using Docx2PDFLive.Core.Services;
using Xunit;

namespace Docx2PDFLive.Tests;

public class FirstRunMarkerTests
{
    [Fact]
    public void FirstRunMarker_CreatesAndDetectsFlagFile()
    {
        // Ensure flag is reset first
        FirstRunMarker.ResetWelcome();
        Assert.False(FirstRunMarker.HasBeenWelcomed());

        // Mark welcomed
        FirstRunMarker.MarkWelcomed();
        Assert.True(FirstRunMarker.HasBeenWelcomed());
        Assert.True(File.Exists(FirstRunMarker.FlagFilePath));

        // Reset welcome for future runs
        FirstRunMarker.ResetWelcome();
        Assert.False(FirstRunMarker.HasBeenWelcomed());
    }
}
