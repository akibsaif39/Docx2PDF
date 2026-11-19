namespace Docx2PDFLive.Core.Services;

public class ConversionResult
{
    public bool Success { get; }
    public string EngineUsed { get; }
    public string? ErrorMessage { get; }
    public long ElapsedMilliseconds { get; }

    private ConversionResult(bool success, string engineUsed, string? errorMessage, long elapsedMilliseconds)
    {
        Success = success;
        EngineUsed = engineUsed;
        ErrorMessage = errorMessage;
        ElapsedMilliseconds = elapsedMilliseconds;
    }

    public static ConversionResult Ok(string engineUsed, long elapsedMs = 0) =>
        new(true, engineUsed, null, elapsedMs);

    public static ConversionResult Fail(string engineUsed, string errorMessage, long elapsedMs = 0) =>
        new(false, engineUsed, errorMessage, elapsedMs);
}
