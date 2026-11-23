namespace Docx2PDFLive.Core.Services;

public class ConversionResult
{
    public bool Success { get; set; }
    public string EngineUsed { get; set; } = string.Empty;
    public string? ErrorMessage { get; set; }
    public long ElapsedMilliseconds { get; set; }

    public static ConversionResult Ok(string engine, long elapsedMs = 0) => new()
    {
        Success = true,
        EngineUsed = engine,
        ElapsedMilliseconds = elapsedMs
    };

    public static ConversionResult Fail(string engine, string error, long elapsedMs = 0) => new()
    {
        Success = false,
        EngineUsed = engine,
        ErrorMessage = error,
        ElapsedMilliseconds = elapsedMs
    };
}
