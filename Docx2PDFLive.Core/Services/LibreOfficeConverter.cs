using System.Diagnostics;
using System.IO;

namespace Docx2PDFLive.Core.Services;

public class LibreOfficeConverter : IConverter
{
    private readonly string _sofficePath;
    private readonly TimeSpan _timeout;

    public string Name => "LibreOffice Headless";

    public LibreOfficeConverter(string? sofficePath = null, TimeSpan? timeout = null)
    {
        if (!string.IsNullOrEmpty(sofficePath) && File.Exists(sofficePath))
        {
            _sofficePath = sofficePath;
        }
        else if (LibreOfficeDetector.IsAvailable(out var detectedPath))
        {
            _sofficePath = detectedPath;
        }
        else
        {
            throw new FileNotFoundException("LibreOffice soffice.exe was not found on this machine.");
        }

        _timeout = timeout ?? TimeSpan.FromSeconds(60);
    }

    public async Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();

        var outDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
        if (string.IsNullOrEmpty(outDir)) outDir = Directory.GetCurrentDirectory();

        var docxFileNameNoExt = Path.GetFileNameWithoutExtension(docxPath);
        var generatedPdfPath = Path.Combine(outDir, docxFileNameNoExt + ".pdf");

        var args = $"--headless --convert-to pdf:writer_pdf_Export \"{Path.GetFullPath(docxPath)}\" --outdir \"{outDir}\"";

        var psi = new ProcessStartInfo
        {
            FileName = _sofficePath,
            Arguments = args,
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = outDir
        };

        psi.EnvironmentVariables.Remove("PYTHONHOME");
        psi.EnvironmentVariables.Remove("PYTHONPATH");

        try
        {
            using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };

            if (!process.Start())
            {
                return ConversionResult.Fail(Name, "Failed to start soffice.exe process.");
            }

            var stdOutTask = process.StandardOutput.ReadToEndAsync();
            var stdErrTask = process.StandardError.ReadToEndAsync();

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(_timeout);

            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch { }

                if (cancellationToken.IsCancellationRequested)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                return ConversionResult.Fail(Name, $"Conversion timed out after {_timeout.TotalSeconds} seconds.");
            }

            var stdOut = await stdOutTask;
            var stdErr = await stdErrTask;
            sw.Stop();

            if (File.Exists(generatedPdfPath))
            {
                if (!string.Equals(Path.GetFullPath(generatedPdfPath), Path.GetFullPath(pdfPath), StringComparison.OrdinalIgnoreCase))
                {
                    File.Move(generatedPdfPath, pdfPath, overwrite: true);
                }
                return ConversionResult.Ok(Name, sw.ElapsedMilliseconds);
            }

            var errorDetail = !string.IsNullOrWhiteSpace(stdErr) ? stdErr : stdOut;
            if (string.IsNullOrWhiteSpace(errorDetail))
            {
                errorDetail = $"Exit code: {process.ExitCode}, PDF output was not created.";
            }

            return ConversionResult.Fail(Name, errorDetail, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return ConversionResult.Fail(Name, $"Error invoking LibreOffice: {ex.Message}", sw.ElapsedMilliseconds);
        }
    }
}
