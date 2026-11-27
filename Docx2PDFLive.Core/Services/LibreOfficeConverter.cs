using System.Diagnostics;
using System.IO;
using Docx2PDFLive.Core.Interop;

namespace Docx2PDFLive.Core.Services;

public class LibreOfficeConverter : IConverter
{
    private readonly string _sofficePath;
    private readonly TimeSpan _timeout;
    private static readonly SemaphoreSlim ProfileLock = new(1, 1);
    private static readonly string PersistentProfileDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Docx2PDF",
        "LOProfile");

    public string Name => "LibreOffice";

    public LibreOfficeConverter(string? customSofficePath = null, TimeSpan? timeout = null)
    {
        _sofficePath = !string.IsNullOrEmpty(customSofficePath)
            ? customSofficePath
            : (LibreOfficeDetector.FindSofficeExecutable() ?? "soffice.exe");
        _timeout = timeout ?? TimeSpan.FromSeconds(35);
    }

    public async Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();

        if (!File.Exists(docxPath))
        {
            return ConversionResult.Fail(Name, $"Source DOCX file not found: {docxPath}");
        }

        var outDir = Path.GetDirectoryName(pdfPath);
        if (string.IsNullOrEmpty(outDir))
        {
            outDir = Path.GetDirectoryName(docxPath) ?? Directory.GetCurrentDirectory();
        }

        if (!Directory.Exists(outDir))
        {
            Directory.CreateDirectory(outDir);
        }

        // Expected output path produced by LibreOffice
        var expectedPdfName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
        var generatedPdfPath = Path.Combine(outDir, expectedPdfName);

        // Delete any existing stale PDF file before starting to ensure freshness
        try
        {
            if (File.Exists(generatedPdfPath))
            {
                File.Delete(generatedPdfPath);
            }
        }
        catch { }

        await ProfileLock.WaitAsync(cancellationToken);
        try
        {
            if (!Directory.Exists(PersistentProfileDir))
            {
                Directory.CreateDirectory(PersistentProfileDir);
            }
            var profileUri = "file:///" + PersistentProfileDir.Replace('\\', '/');

            var args = $"-env:UserInstallation=\"{profileUri}\" --headless --convert-to pdf \"{docxPath}\" --outdir \"{outDir}\"";

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

            // Prevent external Python environment variables from interfering with LibreOffice's internal Python runtime
            psi.EnvironmentVariables.Remove("PYTHONHOME");
            psi.EnvironmentVariables.Remove("PYTHONPATH");

            using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };

            if (!process.Start())
            {
                return ConversionResult.Fail(Name, "Failed to start soffice.exe process.");
            }

            // Assign LibreOffice process to Job Object so it is guaranteed to terminate if app closes
            JobObjectManager.AssignProcess(process);

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

            // Check if PDF file was generated
            if (File.Exists(generatedPdfPath))
            {
                // If requested pdfPath is different from generated (e.g. casing), handle rename
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
            sw.Stop();
            return ConversionResult.Fail(Name, $"LibreOffice execution error: {ex.Message}", sw.ElapsedMilliseconds);
        }
        finally
        {
            ProfileLock.Release();
        }
    }
}
