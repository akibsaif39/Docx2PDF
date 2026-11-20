using System.Diagnostics;
using System.IO;

namespace Docx2PDFLive.Core.Services;

/// <summary>
/// High-fidelity composite converter.
/// Strictly uses Microsoft Word (native, high-speed warm instance) or LibreOffice.
/// Never degrades to simple/openxml format conversions.
/// </summary>
public class CompositeConverter : IConverter, IDisposable, IAsyncDisposable
{
    private readonly WarmWordConverter? _wordConverter;
    private readonly LibreOfficeConverter? _libreOfficeConverter;
    private readonly OfficeEngineType _preferredEngine;
    private bool _disposed;

    public string Name => ActiveEngineName;

    public string ActiveEngineName
    {
        get
        {
            if (_wordConverter != null) return "Microsoft Word (Native High-Speed)";
            if (_libreOfficeConverter != null) return "LibreOffice";
            return "No Engine Available";
        }
    }

    public CompositeConverter()
    {
        var available = OfficeEngineDetector.GetAvailableEngines();

        if (available.HasFlag(OfficeEngineType.MicrosoftWord))
        {
            try
            {
                _wordConverter = new WarmWordConverter();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize WarmWordConverter: {ex.Message}");
            }
        }

        if (available.HasFlag(OfficeEngineType.LibreOffice) || LibreOfficeDetector.IsAvailable(out var loPath))
        {
            try
            {
                _libreOfficeConverter = new LibreOfficeConverter();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize LibreOfficeConverter: {ex.Message}");
            }
        }

        _preferredEngine = OfficeEngineDetector.GetFastestAvailableEngine();
    }

    public async Task<ConversionResult> ConvertAsync(string docxPath, string pdfPath, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(docxPath))
        {
            return ConversionResult.Fail(Name, $"Source DOCX file not found: {docxPath}");
        }

        // 1. Try Primary Engine (Word prioritized for speed and 100% native layout)
        if (_preferredEngine == OfficeEngineType.MicrosoftWord && _wordConverter != null)
        {
            try
            {
                var wordResult = await _wordConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
                if (wordResult.Success && File.Exists(pdfPath))
                {
                    return wordResult;
                }

                Debug.WriteLine($"Word conversion failed ({wordResult.ErrorMessage}). Checking for LibreOffice fallback...");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception invoking Word converter: {ex.Message}");
            }

            // Fallback to LibreOffice if Word encountered an unexpected issue
            if (_libreOfficeConverter != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var loResult = await _libreOfficeConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
                if (loResult.Success && File.Exists(pdfPath))
                {
                    return loResult;
                }
            }

            return ConversionResult.Fail(Name, "Conversion failed with Microsoft Word. Ensure the document is not corrupt or password-protected.");
        }

        // 2. Primary Engine is LibreOffice (e.g. Word is not installed)
        if (_libreOfficeConverter != null)
        {
            try
            {
                var loResult = await _libreOfficeConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
                if (loResult.Success && File.Exists(pdfPath))
                {
                    return loResult;
                }

                Debug.WriteLine($"LibreOffice conversion failed: {loResult.ErrorMessage}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception invoking LibreOffice: {ex.Message}");
            }

            // If LibreOffice failed but Word is available
            if (_wordConverter != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var wordResult = await _wordConverter.ConvertAsync(docxPath, pdfPath, cancellationToken);
                if (wordResult.Success && File.Exists(pdfPath))
                {
                    return wordResult;
                }
            }

            return ConversionResult.Fail(Name, "Conversion failed with LibreOffice. Ensure the document is not corrupt or password-protected.");
        }

        return ConversionResult.Fail(Name, "No high-fidelity conversion engine (Microsoft Word or LibreOffice) is installed on this system.");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _wordConverter?.Dispose();
        GC.SuppressFinalize(this);
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_wordConverter != null)
        {
            await _wordConverter.DisposeAsync();
        }

        GC.SuppressFinalize(this);
    }
}
