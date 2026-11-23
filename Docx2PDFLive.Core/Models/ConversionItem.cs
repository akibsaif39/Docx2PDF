using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Docx2PDFLive.Core.Models;

public class ConversionItem : INotifyPropertyChanged
{
    private ConversionStatus _status;
    private string _engine = "Pending";
    private string _timestamp = string.Empty;
    private string? _errorMessage;

    public string FilePath { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public int RunId { get; set; }

    public ConversionStatus Status
    {
        get => _status;
        set
        {
            if (_status != value)
            {
                _status = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusDisplay));
            }
        }
    }

    public string StatusDisplay => _status switch
    {
        ConversionStatus.Queued => "Queued",
        ConversionStatus.Converting => "Converting...",
        ConversionStatus.Synced => "Synced",
        ConversionStatus.Error => "Error",
        _ => _status.ToString()
    };

    public string Engine
    {
        get => _engine;
        set
        {
            if (_engine != value)
            {
                _engine = value;
                OnPropertyChanged();
            }
        }
    }

    public string Timestamp
    {
        get => _timestamp;
        set
        {
            if (_timestamp != value)
            {
                _timestamp = value;
                OnPropertyChanged();
            }
        }
    }

    public string? ErrorMessage
    {
        get => _errorMessage;
        set
        {
            if (_errorMessage != value)
            {
                _errorMessage = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasError));
            }
        }
    }

    public bool HasError => !string.IsNullOrEmpty(_errorMessage);

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
