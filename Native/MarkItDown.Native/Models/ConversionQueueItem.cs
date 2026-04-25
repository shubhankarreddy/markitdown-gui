using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MarkItDown.Native.Models;

public sealed class ConversionQueueItem : INotifyPropertyChanged
{
    private QueueItemStatus _status = QueueItemStatus.Queued;
    private string _result = string.Empty;
    private string _error = string.Empty;
    private string? _outputPath;
    private string _engine = string.Empty;
    private string _detail = string.Empty;

    public ConversionQueueItem(string source, bool isUrl, string displayName)
    {
        Source = source;
        IsUrl = isUrl;
        DisplayName = displayName;
    }

    public string Source { get; }

    public bool IsUrl { get; }

    public string DisplayName { get; }

    public QueueItemStatus Status
    {
        get => _status;
        set
        {
            if (_status == value)
            {
                return;
            }

            _status = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(StatusLabel));
        }
    }

    public string Result
    {
        get => _result;
        set
        {
            if (_result == value)
            {
                return;
            }

            _result = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasResult));
        }
    }

    public string Error
    {
        get => _error;
        set
        {
            if (_error == value)
            {
                return;
            }

            _error = value;
            OnPropertyChanged();
        }
    }

    public string? OutputPath
    {
        get => _outputPath;
        set
        {
            if (_outputPath == value)
            {
                return;
            }

            _outputPath = value;
            OnPropertyChanged();
        }
    }

    public string Engine
    {
        get => _engine;
        set
        {
            if (_engine == value)
            {
                return;
            }

            _engine = value;
            OnPropertyChanged();
        }
    }

    public string Detail
    {
        get => _detail;
        set
        {
            if (_detail == value)
            {
                return;
            }

            _detail = value;
            OnPropertyChanged();
        }
    }

    public bool HasResult => !string.IsNullOrWhiteSpace(Result);

    public string StatusLabel => Status switch
    {
        QueueItemStatus.Queued => "Queued",
        QueueItemStatus.Working => "Working",
        QueueItemStatus.Done => "Done",
        QueueItemStatus.Error => "Error",
        _ => "Queued"
    };

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public enum QueueItemStatus
{
    Queued,
    Working,
    Done,
    Error
}
