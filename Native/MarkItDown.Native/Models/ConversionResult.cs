namespace MarkItDown.Native.Models;

public sealed class ConversionResult
{
    public bool Success { get; init; }

    public string Text { get; init; } = string.Empty;

    public string Error { get; init; } = string.Empty;

    public string Detail { get; init; } = string.Empty;

    public string Engine { get; init; } = string.Empty;
}
