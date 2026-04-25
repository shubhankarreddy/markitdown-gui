namespace MarkItDown.Native.Models;

public sealed class LlmOptions
{
    public string ApiKey { get; init; } = string.Empty;

    public string Model { get; init; } = "gpt-4o-mini";
}
