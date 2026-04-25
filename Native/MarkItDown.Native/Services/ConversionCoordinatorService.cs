using MarkItDown.Native.Models;

namespace MarkItDown.Native.Services;

public sealed class ConversionCoordinatorService
{
    private readonly NativeConversionService _nativeConversionService = new();
    private readonly MarkItDownBackendService _pythonFallbackService = new();

    public async Task<ConversionResult> ConvertAsync(string source, LlmOptions? llm)
    {
        if (llm is null && _nativeConversionService.CanHandle(source))
        {
            var nativeResult = await _nativeConversionService.ConvertAsync(source);
            if (nativeResult.Success)
            {
                return nativeResult;
            }

            var fallbackResult = await _pythonFallbackService.ConvertAsync(source, llm);
            if (fallbackResult.Success)
            {
                return fallbackResult.WithDetailPrefix(
                    $"Native conversion failed first, so the app fell back to Python.\r\n\r\nNative error:\r\n{nativeResult.Error}");
            }

            return fallbackResult.WithDetailPrefix(
                $"Native conversion failed first.\r\n\r\nNative error:\r\n{nativeResult.Error}");
        }

        return await _pythonFallbackService.ConvertAsync(source, llm);
    }
}

internal static class ConversionResultExtensions
{
    public static ConversionResult WithDetailPrefix(this ConversionResult result, string prefix)
    {
        var combinedDetail = string.IsNullOrWhiteSpace(result.Detail)
            ? prefix
            : prefix + "\r\n\r\nFallback detail:\r\n" + result.Detail;

        return new ConversionResult
        {
            Success = result.Success,
            Text = result.Text,
            Error = result.Error,
            Detail = combinedDetail,
            Engine = result.Engine
        };
    }
}
