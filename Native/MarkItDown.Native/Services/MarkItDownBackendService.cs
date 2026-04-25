using System.Diagnostics;
using System.IO;
using System.Text.Json;
using MarkItDown.Native.Models;

namespace MarkItDown.Native.Services;

public sealed class MarkItDownBackendService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public async Task<ConversionResult> ConvertAsync(string source, LlmOptions? llm)
    {
        Exception? lastLaunchException = null;

        foreach (var command in BuildCommands())
        {
            try
            {
                return await ExecuteAsync(command, source, llm);
            }
            catch (Exception ex) when (ex is FileNotFoundException or InvalidOperationException or System.ComponentModel.Win32Exception)
            {
                lastLaunchException = ex;
            }
        }

        throw new InvalidOperationException(
            "Could not find a MarkItDown backend runtime. Run setup-native.bat to install Python dependencies, or build MarkItDownBackend.exe first.",
            lastLaunchException);
    }

    private static IEnumerable<BackendCommand> BuildCommands()
    {
        foreach (var root in GetSearchRoots())
        {
            foreach (var candidate in new[]
                     {
                         Path.Combine(root, "MarkItDownBackend.exe"),
                         Path.Combine(root, "dist", "MarkItDownBackend.exe"),
                         Path.Combine(root, "native_publish", "win-x64", "MarkItDownBackend.exe")
                     })
            {
                if (File.Exists(candidate))
                {
                    yield return new BackendCommand(candidate, Array.Empty<string>(), Path.GetDirectoryName(candidate)!);
                }
            }
        }

        var scriptPath = FindBackendScriptPath();
        if (scriptPath is null)
        {
            yield break;
        }

        foreach (var command in GetPythonCommands())
        {
            yield return new BackendCommand(command, new[] { scriptPath }, Path.GetDirectoryName(scriptPath)!);
        }
    }

    private static async Task<ConversionResult> ExecuteAsync(BackendCommand command, string source, LlmOptions? llm)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = command.FileName,
            WorkingDirectory = command.WorkingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var argument in command.PrefixArguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        startInfo.ArgumentList.Add("--source");
        startInfo.ArgumentList.Add(source);

        if (llm is not null && !string.IsNullOrWhiteSpace(llm.ApiKey))
        {
            startInfo.ArgumentList.Add("--api-key");
            startInfo.ArgumentList.Add(llm.ApiKey);

            if (!string.IsNullOrWhiteSpace(llm.Model))
            {
                startInfo.ArgumentList.Add("--model");
                startInfo.ArgumentList.Add(llm.Model);
            }
        }

        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
        {
            throw new InvalidOperationException("The backend process could not be started.");
        }

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        var payload = TryParsePayload(stdout);
        if (payload is not null)
        {
            return new ConversionResult
            {
                Success = payload.Success && process.ExitCode == 0,
                Text = payload.Text ?? string.Empty,
                Error = payload.Error ?? (process.ExitCode == 0 ? string.Empty : stderr.Trim()),
                Detail = payload.Detail ?? stderr.Trim(),
                Engine = "Python MarkItDown"
            };
        }

        if (process.ExitCode == 0)
        {
            return new ConversionResult
            {
                Success = true,
                Text = stdout,
                Engine = "Python MarkItDown"
            };
        }

        return new ConversionResult
        {
            Success = false,
            Error = string.IsNullOrWhiteSpace(stderr)
                ? $"Backend exited with code {process.ExitCode}."
                : stderr.Trim(),
            Detail = stdout.Trim(),
            Engine = "Python MarkItDown"
        };
    }

    private static BackendPayload? TryParsePayload(string stdout)
    {
        var trimmed = stdout.Trim();
        if (string.IsNullOrWhiteSpace(trimmed) || !trimmed.StartsWith("{", StringComparison.Ordinal))
        {
            return null;
        }

        try
        {
            return JsonSerializer.Deserialize<BackendPayload>(trimmed, JsonOptions);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static string? FindBackendScriptPath()
    {
        foreach (var root in GetSearchRoots())
        {
            foreach (var candidate in new[]
                     {
                         Path.Combine(root, "Backend", "markitdown_backend.py"),
                         Path.Combine(root, "Native", "MarkItDown.Native", "Backend", "markitdown_backend.py"),
                         Path.Combine(root, "markitdown_backend.py")
                     })
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static IEnumerable<string> GetPythonCommands()
    {
        var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var root in GetSearchRoots())
        {
            var candidate = Path.Combine(root, "venv", "Scripts", "python.exe");
            if (File.Exists(candidate) && yielded.Add(candidate))
            {
                yield return candidate;
            }
        }

        foreach (var command in new[] { "python", "py" })
        {
            if (yielded.Add(command))
            {
                yield return command;
            }
        }
    }

    private static IEnumerable<string> GetSearchRoots()
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var seed in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            if (string.IsNullOrWhiteSpace(seed))
            {
                continue;
            }

            var current = Path.GetFullPath(seed);
            while (visited.Add(current))
            {
                yield return current;

                var parent = Directory.GetParent(current);
                if (parent is null)
                {
                    break;
                }

                current = parent.FullName;
            }
        }
    }

    private sealed record BackendCommand(string FileName, IReadOnlyList<string> PrefixArguments, string WorkingDirectory);

    private sealed class BackendPayload
    {
        public bool Success { get; init; }

        public string? Text { get; init; }

        public string? Error { get; init; }

        public string? Detail { get; init; }
    }
}
