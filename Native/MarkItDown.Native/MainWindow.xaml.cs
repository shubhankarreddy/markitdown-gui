using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using MarkItDown.Native.Models;
using MarkItDown.Native.Services;
using Microsoft.Win32;
using Forms = System.Windows.Forms;

namespace MarkItDown.Native;

public partial class MainWindow : Window
{
    private static readonly string[] SupportedExtensions =
    {
        ".pdf", ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls",
        ".html", ".htm", ".xml", ".txt", ".csv", ".json", ".yaml", ".yml",
        ".md", ".png", ".jpg", ".jpeg", ".gif", ".bmp",
        ".mp3", ".wav", ".mp4", ".ipynb", ".zip"
    };

    private readonly ObservableCollection<ConversionQueueItem> _queue = new();
    private readonly ConversionCoordinatorService _conversionService = new();
    private bool _isConverting;
    private string _outputFolder = string.Empty;

    public MainWindow()
    {
        InitializeComponent();
        QueueListView.ItemsSource = _queue;
        UpdatePreview();
        UpdateUiState();
    }

    private ConversionQueueItem? SelectedItem => QueueListView.SelectedItem as ConversionQueueItem;

    private void AddFilesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Multiselect = true,
            Filter = BuildFileDialogFilter()
        };

        if (dialog.ShowDialog(this) == true)
        {
            AddFiles(dialog.FileNames);
        }
    }

    private void AddUrlButton_Click(object sender, RoutedEventArgs e)
    {
        AddUrlFromInput();
    }

    private void UrlTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddUrlFromInput();
            e.Handled = true;
        }
    }

    private void SelectOutputButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Choose where converted Markdown files should be saved.",
            ShowNewFolderButton = true,
            SelectedPath = string.IsNullOrWhiteSpace(_outputFolder)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                : _outputFolder
        };

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            _outputFolder = dialog.SelectedPath;
            OutputFolderTextBlock.Text = _outputFolder;
            UpdateStatus($"Output folder set to {_outputFolder}");
            UpdateUiState();
            UpdatePreview();
        }
    }

    private void RemoveSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is null || _isConverting)
        {
            return;
        }

        var nextIndex = QueueListView.SelectedIndex;
        _queue.Remove(SelectedItem);

        if (_queue.Count > 0)
        {
            QueueListView.SelectedIndex = Math.Min(nextIndex, _queue.Count - 1);
        }

        UpdateStatus("Selected item removed from the queue.");
        UpdatePreview();
        UpdateUiState();
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isConverting)
        {
            return;
        }

        _queue.Clear();
        UpdateStatus("Queue cleared.");
        UpdatePreview();
        UpdateUiState();
    }

    private void QueueListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdatePreview();
        UpdateUiState();
    }

    private void EnableLlmCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        LlmGroupBox.Visibility = EnableLlmCheckBox.IsChecked == true
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private async void ConvertAllButton_Click(object sender, RoutedEventArgs e)
    {
        var pending = _queue
            .Where(item => item.Status is QueueItemStatus.Queued or QueueItemStatus.Error)
            .ToList();

        if (pending.Count == 0 || _isConverting)
        {
            UpdateStatus("There is nothing queued to convert.");
            return;
        }

        var llm = BuildLlmOptions();
        if (EnableLlmCheckBox.IsChecked == true && llm is null)
        {
            return;
        }

        _isConverting = true;
        UpdateUiState();
        ConversionProgressBar.Visibility = Visibility.Visible;
        ConversionProgressBar.Minimum = 0;
        ConversionProgressBar.Maximum = pending.Count;
        ConversionProgressBar.Value = 0;

        try
        {
            for (var i = 0; i < pending.Count; i++)
            {
                var item = pending[i];
                item.Status = QueueItemStatus.Working;
                item.Error = string.Empty;
                item.Engine = string.Empty;
                item.Detail = string.Empty;
                QueueListView.SelectedItem = item;
                QueueListView.ScrollIntoView(item);
                UpdatePreview();
                UpdateStatus($"Converting {i + 1} of {pending.Count}: {item.DisplayName}");

                var result = await _conversionService.ConvertAsync(item.Source, llm);

                if (result.Success)
                {
                    item.Result = result.Text;
                    item.Error = string.Empty;
                    item.Status = QueueItemStatus.Done;
                    item.Engine = result.Engine;
                    item.Detail = result.Detail;
                    await AutoSaveResultAsync(item);
                }
                else
                {
                    item.Result = string.Empty;
                    item.Error = string.IsNullOrWhiteSpace(result.Error)
                        ? "Conversion failed."
                        : result.Error;
                    item.Status = QueueItemStatus.Error;
                    item.Engine = result.Engine;
                    item.Detail = result.Detail;
                }

                ConversionProgressBar.Value = i + 1;
                UpdatePreview();
            }

            UpdateStatus("Conversion complete.");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "MarkItDown",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);

            UpdateStatus("Conversion stopped because a conversion engine could not be started.");
        }
        finally
        {
            _isConverting = false;
            ConversionProgressBar.Visibility = Visibility.Collapsed;
            UpdateUiState();
        }
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is { HasResult: true } item)
        {
            System.Windows.Clipboard.SetText(item.Result);
            UpdateStatus("Markdown copied to the clipboard.");
        }
    }

    private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is not { HasResult: true } item)
        {
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = GetSafeFileStem(item) + ".md",
            Filter = "Markdown (*.md)|*.md|All files (*.*)|*.*"
        };

        if (dialog.ShowDialog(this) == true)
        {
            await File.WriteAllTextAsync(dialog.FileName, item.Result);
            item.OutputPath = dialog.FileName;
            UpdateStatus($"Saved {Path.GetFileName(dialog.FileName)}");
            UpdatePreview();
            UpdateUiState();
        }
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var folder = GetOpenFolderTarget();
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = "explorer.exe",
            UseShellExecute = true
        };
        startInfo.ArgumentList.Add(folder);
        Process.Start(startInfo);
    }

    private void Window_PreviewDragOver(object sender, System.Windows.DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop)
            ? System.Windows.DragDropEffects.Copy
            : System.Windows.DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            return;
        }

        if (e.Data.GetData(System.Windows.DataFormats.FileDrop) is string[] paths)
        {
            AddFiles(paths);
        }
    }

    private void AddFiles(IEnumerable<string> files)
    {
        var known = new HashSet<string>(
            _queue.Where(item => !item.IsUrl).Select(item => item.Source),
            StringComparer.OrdinalIgnoreCase);

        var added = 0;

        foreach (var file in files)
        {
            if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
            {
                continue;
            }

            if (!IsSupportedFile(file) || !known.Add(file))
            {
                continue;
            }

            _queue.Add(new ConversionQueueItem(file, false, Path.GetFileName(file)));
            added++;
        }

        if (added > 0 && QueueListView.SelectedItem is null)
        {
            QueueListView.SelectedIndex = 0;
        }

        UpdateStatus(added == 0
            ? "No new supported files were added."
            : $"Added {added} file{(added == 1 ? string.Empty : "s")} to the queue.");
        UpdateUiState();
    }

    private void AddUrlFromInput()
    {
        var text = UrlTextBox.Text.Trim();

        if (!Uri.TryCreate(text, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            UpdateStatus("Enter a valid http or https URL first.");
            return;
        }

        if (_queue.Any(item => item.IsUrl && string.Equals(item.Source, text, StringComparison.OrdinalIgnoreCase)))
        {
            UpdateStatus("That URL is already in the queue.");
            return;
        }

        _queue.Add(new ConversionQueueItem(text, true, Shorten(text, 60)));
        UrlTextBox.Clear();

        if (QueueListView.SelectedItem is null)
        {
            QueueListView.SelectedIndex = 0;
        }

        UpdateStatus("URL added to the queue.");
        UpdateUiState();
    }

    private void UpdatePreview()
    {
        if (SelectedItem is not { } item)
        {
            SelectedItemTextBlock.Text = "No item selected";
            StatsTextBlock.Text = "Choose a queued item to preview its status and converted Markdown.";
            MarkdownTextBox.Text = string.Empty;
            DetailsTextBox.Text = "Drag files in, add a URL, or use Add Files to get started.";
            CopyButton.IsEnabled = false;
            SaveButton.IsEnabled = false;
            OpenFolderButton.IsEnabled = !string.IsNullOrWhiteSpace(GetOpenFolderTarget()) &&
                                         Directory.Exists(GetOpenFolderTarget());
            return;
        }

        SelectedItemTextBlock.Text = item.DisplayName;
        StatsTextBlock.Text = BuildStatsText(item);
        MarkdownTextBox.Text = item.Result;
        DetailsTextBox.Text = BuildDetailsText(item);
        CopyButton.IsEnabled = item.HasResult;
        SaveButton.IsEnabled = item.HasResult;
        OpenFolderButton.IsEnabled = !string.IsNullOrWhiteSpace(GetOpenFolderTarget()) &&
                                     Directory.Exists(GetOpenFolderTarget());
    }

    private void UpdateUiState()
    {
        AddFilesButton.IsEnabled = !_isConverting;
        AddUrlButton.IsEnabled = !_isConverting;
        UrlTextBox.IsEnabled = !_isConverting;
        SelectOutputButton.IsEnabled = !_isConverting;
        RemoveSelectedButton.IsEnabled = !_isConverting && SelectedItem is not null;
        ClearButton.IsEnabled = !_isConverting && _queue.Count > 0;
        ConvertAllButton.IsEnabled = !_isConverting &&
                                     _queue.Any(item => item.Status is QueueItemStatus.Queued or QueueItemStatus.Error);
    }

    private string BuildStatsText(ConversionQueueItem item)
    {
        if (!item.HasResult)
        {
            return $"Status: {item.StatusLabel}";
        }

        var words = item.Result
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Length;
        var lines = item.Result.Length == 0 ? 0 : item.Result.Count(ch => ch == '\n') + 1;

        return $"Status: {item.StatusLabel} | {words} words | {lines} lines | {item.Result.Length} chars";
    }

    private string BuildDetailsText(ConversionQueueItem item)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"Source: {item.Source}");
        builder.AppendLine($"Type: {(item.IsUrl ? "URL" : "File")}");
        builder.AppendLine($"Status: {item.StatusLabel}");
        if (!string.IsNullOrWhiteSpace(item.Engine))
        {
            builder.AppendLine($"Engine: {item.Engine}");
        }

        if (!string.IsNullOrWhiteSpace(item.OutputPath))
        {
            builder.AppendLine($"Saved to: {item.OutputPath}");
        }
        else if (!string.IsNullOrWhiteSpace(_outputFolder) && item.Status == QueueItemStatus.Done)
        {
            builder.AppendLine($"Output folder: {_outputFolder}");
        }

        if (!string.IsNullOrWhiteSpace(item.Error))
        {
            builder.AppendLine();
            builder.AppendLine("Error:");
            builder.AppendLine(item.Error);
        }

        if (!string.IsNullOrWhiteSpace(item.Detail))
        {
            builder.AppendLine();
            builder.AppendLine("Detail:");
            builder.AppendLine(item.Detail);
        }

        if (!item.HasResult && string.IsNullOrWhiteSpace(item.Error))
        {
            builder.AppendLine();
            builder.AppendLine("Converted Markdown will appear here after the item finishes.");
        }

        return builder.ToString().TrimEnd();
    }

    private LlmOptions? BuildLlmOptions()
    {
        if (EnableLlmCheckBox.IsChecked != true)
        {
            return null;
        }

        var apiKey = ApiKeyPasswordBox.Password.Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            System.Windows.MessageBox.Show(
                this,
                "Enter your OpenAI API key or turn off LLM descriptions.",
                "MarkItDown",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            ApiKeyPasswordBox.Focus();
            return null;
        }

        var model = (ModelComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString();
        return new LlmOptions
        {
            ApiKey = apiKey,
            Model = string.IsNullOrWhiteSpace(model) ? "gpt-4o-mini" : model
        };
    }

    private async Task AutoSaveResultAsync(ConversionQueueItem item)
    {
        if (string.IsNullOrWhiteSpace(_outputFolder) || !item.HasResult)
        {
            return;
        }

        Directory.CreateDirectory(_outputFolder);
        var outputPath = GetUniqueOutputPath(item);
        await File.WriteAllTextAsync(outputPath, item.Result);
        item.OutputPath = outputPath;
    }

    private string GetUniqueOutputPath(ConversionQueueItem item)
    {
        var baseName = GetSafeFileStem(item);
        var outputPath = Path.Combine(_outputFolder, baseName + ".md");
        var counter = 1;

        while (File.Exists(outputPath))
        {
            outputPath = Path.Combine(_outputFolder, $"{baseName}_{counter}.md");
            counter++;
        }

        return outputPath;
    }

    private string GetSafeFileStem(ConversionQueueItem item)
    {
        var baseName = item.IsUrl ? "web-page" : Path.GetFileNameWithoutExtension(item.DisplayName);
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "output";
        }

        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(baseName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "output" : cleaned;
    }

    private string GetOpenFolderTarget()
    {
        if (SelectedItem is { OutputPath: not null } item && !string.IsNullOrWhiteSpace(item.OutputPath))
        {
            var folder = Path.GetDirectoryName(item.OutputPath);
            if (!string.IsNullOrWhiteSpace(folder))
            {
                return folder;
            }
        }

        return _outputFolder;
    }

    private void UpdateStatus(string message)
    {
        StatusTextBlock.Text = message;
    }

    private static bool IsSupportedFile(string path)
    {
        return SupportedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);
    }

    private static string BuildFileDialogFilter()
    {
        var patterns = string.Join(";", SupportedExtensions.Select(ext => "*" + ext));
        return $"Supported files|{patterns}|All files (*.*)|*.*";
    }

    private static string Shorten(string value, int maxLength)
    {
        return value.Length <= maxLength
            ? value
            : value[..(maxLength - 3)] + "...";
    }
}
