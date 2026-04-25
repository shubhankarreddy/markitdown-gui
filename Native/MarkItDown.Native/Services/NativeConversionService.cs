using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using MarkItDown.Native.Models;
using Microsoft.VisualBasic.FileIO;
using ReverseMarkdown;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MarkItDown.Native.Services;

public sealed class NativeConversionService
{
    private static readonly HttpClient HttpClient = new();
    private static readonly HashSet<string> NativeExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".txt", ".md", ".csv", ".json", ".xml", ".yaml", ".yml", ".html", ".htm",
        ".docx", ".pptx", ".xlsx", ".pdf", ".zip", ".ipynb"
    };

    public bool CanHandle(string source)
    {
        if (IsWebUri(source, out _))
        {
            return true;
        }

        var extension = Path.GetExtension(source);
        return !string.IsNullOrWhiteSpace(extension) && NativeExtensions.Contains(extension);
    }

    public async Task<ConversionResult> ConvertAsync(string source)
    {
        try
        {
            var markdown = IsWebUri(source, out var uri)
                ? await ConvertUrlAsync(uri!)
                : await ConvertFileAsync(source, 0);

            return new ConversionResult
            {
                Success = true,
                Text = markdown.Trim(),
                Engine = "Native C#"
            };
        }
        catch (Exception ex)
        {
            return new ConversionResult
            {
                Success = false,
                Error = ex.Message,
                Detail = ex.ToString(),
                Engine = "Native C#"
            };
        }
    }

    private async Task<string> ConvertUrlAsync(Uri uri)
    {
        using var response = await HttpClient.GetAsync(uri);
        response.EnsureSuccessStatusCode();

        var mediaType = response.Content.Headers.ContentType?.MediaType?.ToLowerInvariant() ?? string.Empty;
        if (mediaType.Contains("html", StringComparison.Ordinal))
        {
            var html = await response.Content.ReadAsStringAsync();
            return ConvertHtmlDocument(html, uri.ToString());
        }

        if (mediaType.StartsWith("text/", StringComparison.Ordinal))
        {
            var text = await response.Content.ReadAsStringAsync();
            return text.Trim();
        }

        var extension = Path.GetExtension(uri.AbsolutePath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            throw new InvalidOperationException("That URL does not point to a natively supported web document.");
        }

        var tempPath = Path.Combine(Path.GetTempPath(), "markitdown-native-" + Guid.NewGuid() + extension);
        try
        {
            await using var input = await response.Content.ReadAsStreamAsync();
            await using var output = File.Create(tempPath);
            await input.CopyToAsync(output);

            return await ConvertFileAsync(tempPath, 0);
        }
        finally
        {
            TryDeleteFile(tempPath);
        }
    }

    private async Task<string> ConvertFileAsync(string path, int depth)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The selected file could not be found.", path);
        }

        if (depth > 2)
        {
            throw new InvalidOperationException("Nested archive conversion is limited to keep the native engine predictable.");
        }

        var extension = Path.GetExtension(path).ToLowerInvariant();
        return extension switch
        {
            ".txt" or ".md" => await File.ReadAllTextAsync(path),
            ".json" => await ConvertJsonAsync(path),
            ".xml" => await ConvertXmlAsync(path),
            ".yaml" or ".yml" => await ConvertCodeBlockAsync(path, "yaml"),
            ".csv" => ConvertCsv(path),
            ".html" or ".htm" => await ConvertHtmlFileAsync(path),
            ".docx" => ConvertDocx(path),
            ".pptx" => ConvertPptx(path),
            ".xlsx" => ConvertXlsx(path),
            ".pdf" => ConvertPdf(path),
            ".zip" => await ConvertZipAsync(path, depth + 1),
            ".ipynb" => await ConvertNotebookAsync(path),
            _ => throw new InvalidOperationException($"Native C# conversion for '{extension}' is not implemented yet.")
        };
    }

    private static async Task<string> ConvertJsonAsync(string path)
    {
        using var document = JsonDocument.Parse(await File.ReadAllTextAsync(path));
        var pretty = JsonSerializer.Serialize(document.RootElement, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        return Fence("json", pretty);
    }

    private static async Task<string> ConvertXmlAsync(string path)
    {
        var document = XDocument.Parse(await File.ReadAllTextAsync(path), LoadOptions.PreserveWhitespace);
        return Fence("xml", document.ToString());
    }

    private static async Task<string> ConvertCodeBlockAsync(string path, string language)
    {
        return Fence(language, await File.ReadAllTextAsync(path));
    }

    private static string ConvertCsv(string path)
    {
        var rows = new List<IReadOnlyList<string>>();
        using var parser = new TextFieldParser(path)
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true
        };
        parser.SetDelimiters(",");

        while (!parser.EndOfData)
        {
            rows.Add(parser.ReadFields()?.Select(value => value ?? string.Empty).ToArray() ?? Array.Empty<string>());
        }

        return MarkdownTable(rows, "Column");
    }

    private static async Task<string> ConvertHtmlFileAsync(string path)
    {
        return ConvertHtmlDocument(await File.ReadAllTextAsync(path), path);
    }

    private static string ConvertHtmlDocument(string html, string sourceLabel)
    {
        var doc = new HtmlAgilityPack.HtmlDocument();
        doc.LoadHtml(html);

        foreach (var node in doc.DocumentNode.SelectNodes("//script|//style|//noscript") ?? Enumerable.Empty<HtmlNode>())
        {
            node.Remove();
        }

        var converter = new Converter();
        var markdown = converter.Convert(doc.DocumentNode.OuterHtml).Trim();
        var title = doc.DocumentNode.SelectSingleNode("//title")?.InnerText?.Trim();

        if (!string.IsNullOrWhiteSpace(title) &&
            !markdown.StartsWith("# " + title, StringComparison.OrdinalIgnoreCase))
        {
            return $"# {title}\n\nSource: {sourceLabel}\n\n{markdown}";
        }

        return markdown;
    }

    private static string ConvertDocx(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart?.Document?.Body
                   ?? throw new InvalidOperationException("The DOCX file does not contain a readable document body.");

        var markdown = new StringBuilder();

        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case W.Paragraph paragraph:
                    AppendParagraph(markdown, paragraph);
                    break;
                case W.Table table:
                    markdown.AppendLine(ConvertWordTable(table));
                    markdown.AppendLine();
                    break;
            }
        }

        return markdown.ToString().Trim();
    }

    private static void AppendParagraph(StringBuilder markdown, W.Paragraph paragraph)
    {
        var text = NormalizeWhitespace(paragraph.InnerText);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        if (style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(new string(style.Where(char.IsDigit).ToArray()), out var level))
        {
            level = Math.Clamp(level, 1, 6);
            markdown.AppendLine($"{new string('#', level)} {text}");
            markdown.AppendLine();
            return;
        }

        markdown.AppendLine(text);
        markdown.AppendLine();
    }

    private static string ConvertWordTable(W.Table table)
    {
        var rows = table.Elements<W.TableRow>()
            .Select(row => (IReadOnlyList<string>)row.Elements<W.TableCell>()
                .Select(cell => NormalizeWhitespace(cell.InnerText))
                .ToArray())
            .Where(row => row.Count > 0)
            .ToList();

        return MarkdownTable(rows, "Column");
    }

    private static string ConvertPptx(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = document.PresentationPart
                               ?? throw new InvalidOperationException("The PPTX file does not contain a presentation part.");
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToList()
                       ?? throw new InvalidOperationException("The PPTX file does not contain any slides.");

        var markdown = new StringBuilder();

        for (var index = 0; index < slideIds.Count; index++)
        {
            var relationshipId = slideIds[index].RelationshipId?.Value
                                 ?? throw new InvalidOperationException("A slide is missing its relationship ID.");
            var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
            var slideTitle = GetSlideTitle(slidePart) ?? $"Slide {index + 1}";

            markdown.AppendLine($"# {slideTitle}");
            markdown.AppendLine();

            foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
            {
                var text = NormalizeWhitespace(shape.TextBody?.InnerText ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(text) &&
                    !string.Equals(text, slideTitle, StringComparison.OrdinalIgnoreCase))
                {
                    markdown.AppendLine($"- {text}");
                }
            }

            foreach (var table in slidePart.Slide.Descendants<A.Table>())
            {
                markdown.AppendLine();
                markdown.AppendLine(MarkdownTable(
                    table.Elements<A.TableRow>()
                        .Select(row => (IReadOnlyList<string>)row.Elements<A.TableCell>()
                            .Select(cell => NormalizeWhitespace(cell.InnerText))
                            .ToArray())
                        .ToList(),
                    "Column"));
            }

            markdown.AppendLine();
        }

        return markdown.ToString().Trim();
    }

    private static string? GetSlideTitle(SlidePart slidePart)
    {
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            var placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();

            var placeholderType = placeholder?.Type?.Value;
            if (placeholderType == P.PlaceholderValues.Title ||
                placeholderType == P.PlaceholderValues.CenteredTitle)
            {
                var title = NormalizeWhitespace(shape.TextBody?.InnerText ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(title))
                {
                    return title;
                }
            }
        }

        return null;
    }

    private static string ConvertXlsx(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbookPart = document.WorkbookPart
                           ?? throw new InvalidOperationException("The XLSX file does not contain a workbook.");
        var sheets = workbookPart.Workbook.Sheets?.Elements<S.Sheet>().ToList()
                     ?? throw new InvalidOperationException("The XLSX file does not contain any worksheets.");

        var markdown = new StringBuilder();

        foreach (var sheet in sheets)
        {
            var sheetName = sheet.Name?.Value ?? "Sheet";
            var relationshipId = sheet.Id?.Value
                                 ?? throw new InvalidOperationException("A worksheet is missing its relationship ID.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
            var rows = ExtractWorksheetRows(worksheetPart, workbookPart);

            markdown.AppendLine($"# {sheetName}");
            markdown.AppendLine();
            markdown.AppendLine(MarkdownTable(rows, "Column"));
            markdown.AppendLine();
        }

        return markdown.ToString().Trim();
    }

    private static List<IReadOnlyList<string>> ExtractWorksheetRows(WorksheetPart worksheetPart, WorkbookPart workbookPart)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<S.SheetData>();
        if (sheetData is null)
        {
            return new List<IReadOnlyList<string>>();
        }

        var rowMaps = new List<Dictionary<int, string>>();
        var maxColumn = 0;

        foreach (var row in sheetData.Elements<S.Row>())
        {
            var values = new Dictionary<int, string>();
            foreach (var cell in row.Elements<S.Cell>())
            {
                var columnIndex = GetColumnIndex(cell.CellReference?.Value);
                values[columnIndex] = GetCellValue(cell, workbookPart);
                maxColumn = Math.Max(maxColumn, columnIndex + 1);
            }

            rowMaps.Add(values);
        }

        var rows = new List<IReadOnlyList<string>>();
        foreach (var rowMap in rowMaps)
        {
            var row = Enumerable.Range(0, maxColumn)
                .Select(index => rowMap.TryGetValue(index, out var value) ? value : string.Empty)
                .ToArray();
            rows.Add(row);
        }

        return rows;
    }

    private static string ConvertPdf(string path)
    {
        using var document = PdfDocument.Open(path);
        var markdown = new StringBuilder();

        foreach (Page page in document.GetPages())
        {
            markdown.AppendLine($"## Page {page.Number}");
            markdown.AppendLine();
            markdown.AppendLine(ContentOrderTextExtractor.GetText(page).Trim());
            markdown.AppendLine();
        }

        return markdown.ToString().Trim();
    }

    private async Task<string> ConvertZipAsync(string path, int depth)
    {
        using var archive = ZipFile.OpenRead(path);
        var parts = new List<string>();
        var tempRoot = Path.Combine(Path.GetTempPath(), "markitdown-native-zip-" + Guid.NewGuid());
        Directory.CreateDirectory(tempRoot);

        try
        {
            foreach (var entry in archive.Entries.Where(entry => !string.IsNullOrWhiteSpace(entry.Name)))
            {
                var extension = Path.GetExtension(entry.Name);
                if (string.IsNullOrWhiteSpace(extension) ||
                    !NativeExtensions.Contains(extension) ||
                    string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var extractedPath = Path.Combine(tempRoot, Path.GetFileName(entry.Name));
                entry.ExtractToFile(extractedPath, overwrite: true);

                var converted = await ConvertFileAsync(extractedPath, depth);
                parts.Add($"# {entry.FullName}\n\n{converted.Trim()}");
            }

            if (parts.Count == 0)
            {
                throw new InvalidOperationException("The ZIP archive did not contain any natively supported documents.");
            }

            return string.Join("\n\n", parts);
        }
        finally
        {
            TryDeleteDirectory(tempRoot);
        }
    }

    private static async Task<string> ConvertNotebookAsync(string path)
    {
        using var document = JsonDocument.Parse(await File.ReadAllTextAsync(path));
        var cells = document.RootElement.TryGetProperty("cells", out var cellArray)
            ? cellArray.EnumerateArray()
            : Enumerable.Empty<JsonElement>();

        var markdown = new StringBuilder();
        foreach (var cell in cells)
        {
            var cellType = cell.TryGetProperty("cell_type", out var typeElement)
                ? typeElement.GetString()
                : string.Empty;
            var source = cell.TryGetProperty("source", out var sourceElement)
                ? string.Concat(sourceElement.EnumerateArray().Select(part => part.GetString()))
                : string.Empty;

            if (string.IsNullOrWhiteSpace(source))
            {
                continue;
            }

            if (string.Equals(cellType, "markdown", StringComparison.OrdinalIgnoreCase))
            {
                markdown.AppendLine(source.Trim());
            }
            else
            {
                markdown.AppendLine(Fence("python", source.TrimEnd()));
            }

            markdown.AppendLine();
        }

        return markdown.ToString().Trim();
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return 0;
        }

        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var index = 0;
        foreach (var letter in letters)
        {
            index *= 26;
            index += char.ToUpperInvariant(letter) - 'A' + 1;
        }

        return Math.Max(0, index - 1);
    }

    private static string GetCellValue(S.Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellValue is null)
        {
            return string.Empty;
        }

        var value = cell.CellValue.InnerText;
        if (cell.DataType?.Value == S.CellValues.SharedString &&
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedIndex))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?
                       .Elements<S.SharedStringItem>()
                       .ElementAtOrDefault(sharedIndex)?
                       .InnerText ?? string.Empty;
        }

        return value;
    }

    private static string MarkdownTable(IReadOnlyList<IReadOnlyList<string>> rows, string defaultHeaderPrefix)
    {
        if (rows.Count == 0)
        {
            return "_No content found._";
        }

        var maxColumns = rows.Max(row => row.Count);
        if (maxColumns == 0)
        {
            return "_No content found._";
        }

        var normalizedRows = rows
            .Select(row => Enumerable.Range(0, maxColumns)
                .Select(index => index < row.Count ? EscapeTableCell(row[index]) : string.Empty)
                .ToArray())
            .ToList();

        string[] header;
        List<string[]> body;

        if (normalizedRows.Count > 1)
        {
            header = normalizedRows[0];
            body = normalizedRows.Skip(1).ToList();
        }
        else
        {
            header = Enumerable.Range(1, maxColumns)
                .Select(index => $"{defaultHeaderPrefix} {index}")
                .ToArray();
            body = normalizedRows;
        }

        var builder = new StringBuilder();
        builder.AppendLine($"| {string.Join(" | ", header)} |");
        builder.AppendLine($"| {string.Join(" | ", Enumerable.Repeat("---", maxColumns))} |");

        foreach (var row in body)
        {
            builder.AppendLine($"| {string.Join(" | ", row)} |");
        }

        return builder.ToString().TrimEnd();
    }

    private static string EscapeTableCell(string value)
    {
        return NormalizeWhitespace(value).Replace("|", "\\|", StringComparison.Ordinal);
    }

    private static string NormalizeWhitespace(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static string Fence(string language, string content)
    {
        return $"```{language}\n{content.TrimEnd()}\n```";
    }

    private static bool IsWebUri(string source, out Uri? uri)
    {
        if (Uri.TryCreate(source, UriKind.Absolute, out var parsedUri) &&
            parsedUri.Scheme is "http" or "https")
        {
            uri = parsedUri;
            return true;
        }

        uri = null;
        return false;
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
        }
    }
}
