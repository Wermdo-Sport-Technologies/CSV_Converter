using System;
using System.Collections.Generic;
//using System.Runtime.Serialization;
using System.Text;
using System.IO;
using ClosedXML.Excel;
using CsvConverterApp.Models;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace CsvConverterApp.Services;

/// <summary>
/// Converts parsed CSV rows into the selected output format.
/// Supported technologies are ClosedXML for Excel .xlsx files, OpenXML for Word
/// .docx files, and plain System.IO text writing for Markdown and .txt files.
/// </summary>
public class ExportService
{

    /// <summary>
    /// Applies the selected export options, optionally removes duplicates,
    /// optionally groups rows by a split column, and writes one file per group.
    /// </summary>
    public void Export(List<Dictionary<string, string>> rows, ExportOptions options)
    {
        var sourceRows = rows;

        if (options.RemoveDuplicates)
        {
            sourceRows = sourceRows
                .GroupBy(row => string.Join("|", row.Values))
                .Select(g => g.First())
                .ToList();
        }

        List<ExportGroup> groups;

        if (string.IsNullOrWhiteSpace(options.SplitByColumn))
        {
            groups =[new ExportGroup{Name = "export", Rows = sourceRows}];
        }
        else
        {
            groups = sourceRows.GroupBy(row => row.TryGetValue(options.SplitByColumn, out var value) ? value : "Unknown").Select(g => new ExportGroup{Name = CleanFileName(g.Key), Rows = g.ToList()}).ToList();
        }

        if (options.ExportFormat.ToLower() == "xlsx" && options.ExcelUseSheetsInsteadOfFiles)
        {
            var path = Path.Combine(options.OutputFolder, "Export.xlsx");
            ExportExcelWithSheets(groups.ToList(), options.SelectedColumns, path);
            return;
        }

        foreach (var group in groups)
        {
            var exportRows = SelectColumns(group.Rows, options.SelectedColumns);

            var path = Path.Combine(
                options.OutputFolder,
                $"{group.Name}.{options.ExportFormat}");
            switch (options.ExportFormat.ToLower())
            {
                case "xlsx":
                    ExportExcel(exportRows, path);
                    break;
                case "docx":
                    ExportWord(exportRows, path);
                    break;
                case "md":
                    ExportMarkdown(exportRows, path);
                    break;
                case "txt":
                    ExportText(exportRows, path);
                    break;
                default:
                    throw new NotSupportedException($"Export format {options.ExportFormat} is not supported.");
            }
        }
    }

    private List<Dictionary<string, string>> SelectColumns(
    List<Dictionary<string, string>> rows,
    List<string> selectedColumns)
    {
        return rows
            .Select(row => selectedColumns.ToDictionary(
                col => col,
                col => row.TryGetValue(col, out var value) ? value : ""))
            .ToList();
    }

    /// <summary>
    /// Creates a Word document using the OpenXML SDK.
    /// Each selected row is written as readable label/value paragraphs.
    /// </summary>
    private void ExportWord(List<Dictionary<string, string>> rows, string filePath)
    {
        using var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);

        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();

        body.AppendChild(new Paragraph(new Run(new Text("Exported Data"))));

        body.AppendChild(new Paragraph(new Run(new Text($"Generated: {DateTime.Now:HH:mm dd-MM-yyyy}"))));

        body.AppendChild(new Paragraph(new Run(new Text("")))); // empty line

        foreach (var row in rows)
        {
            foreach(var item in row)
            {
                var paragraph = new Paragraph(
                new Run(
                    new Text($"{item.Key}: {item.Value}")));

                body.AppendChild(paragraph);
            }
            body.AppendChild(new Paragraph(
            new Run(
                new Text(""))));
        }

        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }

    /// <summary>
    /// Creates an Excel workbook using ClosedXML.
    /// The first row contains headers and the following rows contain exported CSV values.
    /// Creates one sheet per group when the "split by column" option is used, otherwise creates a single sheet with all rows.
    /// </summary>
    private void ExportExcelWithSheets(List<ExportGroup> groups, List<string> _selectedColumns, string filePath)
    {
        using var workBook = new XLWorkbook();

        foreach (var group in groups)
        {
            var exportRows = SelectColumns(group.Rows, _selectedColumns);
            var sheetName = CleanWorksheetName(group.Name);

            var ws = workBook.Worksheets.Add(sheetName);

            if(exportRows.Count == 0)
            {
                continue;
            }

            var headers = exportRows.First().Keys.ToList();
            for (int i = 0; i < headers.Count; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
            }

            for (int r = 0; r < exportRows.Count; r++)
            {
                for (int c = 0; c < headers.Count; c++)
                {
                    ws.Cell(r + 2, c + 1).Value = exportRows[r][headers[c]];
                }
            }
            ws.Columns().AdjustToContents();
        }
        workBook.SaveAs(filePath);
    }

    /// <summary>
    /// Creates an Excel workbook using ClosedXML.
    /// The first row contains headers and the following rows contain exported CSV values.
    /// </summary>
    private void ExportExcel(List<Dictionary<string, string>> rows, string filePath)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("data");

        if(rows.Count == 0)
        {
            workbook.SaveAs(filePath);
            return;
        }

        var headers = rows.First().Keys.ToList();

        for (int i = 0; i < headers.Count; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        for(int r = 0; r < rows.Count; r++)
        {
            for(int c = 0; c < headers.Count; c++)
            {
                ws.Cell(r + 2, c + 1).Value = rows[r][headers[c]];
            }
        }

        ws.Columns().AdjustToContents();
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Writes rows as a Markdown table using plain text output.
    /// This format is useful for documentation, GitHub comments, or simple sharing.
    /// </summary>
    private void ExportMarkdown(List<Dictionary<string, string>> rows, string filePath)
    {
        if(rows.Count == 0)
        {
            File.WriteAllText(filePath, "");
            return;
        }

        var headers = rows.First().Keys.ToList();
        var sb = new StringBuilder();
        
        sb.AppendLine("| " + string.Join(" | ", headers) + " | ");
        sb.AppendLine("| " + string.Join(" | ", headers.Select(_ => "---")) + " |");

        foreach(var row in rows)
        {
            sb.AppendLine("| " + string.Join(" | ", headers.Select(h => row[h])) + " |");
        }

        File.WriteAllText(filePath, sb.ToString());
    }

    /// <summary>
    /// Writes rows as simple space-separated text.
    /// This is the lightest output format and does not require any document library.
    /// </summary>
    private void ExportText(List<Dictionary<string, string>> rows, string path)
    {
        var sb = new StringBuilder();

        foreach (var row in rows)
        {
            foreach (var item in row)
            {
                sb.Append($"{item.Value} ");
            }

            sb.AppendLine();
        }

        File.WriteAllText(path, sb.ToString());
    }

    /// <summary>
    /// Replaces characters that Windows does not allow in file names.
    /// Empty split values become "Unknown" so every group can still be exported.
    /// </summary>
    private static string CleanFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }

        return string.IsNullOrWhiteSpace(name) ? "Unknown" : name;
    }

    /// <summary>
    /// Escapes values for CSV-style text output.
    /// This helper is currently unused because the app exports txt, md, xlsx, and docx.
    /// </summary>
    private static string EscapeCsv(string value)
    {
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }

    private static string CleanWorksheetName(string name)
    {
        char[] invalidChars = ['\\', '/', '*', '[', ']', ':', '?'];

        foreach (var c in invalidChars)
        {
            name = name.Replace(c, '_');
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            name = "Sheet";
        }

        return name.Length > 31 ? name[..31] : name;
    }
}
