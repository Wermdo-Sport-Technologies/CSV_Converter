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

public class ExportService
{
    public void Export(List<Dictionary<string, string>> rows, ExportOptions options)
    {
        //var filterdRows = rows
        //    .Select(row => options.SelectedColumns.ToDictionary(
        //        col => col,
        //        col => row.TryGetValue(col, out var value) ? value : ""))
        //    .ToList();

        //if (options.RemoveDuplicates)
        //{
        //    filterdRows = filterdRows.DistinctBy(row => string.Join("|", row.Values)).ToList();
        //}

        //var groups = string.IsNullOrWhiteSpace(options.SplitByColumn) ? new[] { new { Name = "export", Rows = filterdRows } } : filterdRows
        //    .GroupBy(row => row.GetValueOrDefault(options.SplitByColumn, "Unknown"))
        //    .Select(g => new { Name = CleanFileName(g.Key), Rows = g.ToList() });

        var sourceRows = rows;

        if (options.RemoveDuplicates)
        {
            sourceRows = sourceRows
                .DistinctBy(row => string.Join("|", row.Values))
                .ToList();
        }

        var groups = string.IsNullOrWhiteSpace(options.SplitByColumn)
            ? new[] { new { Name = "export", Rows = sourceRows } }
            : sourceRows
                .GroupBy(row =>
                    row.TryGetValue(options.SplitByColumn, out var value)
                        ? value
                        : "Unknown")
                .Select(g => new
                {
                    Name = CleanFileName(g.Key),
                    Rows = g.ToList()
                });

        foreach (var group in groups)
        {
            var path = Path.Combine(options.OutputFolder, $"{group.Name}.{options.ExportFormat}");

            switch(options.ExportFormat.ToLower())
            {
                case "xlsx":
                    foreach (var _group in groups)
                    {
                        var exportRows = group.Rows
                            .Select(row => options.SelectedColumns.ToDictionary(
                                col => col,
                                col => row.TryGetValue(col, out var value)
                                    ? value
                                    : ""))
                            .ToList();

                        var _path = Path.Combine(
                            options.OutputFolder,
                            $"{group.Name}.{options.ExportFormat}");

                        ExportExcel(exportRows, _path);
                    }
                    break;

                case "md":
                    foreach (var _group in groups)
                    {
                        var exportRows = group.Rows
                            .Select(row => options.SelectedColumns.ToDictionary(
                                col => col,
                                col => row.TryGetValue(col, out var value)
                                    ? value
                                    : ""))
                            .ToList();

                        var _path = Path.Combine(
                            options.OutputFolder,
                            $"{group.Name}.{options.ExportFormat}");

                        ExportMarkdown(exportRows, _path);
                    }
                    break;

                case "txt":
                    foreach (var _group in groups)
                    {
                        var exportRows = group.Rows
                            .Select(row => options.SelectedColumns.ToDictionary(
                                col => col,
                                col => row.TryGetValue(col, out var value)
                                    ? value
                                    : ""))
                            .ToList();

                        var _path = Path.Combine(
                            options.OutputFolder,
                            $"{group.Name}.{options.ExportFormat}");

                        ExportText(exportRows, _path);
                    }
                    break;
                case "docx":
                    foreach (var _group in groups)
                    {
                        var exportRows = group.Rows
                            .Select(row => options.SelectedColumns.ToDictionary(
                                col => col,
                                col => row.TryGetValue(col, out var value)
                                    ? value
                                    : ""))
                            .ToList();

                        var _path = Path.Combine(
                            options.OutputFolder,
                            $"{group.Name}.{options.ExportFormat}");

                        ExportWord(exportRows, _path);
                    }
                    break;

            }
        }
    }

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

    private static string CleanFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }

        return string.IsNullOrWhiteSpace(name) ? "Unknown" : name;
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }
}
