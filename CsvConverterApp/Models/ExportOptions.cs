using System;
using System.Collections.Generic;
using System.Text;

namespace CsvConverterApp.Models;

/// <summary>
/// Holds all choices the user makes in the WPF window before export starts.
/// Keeping these settings in one model makes ExportService independent from
/// UI controls such as buttons, list boxes, and combo boxes.
/// </summary>
public class ExportOptions
{
    /// <summary>
    /// Full path to the CSV file selected by the user.
    /// </summary>
    public string InputPath { get; set; } = "";

    /// <summary>
    /// Folder where the generated export files will be written.
    /// </summary>
    public string OutputFolder { get; set; } = "";

    /// <summary>
    /// Header names selected in the UI. Only these columns are included in exports.
    /// </summary>
    public List<string> SelectedColumns { get; set; } = new();

    /// <summary>
    /// When true, rows with the same full source values are exported only once.
    /// </summary>
    public bool RemoveDuplicates { get; set; } // TODO: add "what is considered a duplicate" option

    /// <summary>
    /// Output file type chosen in the UI, such as txt, xlsx, md, or docx.
    /// </summary>
    public string ExportFormat { get; set; } = "txt";

    /// <summary>
    /// Optional header name used to split rows into separate files.
    /// For example, choosing a grade column creates one file per grade value.
    /// </summary>
    public string? SplitByColumn { get; set; }
}
