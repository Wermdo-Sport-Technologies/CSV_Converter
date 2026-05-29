using System;
using System.Collections.Generic;
using System.Text;

namespace CsvConverterApp.Models;

public class ExportOptions
{
    public string InputPath { get; set; } = "";
    public string OutputFolder { get; set; } = "";
    public List<string> SelectedColumns { get; set; } = new();
    public bool RemoveDuplicates { get; set; } // TODO: add "what is considered a duplicate" option
    public string ExportFormat { get; set; } = "txt";
    public string? SplitByColumn { get; set; }
}
