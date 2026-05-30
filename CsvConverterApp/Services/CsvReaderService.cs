using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Globalization;
using CsvHelper;
using System.Linq;

namespace CsvConverterApp.Services;

/// <summary>
/// Reads CSV files for the converter.
/// The app uses CsvHelper so quoted values, commas inside fields, and header
/// parsing are handled by a CSV library instead of manual string splitting.
/// </summary>
public class CsvReaderService
{
    /// <summary>
    /// Reads only the header row from the CSV file.
    /// The headers are used by the WPF UI so the user can choose which columns
    /// to export and which column should split the output files.
    /// </summary>
    public List<String> ReadHeaders(string path)
    {
        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

        csv.Read();
        csv.ReadHeader();

        return csv.HeaderRecord?.ToList() ?? [];
    }

    /// <summary>
    /// Reads every data row from the CSV file into dictionaries keyed by header name.
    /// This structure makes the export code format-independent because each output
    /// writer can ask for values by column name instead of by numeric index.
    /// </summary>
    public List<Dictionary<string, string>> ReadRows(string path)
    {
        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

        var rows = new List<Dictionary<string, string>>();

        csv.Read();
        csv.ReadHeader();

        while (csv.Read())
        {
            // Each CSV record becomes one dictionary: header name -> cell value.
            var row = new Dictionary<string, string>();
            foreach (var header in csv.HeaderRecord)
            {
                row[header] = csv.GetField(header) ?? "";
            }
            rows.Add(row);
        }

        return rows;
    }
}
