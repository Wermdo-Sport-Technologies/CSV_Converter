using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Globalization;
using CsvHelper;

namespace CsvConverterApp.Services;

public class CsvReaderService
{
    public List<String> ReadHeaders(string path)
    {
        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

        csv.Read();
        csv.ReadHeader();

        return csv.HeaderRecord?.ToList() ?? [];
    }

    public List<Dictionary<string, string>> ReadRows(string path)
    {
        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

        var rows = new List<Dictionary<string, string>>();

        csv.Read();
        csv.ReadHeader();

        while (csv.Read())
        {
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
