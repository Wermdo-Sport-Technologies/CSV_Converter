using System;
using System.Collections.Generic;
//using System.Runtime.Serialization;
using System.Text;
using System.IO;
using ClosedXML.Excel;
using CsvConverterApp.Models;


namespace CsvConverterApp.Services;

public class ExportService
{
    internal class ExportService
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

    }
}
