using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvConverterApp.Models;

public class ExportGroup
{
    public string Name { get; set; } = "";
    public List<Dictionary<string, string>> Rows { get; set; } = new();
}