using CsvConverterApp.Models;
using CsvConverterApp.Services;
using Microsoft.Win32;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CsvConverterApp;

/// <summary>
/// Main WPF window for the CSV Converter.
/// This class coordinates the UI workflow: choose a CSV file, choose output
/// settings, read rows with CsvReaderService, and export them with ExportService.
/// </summary>
public partial class MainWindow : Window
{
    // Service classes keep CSV parsing and file generation separate from WPF UI code.
    private readonly CsvReaderService _csvReader = new();
    private readonly ExportService _exportService = new();

    // These paths are chosen by the user through Windows dialogs before conversion.
    private string _inputPath = "";
    private string _outputFolder = "";

    public MainWindow()
    {
        InitializeComponent();
        FormatComboBox.SelectedIndex = 0;
    }

    private void ChooseInputFile_Click(object sender, RoutedEventArgs e)
    {
        // Microsoft.Win32.OpenFileDialog is the standard WPF-friendly file picker.
        var dialog = new OpenFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {
            _inputPath = dialog.FileName;

            // The first CSV row is treated as headers and becomes the list of selectable columns.
            var headers = _csvReader.ReadHeaders(_inputPath);

            ColumnsList.ItemsSource = headers;
            SplitColumnComboBox.ItemsSource = headers;

            StatusText.Text = $"Laddade {headers.Count} kolumner.";
        }
    }

    private void ChooseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        // WPF does not include its own simple folder picker, so the app uses Windows Forms here.
        var dialog = new System.Windows.Forms.FolderBrowserDialog();

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _outputFolder = dialog.SelectedPath;
            StatusText.Text = $"Output-mapp: {_outputFolder}";
        }
    }

    private void Convert_Click(object sender, RoutedEventArgs e)
    {
        // Conversion needs both an input file and an output folder because exports are written to disk.
        if (string.IsNullOrWhiteSpace(_inputPath) || string.IsNullOrWhiteSpace(_outputFolder))
        {
            MessageBox.Show("Välj både CSV-fil och output-mapp.");
            return;
        }

        var selectedColumns = ColumnsList.SelectedItems
            .Cast<string>()
            .ToList();

        if (selectedColumns.Count == 0)
        {
            MessageBox.Show("Välj minst en kolumn.");
            return;
        }

        // ComboBox values are stored as WPF ComboBoxItem controls, so the text content is extracted here.
        var selectedItem = FormatComboBox.SelectedItem as ComboBoxItem;
        var format = selectedItem?.Content?.ToString() ?? "xlsx";

        // ExportOptions is a small data object that passes all UI choices to the export layer.
        var options = new ExportOptions
        {
            InputPath = _inputPath,
            OutputFolder = _outputFolder,
            SelectedColumns = selectedColumns,
            RemoveDuplicates = RemoveDuplicatesCheckBox.IsChecked == true,
            ExportFormat = format,
            SplitByColumn = SplitColumnComboBox.SelectedItem as string
        };

        // CsvReaderService returns rows as dictionaries keyed by column header.
        // ExportService then filters, splits, and writes the selected output format.
        var rows = _csvReader.ReadRows(_inputPath);
        _exportService.Export(rows, options);

        StatusText.Text = "Konvertering klar.";
    }
}
