//using CsvConverterApp.Services;
//using Microsoft.Win32;
//using System.Text;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;

//namespace CsvConverterApp;

///// <summary>
///// Interaction logic for MainWindow.xaml
///// </summary>
//public partial class MainWindow : Window
//{
//    private readonly CsvReaderService _csvReader = new();
//    private readonly ExportService _exportService = new();

//    private string _inputPath = "";
//    private string _outputFolder = "";

//    public MainWindow()
//    {
//        InitializeComponent();
//        FormatComboBox.SelectedIndex = 0;
//    }

//    private void ChooseInputFile_Click(object sender, RoutedEventArgs e)
//    {
//        var dialog = new OpenFileDialog
//        {
//            Filter = "CSV files (*.csv)|*.csv"
//        };

//        if (dialog.ShowDialog() == true)
//        {
//            _inputPath = dialog.FileName;

//            var headers = _csvReader.ReadHeaders(_inputPath);

//            ColumnsList.ItemsSource = headers;
//            SplitColumnComboBox.ItemsSource = headers;

//            StatusText.Text = $"Laddade {headers.Count} kolumner.";
//        }
//    }

//    private void ChooseOutputFolder_Click(object sender, RoutedEventArgs e)
//    {
//        var dialog = new System.Windows.Forms.FolderBrowserDialog();

//        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
//        {
//            _outputFolder = dialog.SelectedPath;
//            StatusText.Text = $"Output-mapp: {_outputFolder}";
//        }
//    }

//    private void Convert_Click(object sender, RoutedEventArgs e)
//    {
//        if (string.IsNullOrWhiteSpace(_inputPath) || string.IsNullOrWhiteSpace(_outputFolder))
//        {
//            MessageBox.Show("Välj både CSV-fil och output-mapp.");
//            return;
//        }

//        var selectedColumns = ColumnsList.SelectedItems
//            .Cast<string>()
//            .ToList();

//        if (selectedColumns.Count == 0)
//        {
//            MessageBox.Show("Välj minst en kolumn.");
//            return;
//        }

//        var format = ((System.Windows.Controls.ComboBoxItem)FormatComboBox.SelectedItem)
//            .Content
//            .ToString()!;

//        var options = new ExportOptions
//        {
//            InputPath = _inputPath,
//            OutputFolder = _outputFolder,
//            SelectedColumns = selectedColumns,
//            RemoveDuplicates = RemoveDuplicatesCheckBox.IsChecked == true,
//            ExportFormat = format,
//            SplitByColumn = SplitColumnComboBox.SelectedItem as string
//        };

//        var rows = _csvReader.ReadRows(_inputPath);
//        _exportService.Export(rows, options);

//        StatusText.Text = "Konvertering klar.";
//    }
//}

using CsvConverterApp.Models;
using CsvConverterApp.Services;
using Microsoft.Win32;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CsvConverterApp;

public partial class MainWindow : Window
{
    private readonly CsvReaderService _csvReader = new();
    private readonly ExportService _exportService = new();

    private string _inputPath = "";
    private string _outputFolder = "";

    public MainWindow()
    {
        InitializeComponent();
        FormatComboBox.SelectedIndex = 0;
    }

    private void ChooseInputFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {
            _inputPath = dialog.FileName;

            var headers = _csvReader.ReadHeaders(_inputPath);

            ColumnsList.ItemsSource = headers;
            SplitColumnComboBox.ItemsSource = headers;

            StatusText.Text = $"Laddade {headers.Count} kolumner.";
        }
    }

    private void ChooseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _outputFolder = dialog.SelectedPath;
            StatusText.Text = $"Output-mapp: {_outputFolder}";
        }
    }

    private void Convert_Click(object sender, RoutedEventArgs e)
    {
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

        var selectedItem = FormatComboBox.SelectedItem as ComboBoxItem;
        var format = selectedItem?.Content?.ToString() ?? "xlsx";

        var options = new ExportOptions
        {
            InputPath = _inputPath,
            OutputFolder = _outputFolder,
            SelectedColumns = selectedColumns,
            RemoveDuplicates = RemoveDuplicatesCheckBox.IsChecked == true,
            ExportFormat = format,
            SplitByColumn = SplitColumnComboBox.SelectedItem as string
        };

        var rows = _csvReader.ReadRows(_inputPath);
        _exportService.Export(rows, options);

        StatusText.Text = "Konvertering klar.";
    }
}