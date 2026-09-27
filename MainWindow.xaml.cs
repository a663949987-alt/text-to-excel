using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace TextToExcel;

public partial class MainWindow : Window
{
    private string? _currentFile;
    private List<string> _headers = [];
    private List<List<string>> _rows = [];
    private string _detectedName = "等待识别";
    private bool _ready;

    public MainWindow()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        InitializeComponent();
        _ready = true;
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = HasTextFile(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void DropArea_DragEnter(object sender, DragEventArgs e)
    {
        if (HasTextFile(e.Data)) DropArea.Background = new SolidColorBrush(Color.FromRgb(232, 247, 239));
    }

    private void DropArea_DragLeave(object sender, DragEventArgs e) => DropArea.Background = new SolidColorBrush(Color.FromRgb(250, 252, 253));

    private void Window_Drop(object sender, DragEventArgs e)
    {
        DropArea.Background = new SolidColorBrush(Color.FromRgb(250, 252, 253));
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0) LoadFile(files[0]);
    }

    private static bool HasTextFile(IDataObject data) => data.GetData(DataFormats.FileDrop) is string[] files && files.Any(f => new[] { ".txt", ".csv", ".tsv" }.Contains(Path.GetExtension(f).ToLowerInvariant()));

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "文本文件 (*.txt;*.csv;*.tsv)|*.txt;*.csv;*.tsv|所有文件 (*.*)|*.*" };
        if (dlg.ShowDialog() == true) LoadFile(dlg.FileName);
    }

    private void LoadFile(string path)
    {
        if (!File.Exists(path)) return;
        try
        {
            _currentFile = path;
            InputText.Text = File.ReadAllText(path, SelectedEncoding());
            FileInfo.Text = $"已载入：{Path.GetFileName(path)}  ·  {new FileInfo(path).Length / 1024d:F1} KB";
            StatusText.Text = "文件已载入，正在预览";
        }
        catch (Exception ex) { MessageBox.Show($"无法读取文件：\n{ex.Message}", "读取失败", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private Encoding SelectedEncoding()
    {
        var tag = (EncodingBox.SelectedItem as ComboBoxItem)?.Tag?.ToString();
        return tag switch { "gb18030" => Encoding.GetEncoding("GB18030"), "utf-16" => Encoding.Unicode, _ => new UTF8Encoding(false) };
    }

    private void Encoding_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_ready && _currentFile != null) LoadFile(_currentFile);
    }

    private void InputText_TextChanged(object sender, TextChangedEventArgs e) { if (_ready) ParseAndPreview(); }
    private void Option_Changed(object sender, RoutedEventArgs e) { if (_ready) ParseAndPreview(); }
    private void Option_Changed(object sender, SelectionChangedEventArgs e) { if (_ready) ParseAndPreview(); }

    private void ParseAndPreview()
    {
        var lines = InputText.Text.Replace("\r\n", "\n").Replace('\r', '\n').TrimStart('\uFEFF').Split('\n').ToList();
        if (SkipEmptyCheck.IsChecked == true) lines = lines.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        if (lines.Count == 0 || lines.All(string.IsNullOrWhiteSpace)) { ClearPreview(); return; }
        var delimiter = GetDelimiter(lines);
        _detectedName = DelimiterName(delimiter);
        var parsed = lines.Select(x => SplitLine(x, delimiter).Select(v => TrimCheck.IsChecked == true ? v.Trim() : v).ToList()).ToList();
        var width = parsed.Max(r => r.Count);
        parsed.ForEach(r => { while (r.Count < width) r.Add(""); });
        if (HeaderCheck.IsChecked == true)
        {
            _headers = parsed[0].Select((h, i) => string.IsNullOrWhiteSpace(h) ? $"字段 {i + 1}" : h).ToList();
            _rows = parsed.Skip(1).ToList();
        }
        else
        {
            _headers = Enumerable.Range(1, width).Select(i => $"字段 {i}").ToList();
            _rows = parsed;
        }
        ShowPreview();
    }

    private string GetDelimiter(List<string> lines)
    {
        var tag = (DelimiterBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "auto";
        if (tag != "auto") return tag switch { "tab" => "\t", "comma" => ",", "cncomma" => "，", "pipe" => "|", "semicolon" => ";", "space" => "__SPACE__", _ => "," };
        var candidates = new[] { "\t", ",", "，", "|", ";", "__SPACE__" };
        return candidates.OrderByDescending(d => ScoreDelimiter(lines, d)).First();
    }

    private static double ScoreDelimiter(List<string> lines, string d)
    {
        var counts = lines.Take(20).Select(x => SplitLine(x, d).Count).ToList();
        if (counts.Max() < 2) return -1;
        var mode = counts.GroupBy(x => x).OrderByDescending(g => g.Count()).First();
        return mode.Count() * 4 - (counts.Max() - counts.Min()) * 2 + mode.Key * .1;
    }

    private static List<string> SplitLine(string line, string delimiter)
    {
        if (delimiter == "__SPACE__") return Regex.Split(line.Trim(), @"\s{2,}").ToList();
        if ((delimiter == "," || delimiter == ";") && line.Contains('"'))
        {
            var result = new List<string>(); var value = new StringBuilder(); var quoted = false;
            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { value.Append('"'); i++; }
                else if (c == '"') quoted = !quoted;
                else if (!quoted && c.ToString() == delimiter) { result.Add(value.ToString()); value.Clear(); }
                else value.Append(c);
            }
            result.Add(value.ToString()); return result;
        }
        return line.Split(delimiter).ToList();
    }

    private static string DelimiterName(string d) => d switch { "\t" => "Tab", "," => "逗号", "，" => "中文逗号", "|" => "竖线", ";" => "分号", "__SPACE__" => "空格", _ => d };

    private void ShowPreview()
    {
        var table = new DataTable();
        for (var i = 0; i < _headers.Count; i++)
        {
            var name = _headers[i]; var baseName = name; var n = 2;
            while (table.Columns.Contains(name)) name = $"{baseName}_{n++}";
            table.Columns.Add(name);
        }
        foreach (var row in _rows.Take(300)) table.Rows.Add(row.Cast<object>().ToArray());
        PreviewGrid.ItemsSource = table.DefaultView;
        EmptyPanel.Visibility = Visibility.Collapsed;
        StatsText.Text = $"{_rows.Count} 行 · {_headers.Count} 列 · {_detectedName}";
        ExportButton.IsEnabled = _rows.Count > 0;
        StatusText.Text = $"已识别 {_rows.Count} 行数据";
    }

    private void ClearPreview()
    {
        _headers = []; _rows = []; PreviewGrid.ItemsSource = null; EmptyPanel.Visibility = Visibility.Visible;
        StatsText.Text = "0 行 · 0 列 · 等待识别"; ExportButton.IsEnabled = false; StatusText.Text = "就绪";
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_rows.Count == 0) return;
        var name = SanitizeSheetName(SheetNameBox.Text);
        var dlg = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", FileName = $"{name}_{DateTime.Now:yyyy-MM-dd}.xlsx", DefaultExt = ".xlsx", AddExtension = true };
        if (dlg.ShowDialog() != true) return;
        try
        {
            ExportButton.IsEnabled = false; StatusText.Text = "正在生成 Excel…";
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add(name);
            for (var c = 0; c < _headers.Count; c++) ws.Cell(1, c + 1).Value = _headers[c];
            for (var r = 0; r < _rows.Count; r++)
                for (var c = 0; c < _headers.Count; c++) SetCellValue(ws.Cell(r + 2, c + 1), _rows[r][c]);
            var used = ws.Range(1, 1, _rows.Count + 1, _headers.Count);
            var header = ws.Range(1, 1, 1, _headers.Count);
            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#183653"); header.Style.Font.FontColor = XLColor.White; header.Style.Font.Bold = true; header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; header.Style.Border.BottomBorder = XLBorderStyleValues.Medium; header.Style.Border.BottomBorderColor = XLColor.FromHtml("#16865B");
            used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; used.Style.Border.InsideBorder = XLBorderStyleValues.Hair; used.Style.Border.InsideBorderColor = XLColor.FromHtml("#DCE3E8"); used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            for (var r = 3; r <= _rows.Count + 1; r += 2) ws.Range(r, 1, r, _headers.Count).Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F7F9");
            ws.SheetView.FreezeRows(1); used.SetAutoFilter(); ws.Row(1).Height = 23;
            ws.Columns(1, _headers.Count).AdjustToContents(1, Math.Min(_rows.Count + 1, 1000));
            foreach (var col in ws.Columns(1, _headers.Count)) { if (col.Width < 10) col.Width = 10; if (col.Width > 42) col.Width = 42; }
            workbook.SaveAs(dlg.FileName);
            StatusText.Text = $"已保存：{dlg.FileName}";
            MessageBox.Show($"Excel 已生成\n\n{dlg.FileName}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { MessageBox.Show($"生成失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); StatusText.Text = "生成失败"; }
        finally { ExportButton.IsEnabled = _rows.Count > 0; }
    }

    private static void SetCellValue(IXLCell cell, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) { cell.Value = ""; return; }
        if (!Regex.IsMatch(value, @"^0\d+") && double.TryParse(value, out var number)) cell.Value = number;
        else if (DateTime.TryParse(value, out var date) && Regex.IsMatch(value, @"[-/年]")) { cell.Value = date; cell.Style.DateFormat.Format = "yyyy-mm-dd"; }
        else cell.Value = value;
    }

    private static string SanitizeSheetName(string name)
    {
        var safe = Regex.Replace(string.IsNullOrWhiteSpace(name) ? "整理结果" : name, @"[\\/\*\?\:\[\]]", " ").Trim();
        return safe.Length > 31 ? safe[..31] : safe;
    }

    private void Clear_Click(object sender, RoutedEventArgs e) { _currentFile = null; InputText.Clear(); FileInfo.Text = "也可以直接在下方粘贴文本"; ClearPreview(); }
    private void Sample_Click(object sender, RoutedEventArgs e)
    {
        _currentFile = null; FileInfo.Text = "示例数据";
        InputText.Text = "订单号,客户,城市,金额,日期\nA-1001,上海启明科技,上海,12800,2026-09-21\nA-1002,杭州云帆贸易,杭州,9650,2026-09-22\nA-1003,苏州新程制造,苏州,18320,2026-09-23";
    }
}
