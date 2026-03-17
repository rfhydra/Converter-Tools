// ================================================================
//  MainWindow.xaml.cs — CashShop SQL Generator
//  Sentinel Dev
//
//  Fungsi:
//    - Baca CashShop.xlsx (kolom Code, CashCode, CsPrice)
//    - Tampilkan di DataGrid dengan preview SQL
//    - Generate INSERT INTO BILLING.dbo.tbl_cashList
//    - Simpan ke file .sql
// ================================================================
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using OfficeOpenXml;

namespace CashShopTool
{
    // ── Model ─────────────────────────────────────────────────────────
    public class CashItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        public int    RowNum     { get; set; }
        public string Code       { get; set; }  // → name
        public string CashCode   { get; set; }  // → id
        public string CsPrice    { get; set; }  // → cost
        public string PreviewSql => BuildSql();

        public string BuildSql()
        {
            string name = Code?.Trim()     ?? "";
            string id   = CashCode?.Trim() ?? "";
            string cost = CsPrice?.Trim()  ?? "0";
            return $"INSERT INTO BILLING.dbo.tbl_cashList ([name],[id],[cost]) VALUES (N'{name}', N'{id}', {cost});";
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged(string p) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
    }

    // ── MainWindow ────────────────────────────────────────────────────
    public partial class MainWindow : Window
    {
        ObservableCollection<CashItem> _items = new ObservableCollection<CashItem>();
        string _loadedFile = "";

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            dgItems.ItemsSource = _items;
            _items.CollectionChanged += (s, e) => UpdateStats();
        }

        // ── Window controls ───────────────────────────────────────────
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }
        private void btnClose_Click(object sender, RoutedEventArgs e) => Close();
        private void btnMinimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

        // ── Browse ────────────────────────────────────────────────────
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title  = "Select CashShop Excel file",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                _loadedFile        = dlg.FileName;
                txtFilePath.Text   = dlg.FileName;
                btnLoad.IsEnabled  = true;
                SetStatus("File selected. Click ⚡ Load to import.", "#C8A040");
            }
        }

        // ── Load Excel ────────────────────────────────────────────────
        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(_loadedFile))
            {
                SetStatus("File not found!", "#D72020");
                return;
            }

            try
            {
                _items.Clear();
                SetStatus("Loading...", "#C8A040");

                using (var pkg = new ExcelPackage(new FileInfo(_loadedFile)))
                {
                    var ws = pkg.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null)
                    {
                        SetStatus("No worksheet found!", "#D72020");
                        return;
                    }

                    // Cari kolom header (row 1 atau 2)
                    int headerRow = 1;
                    int colCode = -1, colCashCode = -1, colCsPrice = -1;

                    for (int row = 1; row <= Math.Min(3, ws.Dimension?.Rows ?? 1); row++)
                    {
                        for (int col = 1; col <= ws.Dimension.Columns; col++)
                        {
                            string cell = ws.Cells[row, col].Text?.Trim().ToLower() ?? "";
                            if (cell == "code")     { colCode     = col; headerRow = row; }
                            if (cell == "cashcode") { colCashCode = col; headerRow = row; }
                            if (cell == "csprice")  { colCsPrice  = col; headerRow = row; }
                        }
                        if (colCode > 0 && colCashCode > 0 && colCsPrice > 0) break;
                    }

                    // Fallback: coba kolom B, C, D jika tidak ada header
                    if (colCode < 0)     colCode     = 2;
                    if (colCashCode < 0) colCashCode = 3; // sesuai gambar: B=CashCode, C=CsPrice
                    if (colCsPrice < 0)  colCsPrice  = 4;

                    // Sesuaikan berdasarkan gambar: A=Code, B=CashCode, C=CsPrice
                    // Coba deteksi otomatis dari header row ke-2 (gambar row 2 = Code, CashCode, CsPrice)
                    int fc = FindColumn(ws, headerRow, "code");
                    int fcc = FindColumn(ws, headerRow, "cashcode");
                    int fcp = FindColumn(ws, headerRow, "csprice");
                    colCode     = fc  > 0 ? fc  : 1;
                    colCashCode = fcc > 0 ? fcc : 2;
                    colCsPrice  = fcp > 0 ? fcp : 3;

                    int dataStart = headerRow + 1;
                    int rowNum    = 0;

                    for (int row = dataStart; row <= ws.Dimension.Rows; row++)
                    {
                        string code = (ws.Cells[row, colCode].Text != null ? ws.Cells[row, colCode].Text.Trim() : "");
                        if (string.IsNullOrEmpty(code)) continue;

                        string cashCode = (ws.Cells[row, colCashCode].Text != null ? ws.Cells[row, colCashCode].Text.Trim() : "");
                        string csPrice  = ws.Cells[row, colCsPrice].Text?.Trim()  ?? "0";

                        // Bersihkan angka dari format
                        if (double.TryParse(csPrice, out double priceVal))
                            csPrice = ((long)priceVal).ToString();

                        rowNum++;
                        _items.Add(new CashItem
                        {
                            RowNum   = rowNum,
                            Code     = code,
                            CashCode = cashCode,
                            CsPrice  = csPrice,
                            IsSelected = true
                        });
                    }
                }

                UpdateStats();
                UpdatePreview();
                SetButtonsEnabled(true);
                SetStatus($"Loaded {_items.Count} items from {Path.GetFileName(_loadedFile)}", "#60C860");
            }
            catch (Exception ex)
            {
                SetStatus("Error: " + ex.Message, "#D72020");
                MessageBox.Show("Failed to load Excel file:\n\n" + ex.Message,
                    "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        int FindColumn(OfficeOpenXml.ExcelWorksheet ws, int headerRow, string name)
        {
            if (ws.Dimension == null) return -1;
            for (int col = 1; col <= ws.Dimension.Columns; col++)
            {
                string cell = ws.Cells[headerRow, col].Text?.Trim().ToLower() ?? "";
                if (cell == name.ToLower()) return col;
            }
            return null;
        }

        // ── Preview ───────────────────────────────────────────────────
        private void UpdatePreview()
        {
            var selected = _items.Where(x => x.IsSelected).ToList();
            if (!selected.Any())
            {
                txtPreview.Text = "-- No items selected";
                return;
            }

            // Show first 20 lines as preview
            int show = Math.Min(20, selected.Count);
            var sb = new StringBuilder();
            sb.AppendLine($"-- CashShop SQL — Generated by Sentinel Dev CashShop Tool");
            sb.AppendLine($"-- Total: {selected.Count} queries | {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            sb.AppendLine($"-- Target: BILLING.dbo.tbl_cashList");
            sb.AppendLine("--");
            for (int i = 0; i < show; i++)
            {
                sb.AppendLine(selected[i].BuildSql());
                sb.AppendLine("GO");
            }
            if (selected.Count > show)
                sb.AppendLine($"-- ... and {selected.Count - show} more rows");

            txtPreview.Text = sb.ToString();
        }

        private void btnPreviewAll_Click(object sender, RoutedEventArgs e) => UpdatePreview();

        private void btnCopyPreview_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtPreview.Text))
            {
                Clipboard.SetText(txtPreview.Text);
                SetStatus("Preview copied to clipboard!", "#60C860");
            }
        }

        // ── Generate SQL ──────────────────────────────────────────────
        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            GenerateSql(_items.Where(x => x.IsSelected).ToList());
        }

        private void btnGenerateAll_Click(object sender, RoutedEventArgs e)
        {
            GenerateSql(_items.ToList());
        }

        private void GenerateSql(List<CashItem> items)
        {
            if (!items.Any())
            {
                SetStatus("No items to generate!", "#D72020");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title      = "Save SQL File",
                Filter     = "SQL Files (*.sql)|*.sql|All Files (*.*)|*.*",
                FileName   = $"CashShop_{DateTime.Now:yyyyMMdd_HHmmss}.sql",
                DefaultExt = "sql"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("-- ============================================================");
                sb.AppendLine("-- CashShop SQL Insert Script");
                sb.AppendLine("-- Generated by Sentinel Dev CashShop SQL Generator");
                sb.AppendLine($"-- Date    : {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                sb.AppendLine($"-- Source  : {Path.GetFileName(_loadedFile)}");
                sb.AppendLine($"-- Total   : {items.Count} records");
                sb.AppendLine("-- Target  : BILLING.dbo.tbl_cashList");
                sb.AppendLine("-- ============================================================");
                sb.AppendLine();
                sb.AppendLine("USE BILLING;");
                sb.AppendLine("GO");
                sb.AppendLine();

                foreach (var item in items)
                {
                    sb.AppendLine(item.BuildSql());
                    sb.AppendLine("GO");
                }

                sb.AppendLine();
                sb.AppendLine($"-- Done. {items.Count} records inserted.");

                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);

                SetStatus($"✓ Saved {items.Count} queries → {Path.GetFileName(dlg.FileName)}", "#60C860");
                txtBottomStatus.Text = $"Saved: {dlg.FileName}";

                // Ask to open
                var result = MessageBox.Show(
                    $"SQL file saved!\n\nLocation:\n{dlg.FileName}\n\nTotal: {items.Count} INSERT queries\n\nOpen file?",
                    "✓ Generate Success", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(dlg.FileName);
            }
            catch (Exception ex)
            {
                SetStatus("Save failed: " + ex.Message, "#D72020");
                MessageBox.Show("Failed to save:\n\n" + ex.Message, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Select all checkbox in header ─────────────────────────────
        private void ChkAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.IsSelected = true;
            UpdateStats();
            UpdatePreview();
        }

        private void ChkAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.IsSelected = false;
            UpdateStats();
            UpdatePreview();
        }

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.IsSelected = true;
            UpdateStats(); UpdatePreview();
        }

        private void btnClearSel_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.IsSelected = false;
            UpdateStats(); UpdatePreview();
        }

        private void dgItems_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateStats();
        }

        // ── Helpers ───────────────────────────────────────────────────
        void UpdateStats()
        {
            txtRowCount.Text      = _items.Count.ToString();
            txtSelectedCount.Text = _items.Count(x => x.IsSelected).ToString();
        }

        void SetStatus(string msg, string color = "#5A7A5A")
        {
            txtStatus.Text       = msg;
            var bc = new System.Windows.Media.BrushConverter();
            txtStatus.Foreground = bc.ConvertFromString(color) as System.Windows.Media.Brush;
        }

        void SetButtonsEnabled(bool enabled)
        {
            btnPreviewAll.IsEnabled  = enabled;
            btnGenerate.IsEnabled    = enabled;
            btnGenerateAll.IsEnabled = enabled;
            btnSelectAll.IsEnabled   = enabled;
            btnClearSel.IsEnabled    = enabled;
        }
    }
}
