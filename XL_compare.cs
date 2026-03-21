using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace XLCompareApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        // UI 控制項
        private TextBox txtFile1 = new(), txtFile2 = new();
        private ComboBox comboSheet1 = new(), comboSheet2 = new();
        private DataGridView grid1 = new(), grid2 = new();
        private Button btnCompare = new(), btnSyncToggle = new(), btnSwap = new();
        private Button btnBrowse1 = new(), btnBrowse2 = new();
        private Button btnSave1 = new(), btnSave2 = new();
        private SplitContainer splitContainer = new();
        private Panel navPanel = new(); // WinMerge 風格導航條

        // 狀態變數
        private bool isSyncEnabled = true;
        private bool isSyncing = false;
        private bool isResizingColumn = false;
        private List<int> diffRows = new(); // 紀錄有差異的行號

        public MainForm()
        {
            this.Width = (int)(Screen.PrimaryScreen!.Bounds.Width * 0.8);
            this.Height = (int)(Screen.PrimaryScreen!.Bounds.Height * 0.8);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel Professional Comparator (WinMerge Style)";
            SetupDynamicLayout();
        }

        private void SetupDynamicLayout()
        {
            // --- 頂部面板設定 ---
            Panel topPanel = new Panel {
                Dock = DockStyle.Top, Height = 165,
                BackColor = Color.FromArgb(245, 245, 245), BorderStyle = BorderStyle.FixedSingle
            };

            int margin = 20;

            // 輔助函式：取得資源圖示
            Image? GetResourceImage(string fileName, int w, int h) {
                try {
                    var assembly = Assembly.GetExecutingAssembly();
                    string? targetKey = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase));
                    if (targetKey != null) {
                        using Stream? stream = assembly.GetManifestResourceStream(targetKey);
                        if (stream != null) {
                            Bitmap resized = new Bitmap(w, h);
                            using (Graphics g = Graphics.FromImage(resized)) {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.DrawImage(Image.FromStream(stream), 0, 0, w, h);
                            }
                            return resized;
                        }
                    }
                } catch { }
                return null;
            }

            // 按鈕樣式
            void StyleIconButton(Button btn, string imgName, string text) {
                btn.Size = new Size(80, 80);
                btn.Text = text; btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                btn.TextAlign = ContentAlignment.BottomCenter; btn.TextImageRelation = TextImageRelation.ImageAboveText;
                btn.BackColor = Color.White; btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderColor = Color.LightGray; btn.Cursor = Cursors.Hand;
                btn.Image = GetResourceImage(imgName, 45, 45);
            }

            void StyleBrowseIconOnly(Button btn, string imgName) {
                btn.Size = new Size(36, 26); btn.BackColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat; btn.FlatAppearance.BorderColor = Color.Silver;
                btn.Cursor = Cursors.Hand; btn.Image = GetResourceImage(imgName, 18, 18);
            }

            StyleIconButton(btnCompare, "compare.png", "COMPARE");
            btnCompare.Location = new Point(margin, 10);
            StyleIconButton(btnSwap, "swap.png", "SWAP");
            btnSwap.Location = new Point(margin + 90, 10);
            StyleIconButton(btnSyncToggle, "sync.png", "SYNC: ON");
            btnSyncToggle.Location = new Point(margin + 180, 10);
            btnSyncToggle.BackColor = Color.Honeydew;

            Label lbl1 = new Label { Text = "File 1:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            StyleBrowseIconOnly(btnBrowse1, "load_file.png");
            StyleBrowseIconOnly(btnSave1, "save_file.png");
            Label lblS1 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };

            Label lbl2 = new Label { Text = "File 2:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            StyleBrowseIconOnly(btnBrowse2, "load_file.png");
            StyleBrowseIconOnly(btnSave2, "save_file.png");
            Label lblS2 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };

            txtFile1.ReadOnly = txtFile2.ReadOnly = true;
            comboSheet1.Width = comboSheet2.Width = 120;
            comboSheet1.DropDownStyle = comboSheet2.DropDownStyle = ComboBoxStyle.DropDownList;

            // --- 導航列設定 (WinMerge Style) ---
            navPanel.Dock = DockStyle.Right;
            navPanel.Width = 25;
            navPanel.BackColor = Color.FromArgb(23, 23, 23);
            navPanel.Cursor = Cursors.Hand;
            navPanel.Paint += NavPanel_Paint;
            navPanel.MouseClick += NavPanel_MouseClick;

            // 響應式佈局
            this.Resize += (s, e) => {
                if (this.WindowState == FormWindowState.Minimized) return;
                int midX = (topPanel.ClientSize.Width - navPanel.Width) / 2;
                int inputWidth = Math.Max(100, midX - 200);

                lbl1.Location = new Point(margin, 100);
                txtFile1.Bounds = new Rectangle(margin + 50, 100, inputWidth, 25);
                btnBrowse1.Location = new Point(txtFile1.Right + 3, 99);
                btnSave1.Location = new Point(btnBrowse1.Right + 3, 99);
                lblS1.Location = new Point(margin, 130);
                comboSheet1.Location = new Point(margin + 50, 130);

                lbl2.Location = new Point(midX + margin, 100);
                txtFile2.Bounds = new Rectangle(midX + margin + 50, 100, inputWidth, 25);
                btnBrowse2.Location = new Point(txtFile2.Right + 3, 99);
                btnSave2.Location = new Point(btnBrowse2.Right + 3, 99);
                lblS2.Location = new Point(midX + margin, 130);
                comboSheet2.Location = new Point(midX + margin + 50, 130);

                splitContainer.SplitterDistance = midX;
                navPanel.Invalidate();
            };

            // 事件綁定
            btnBrowse1.Click += (s, e) => LoadExcel(txtFile1, comboSheet1);
            btnBrowse2.Click += (s, e) => LoadExcel(txtFile2, comboSheet2);
            btnSave1.Click += (s, e) => SaveExcel(txtFile1.Text, comboSheet1.Text, grid1);
            btnSave2.Click += (s, e) => SaveExcel(txtFile2.Text, comboSheet2.Text, grid2);
            comboSheet1.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile1.Text, comboSheet1.Text, grid1);
            comboSheet2.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile2.Text, comboSheet2.Text, grid2);
            btnCompare.Click += RunComparison;
            btnSyncToggle.Click += (s, e) => {
                isSyncEnabled = !isSyncEnabled;
                btnSyncToggle.Text = isSyncEnabled ? "SYNC: ON" : "SYNC: OFF";
                btnSyncToggle.BackColor = isSyncEnabled ? Color.Honeydew : Color.LightGray;
            };
            btnSwap.Click += SwapGrids;

            topPanel.Controls.AddRange(new Control[] { 
                btnCompare, btnSwap, btnSyncToggle, lbl1, txtFile1, btnBrowse1, btnSave1, lblS1, comboSheet1, 
                lbl2, txtFile2, btnBrowse2, btnSave2, lblS2, comboSheet2 
            });

            splitContainer = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BorderStyle = BorderStyle.Fixed3D };
            ConfigGrid(grid1); ConfigGrid(grid2);
            
            grid1.ColumnWidthChanged += (s, e) => SyncColumnWidth(grid1, grid2, e);
            grid2.ColumnWidthChanged += (s, e) => SyncColumnWidth(grid2, grid1, e);
            grid1.Scroll += (s, e) => { if (isSyncEnabled) SyncGrids(grid1, grid2, e); navPanel.Invalidate(); };
            grid2.Scroll += (s, e) => { if (isSyncEnabled) SyncGrids(grid2, grid1, e); navPanel.Invalidate(); };

            splitContainer.Panel1.Controls.Add(grid1);
            splitContainer.Panel2.Controls.Add(grid2);
            
            this.Controls.Add(splitContainer);
            this.Controls.Add(navPanel);
            this.Controls.Add(topPanel);
            this.OnResize(EventArgs.Empty);
        }

        private void ConfigGrid(DataGridView g) { 
            g.Dock = DockStyle.Fill; g.BackgroundColor = Color.White; 
            g.ReadOnly = true; g.RowHeadersVisible = false; 
            g.AllowUserToAddRows = false; g.BorderStyle = BorderStyle.None;
            g.EnableHeadersVisualStyles = false;
            g.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(g, true);
        }

        // --- 核心邏輯：比對 (非同步 + 行列不一致處理) ---
        private async void RunComparison(object? sender, EventArgs e) {
            if (grid1.DataSource == null || grid2.DataSource == null) return;
            
            btnCompare.Enabled = false;
            btnCompare.Text = "WAIT...";
            diffRows.Clear();

            DataTable dt1 = (DataTable)grid1.DataSource;
            DataTable dt2 = (DataTable)grid2.DataSource;

            var result = await Task.Run(() => {
                var diffs = new List<Point>();
                int maxR = Math.Max(dt1.Rows.Count, dt2.Rows.Count);
                int maxC = Math.Max(dt1.Columns.Count, dt2.Columns.Count);
                var rowsWithDiff = new HashSet<int>();

                for (int r = 0; r < maxR; r++) {
                    bool rowDiff = false;
                    for (int c = 0; c < maxC; c++) {
                        string v1 = (r < dt1.Rows.Count && c < dt1.Columns.Count) ? dt1.Rows[r][c]?.ToString() ?? "" : "[NULL]";
                        string v2 = (r < dt2.Rows.Count && c < dt2.Columns.Count) ? dt2.Rows[r][c]?.ToString() ?? "" : "[NULL]";
                        if (v1 != v2) {
                            diffs.Add(new Point(c, r));
                            rowDiff = true;
                        }
                    }
                    if (rowDiff) rowsWithDiff.Add(r);
                }
                return new { Points = diffs, Rows = rowsWithDiff.OrderBy(x => x).ToList() };
            });

            grid1.SuspendLayout(); grid2.SuspendLayout();
            ResetGridStyle(grid1); ResetGridStyle(grid2);

            foreach (var p in result.Points) {
                if (p.Y < grid1.RowCount && p.X < grid1.ColumnCount) grid1[p.X, p.Y].Style.BackColor = Color.MistyRose;
                if (p.Y < grid2.RowCount && p.X < grid2.ColumnCount) grid2[p.X, p.Y].Style.BackColor = Color.MistyRose;
            }

            diffRows = result.Rows;
            grid1.ResumeLayout(); grid2.ResumeLayout();
            navPanel.Invalidate();
            btnCompare.Enabled = true; btnCompare.Text = "COMPARE";
            MessageBox.Show($"Found {diffRows.Count} different rows.", "Done");
        }

        // --- 導航列繪製邏輯 ---
        private void NavPanel_Paint(object? sender, PaintEventArgs e) {
            if (grid1.RowCount == 0) return;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // 畫差異點
            using var p = new Pen(Color.Orange, 2);
            foreach (int row in diffRows) {
                float y = (float)row / grid1.RowCount * navPanel.Height;
                e.Graphics.DrawLine(p, 0, y, navPanel.Width, y);
            }

            // 畫目前視窗位置
            int visibleRows = grid1.DisplayedRowCount(false);
            if (visibleRows > 0) {
                float viewY = (float)grid1.FirstDisplayedScrollingRowIndex / grid1.RowCount * navPanel.Height;
                float viewH = (float)visibleRows / grid1.RowCount * navPanel.Height;
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(80, Color.Gray)), 0, viewY, navPanel.Width, Math.Max(5, viewH));
            }
        }

        private void NavPanel_MouseClick(object? sender, MouseEventArgs e) {
            if (grid1.RowCount == 0) return;
            float pct = (float)e.Y / navPanel.Height;
            int row = (int)(pct * grid1.RowCount);
            if (row >= 0 && row < grid1.RowCount) grid1.FirstDisplayedScrollingRowIndex = row;
        }

        // --- 輔助功能 ---
        private void SwapGrids(object? sender, EventArgs e) {
            isResizingColumn = true;
            var tempFile = txtFile1.Text; txtFile1.Text = txtFile2.Text; txtFile2.Text = tempFile;
            var tempItems = comboSheet1.Items.Cast<object>().ToArray();
            var tempIdx = comboSheet1.SelectedIndex;
            comboSheet1.Items.Clear(); comboSheet1.Items.AddRange(comboSheet2.Items.Cast<object>().ToArray());
            comboSheet2.Items.Clear(); comboSheet2.Items.AddRange(tempItems);
            var ds1 = grid1.DataSource; grid1.DataSource = grid2.DataSource; grid2.DataSource = ds1;
            isResizingColumn = false;
            ResetGridStyle(grid1); ResetGridStyle(grid2);
            navPanel.Invalidate();
        }

        private void SyncColumnWidth(DataGridView src, DataGridView tar, DataGridViewColumnEventArgs e) {
            if (!isSyncEnabled || isResizingColumn) return;
            isResizingColumn = true;
            try { if (e.Column.Index < tar.ColumnCount) tar.Columns[e.Column.Index].Width = e.Column.Width; }
            finally { isResizingColumn = false; }
        }

        private void SyncGrids(DataGridView src, DataGridView tar, ScrollEventArgs e) {
            if (isSyncing || tar.RowCount == 0) return;
            isSyncing = true;
            try {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) {
                    if (src.FirstDisplayedScrollingRowIndex >= 0 && src.FirstDisplayedScrollingRowIndex < tar.RowCount)
                        tar.FirstDisplayedScrollingRowIndex = src.FirstDisplayedScrollingRowIndex;
                } else tar.HorizontalScrollingOffset = src.HorizontalScrollingOffset;
            } catch { }
            isSyncing = false;
        }

        private void LoadExcel(TextBox t, ComboBox c) {
            using OpenFileDialog ofd = new() { Filter = "Excel|*.xlsx;*.xls" };
            if (ofd.ShowDialog() == DialogResult.OK) {
                t.Text = ofd.FileName; c.Items.Clear();
                try {
                    using var wb = new XLWorkbook(t.Text);
                    foreach (var s in wb.Worksheets) c.Items.Add(s.Name);
                    if (c.Items.Count > 0) c.SelectedIndex = 0;
                } catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private void DisplaySheet(string p, string s, DataGridView g) {
            if (string.IsNullOrEmpty(p) || string.IsNullOrEmpty(s)) return;
            DataTable dt = new();
            try {
                using var wb = new XLWorkbook(p);
                var ws = wb.Worksheet(s);
                var range = ws.RangeUsed();
                if (range == null) return;
                int lc = range.LastColumn().ColumnNumber();
                for (int i = 1; i <= lc; i++) dt.Columns.Add(ws.Cell(1, i).Value.ToString() ?? "Col" + i);
                foreach (var r in ws.RowsUsed().Skip(1)) {
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count; i++) dr[i] = r.Cell(i + 1).Value.ToString();
                    dt.Rows.Add(dr);
                }
                g.DataSource = dt;
                g.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                diffRows.Clear(); navPanel.Invalidate();
            } catch { }
        }

        private void SaveExcel(string filePath, string sheetName, DataGridView g) {
            if (string.IsNullOrEmpty(filePath) || g.DataSource == null) return;
            try {
                using var wb = new XLWorkbook(filePath);
                var ws = wb.Worksheet(sheetName);
                DataTable dt = (DataTable)g.DataSource;
                for (int i = 0; i < dt.Columns.Count; i++) ws.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                for (int r = 0; r < dt.Rows.Count; r++)
                    for (int c = 0; c < dt.Columns.Count; c++)
                        ws.Cell(r + 2, c + 1).Value = dt.Rows[r][c]?.ToString();
                wb.Save();
                MessageBox.Show("Saved!");
            } catch (Exception ex) { MessageBox.Show("Failed: " + ex.Message); }
        }

        private void ResetGridStyle(DataGridView g) {
            if (g.DataSource == null) return;
            foreach (DataGridViewRow row in g.Rows)
                foreach (DataGridViewCell cell in row.Cells) cell.Style.BackColor = Color.White;
        }
    }
}