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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private TextBox txtFile1 = new TextBox(), txtFile2 = new TextBox();
        private ComboBox comboSheet1 = new ComboBox(), comboSheet2 = new ComboBox();
        private DataGridView grid1 = new DataGridView(), grid2 = new DataGridView();
        private Button btnCompare = new Button(), btnSyncToggle = new Button(), btnSwap = new Button(), btnIgnoreBlank = new Button();
        private Button btnBrowse1 = new Button(), btnBrowse2 = new Button();
        private Button btnSave1 = new Button(), btnSave2 = new Button();
        private Button btnSaveAs1 = new Button(), btnSaveAs2 = new Button();
        private SplitContainer splitContainer = new SplitContainer();
        private Panel navPanel = new Panel(); 

        private bool isSyncEnabled = true;
        private bool isIgnoreBlank = false;
        private bool isSyncing = false;
        private bool isResizingColumn = false;
        private List<int> diffRows = new List<int>(); 
        private int totalDiffCells = 0;

        public MainForm()
        {
            this.Width = (int)(Screen.PrimaryScreen!.Bounds.Width * 0.8);
            this.Height = (int)(Screen.PrimaryScreen!.Bounds.Height * 0.8);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel Comparator (v4.1 - 3px & Ignore Blank)";
            try { if (File.Exists("XL_compare.ico")) this.Icon = new Icon("XL_compare.ico"); } catch { }
            SetupDynamicLayout();
        }

        private void SetupDynamicLayout()
        {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 165, BackColor = Color.FromArgb(245, 245, 245), BorderStyle = BorderStyle.FixedSingle };
            int margin = 20;

            void StyleIconButton(Button btn, string imgName, string text) {
                btn.Size = new Size(85, 80);
                btn.Text = text; btn.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                btn.TextAlign = ContentAlignment.BottomCenter; btn.TextImageRelation = TextImageRelation.ImageAboveText;
                btn.BackColor = Color.White; btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderColor = Color.LightGray; btn.Cursor = Cursors.Hand;
                btn.Image = GetResourceImage(imgName, 40, 40);
            }

            void StyleBrowseIconOnly(Button btn, string imgName) {
                btn.Size = new Size(36, 26); btn.BackColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat; btn.FlatAppearance.BorderColor = Color.Silver;
                btn.Image = GetResourceImage(imgName, 22, 22);
            }

            // 按鈕佈局 (由左至右)
            StyleIconButton(btnCompare, "compare.png", "COMPARE");
            btnCompare.Location = new Point(margin, 10);
            
            StyleIconButton(btnSwap, "swap.png", "SWAP");
            btnSwap.Location = new Point(margin + 90, 10);

            StyleIconButton(btnIgnoreBlank, "ignore.png", "IGNORE BLK: OFF");
            btnIgnoreBlank.Location = new Point(margin + 180, 10);

            StyleIconButton(btnSyncToggle, "sync.png", "SYNC: ON");
            btnSyncToggle.Location = new Point(margin + 270, 10);
            btnSyncToggle.BackColor = Color.Honeydew;

            Label lbl1 = new Label { Text = "File 1:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            Label lbl2 = new Label { Text = "File 2:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            StyleBrowseIconOnly(btnBrowse1, "load_file.png");
            StyleBrowseIconOnly(btnSave1, "save_file.png");
            StyleBrowseIconOnly(btnSaveAs1, "save_as.png");
            Label lblS1 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };

            //Label lbl2 = new Label { Text = "File 2:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            StyleBrowseIconOnly(btnBrowse2, "load_file.png");
            StyleBrowseIconOnly(btnSave2, "save_file.png");
            StyleBrowseIconOnly(btnSaveAs2, "save_as.png");
            Label lblS2 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };

            txtFile1.ReadOnly = txtFile2.ReadOnly = true;
            comboSheet1.DropDownStyle = comboSheet2.DropDownStyle = ComboBoxStyle.DropDownList;

            // 設定 ToolTip
            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(btnBrowse1, "Open Excel File");
            toolTip.SetToolTip(btnSave1, "Save File");
            toolTip.SetToolTip(btnSaveAs1, "Save As");
            toolTip.SetToolTip(btnBrowse2, "Open Excel File");
            toolTip.SetToolTip(btnSave2, "Save File");
            toolTip.SetToolTip(btnSaveAs2, "Save As");

            // 導覽列與 3px 間距
            navPanel.Dock = DockStyle.Right; navPanel.Width = 35; navPanel.BackColor = Color.FromArgb(30, 30, 30);
            navPanel.Paint += NavPanel_Paint; navPanel.MouseClick += NavPanel_MouseClick;
            Panel rightMargin = new Panel { Dock = DockStyle.Right, Width = 3 }; 
            Panel middleSpacer = new Panel { Dock = DockStyle.Right, Width = 3 };

            EnableDragDrop(grid1, txtFile1, comboSheet1);
            EnableDragDrop(txtFile1, txtFile1, comboSheet1);
            EnableDragDrop(grid2, txtFile2, comboSheet2);
            EnableDragDrop(txtFile2, txtFile2, comboSheet2);

            //topPanel.Controls.AddRange(new Control[] { btnCompare, btnSwap, btnIgnoreBlank, btnSyncToggle, lbl1, txtFile1, btnBrowse1, btnSave1, comboSheet1, lbl2, txtFile2, btnBrowse2, btnSave2, comboSheet2 });

            topPanel.Controls.AddRange(new Control[] { 
                btnCompare, btnSwap, btnIgnoreBlank, btnSyncToggle, lbl1, txtFile1, btnBrowse1, btnSave1, btnSaveAs1, lblS1, comboSheet1, 
                lbl2, txtFile2, btnBrowse2, btnSave2, btnSaveAs2, lblS2, comboSheet2 
            });

            splitContainer = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BorderStyle = BorderStyle.Fixed3D };
            ConfigGrid(grid1); ConfigGrid(grid2);
            splitContainer.Panel1.Controls.Add(grid1); splitContainer.Panel2.Controls.Add(grid2);

            // 绑定滚动同步事件
            grid1.Scroll += (s, e) => SyncGrids(grid1, grid2, e);
            grid2.Scroll += (s, e) => SyncGrids(grid2, grid1, e);
            
            // 绑定列宽变化同步事件
            grid1.ColumnWidthChanged += (s, e) => {
                if (!isResizingColumn && isSyncEnabled) {
                    isResizingColumn = true;
                    try {
                        for (int i = 0; i < Math.Min(grid1.ColumnCount, grid2.ColumnCount); i++) {
                            grid2.Columns[i].Width = grid1.Columns[i].Width;
                        }
                    } finally {
                        isResizingColumn = false;
                    }
                }
            };
            grid2.ColumnWidthChanged += (s, e) => {
                if (!isResizingColumn && isSyncEnabled) {
                    isResizingColumn = true;
                    try {
                        for (int i = 0; i < Math.Min(grid1.ColumnCount, grid2.ColumnCount); i++) {
                            grid1.Columns[i].Width = grid2.Columns[i].Width;
                        }
                    } finally {
                        isResizingColumn = false;
                    }
                }
            };

            this.Controls.Add(splitContainer); this.Controls.Add(middleSpacer); this.Controls.Add(navPanel); this.Controls.Add(rightMargin); this.Controls.Add(topPanel);

            // 事件繫結
            btnIgnoreBlank.Click += (s, e) => {
                isIgnoreBlank = !isIgnoreBlank;
                btnIgnoreBlank.Text = isIgnoreBlank ? "IGNORE BLK: ON" : "IGNORE BLK: OFF";
                btnIgnoreBlank.BackColor = isIgnoreBlank ? Color.NavajoWhite : Color.White;
            };
            btnSyncToggle.Click += (s, e) => {
                isSyncEnabled = !isSyncEnabled;
                btnSyncToggle.Text = isSyncEnabled ? "SYNC: ON" : "SYNC: OFF";
                btnSyncToggle.BackColor = isSyncEnabled ? Color.Honeydew : Color.LightGray;
            };
            btnBrowse1.Click += (s, e) => LoadExcel(txtFile1, comboSheet1);
            btnBrowse2.Click += (s, e) => LoadExcel(txtFile2, comboSheet2);
            btnCompare.Click += RunComparison;
            btnSwap.Click += SwapGrids;
            btnSave1.Click += (s, e) => SaveExcel(txtFile1.Text, comboSheet1.Text, grid1);
            btnSave2.Click += (s, e) => SaveExcel(txtFile2.Text, comboSheet2.Text, grid2);
            btnSaveAs1.Click += (s, e) => SaveExcelAs(txtFile1, comboSheet1, grid1);
            btnSaveAs2.Click += (s, e) => SaveExcelAs(txtFile2, comboSheet2, grid2);
            comboSheet1.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile1.Text, comboSheet1.Text, grid1);
            comboSheet2.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile2.Text, comboSheet2.Text, grid2);

            this.Resize += (s, e) => {
                if (this.WindowState == FormWindowState.Minimized) return;
                int midX = (topPanel.Width - 50) / 2;
                lbl1.Location = new Point(margin, 100);
                txtFile1.Bounds = new Rectangle(margin + 50, 100, midX - 250, 25);
                btnBrowse1.Location = new Point(txtFile1.Right + 3, 99);
                btnSave1.Location = new Point(btnBrowse1.Right + 3, 99);
                btnSaveAs1.Location = new Point(btnSave1.Right + 3, 99);
                lblS1.Location = new Point(margin, 130);
                comboSheet1.Bounds = new Rectangle(margin + 50, 130, 120, 25);
                
                lbl2.Location = new Point(midX + 20, 100);
                txtFile2.Bounds = new Rectangle(midX + 70, 100, midX - 250, 25);
                btnBrowse2.Location = new Point(txtFile2.Right + 3, 99);
                btnSave2.Location = new Point(btnBrowse2.Right + 3, 99);
                btnSaveAs2.Location = new Point(btnSave2.Right + 3, 99);
                lblS2.Location = new Point(midX + margin, 130);
                comboSheet2.Bounds = new Rectangle(midX + 70, 130, 120, 25);
                if (midX > 0) splitContainer.SplitterDistance = midX;
            };
            this.OnResize(EventArgs.Empty);
        }

        private async void RunComparison(object? sender, EventArgs e) {
            if (grid1.DataSource == null || grid2.DataSource == null) return;
            btnCompare.Enabled = false; btnCompare.Text = "WAIT...";
            DataTable dt1 = (DataTable)grid1.DataSource;
            DataTable dt2 = (DataTable)grid2.DataSource;

            var result = await Task.Run(() => {
                var diffs = new List<Point>();
                var rows = new HashSet<int>();
                int maxR = Math.Max(dt1.Rows.Count, dt2.Rows.Count);
                int maxC = Math.Max(dt1.Columns.Count, dt2.Columns.Count);
                for (int r = 0; r < maxR; r++) {
                    bool rowDiff = false;
                    for (int c = 0; c < maxC; c++) {
                        string v1 = (r < dt1.Rows.Count && c < dt1.Columns.Count) ? dt1.Rows[r][c]?.ToString() ?? "" : "[NULL]";
                        string v2 = (r < dt2.Rows.Count && c < dt2.Columns.Count) ? dt2.Rows[r][c]?.ToString() ?? "" : "[NULL]";
                        if (isIgnoreBlank) { v1 = v1.Trim(); v2 = v2.Trim(); }
                        if (v1 != v2) { diffs.Add(new Point(c, r)); rowDiff = true; }
                    }
                    if (rowDiff) rows.Add(r);
                }
                return new { Points = diffs, Rows = rows.OrderBy(x => x).ToList() };
            });

            grid1.SuspendLayout(); grid2.SuspendLayout();
            ResetGridStyle(grid1); ResetGridStyle(grid2);
            foreach (var p in result.Points) {
                HighlightDiffCell(grid1, p.X, p.Y);
                HighlightDiffCell(grid2, p.X, p.Y);
            }
            diffRows = result.Rows; totalDiffCells = result.Points.Count;
            grid1.ResumeLayout(true); grid2.ResumeLayout(true);
            grid1.Refresh(); grid2.Refresh();
            navPanel.Invalidate(); btnCompare.Enabled = true; btnCompare.Text = "COMPARE";
            MessageBox.Show($"差異列數: {diffRows.Count}");
        }

        private void NavPanel_Paint(object? sender, PaintEventArgs e) {
            if (grid1.RowCount == 0) return;
            Graphics g = e.Graphics;
            using (Pen p = new Pen(Color.DodgerBlue, 3)) {
                Rectangle r = navPanel.ClientRectangle; r.Inflate(-1, -1);
                g.DrawRectangle(p, r);
            }
            float ratio = (float)navPanel.Height / grid1.RowCount;
            int first = grid1.FirstDisplayedScrollingRowIndex;
            if (first >= 0) {
                using (SolidBrush b = new SolidBrush(Color.FromArgb(100, 30, 144, 255)))
                    g.FillRectangle(b, 5, first * ratio, navPanel.Width - 10, Math.Max(grid1.DisplayedRowCount(false) * ratio, 5));
            }
            using (Pen orange = new Pen(Color.Orange, 3)) {
                foreach (int row in diffRows) {
                    float y = row * ratio;
                    g.DrawLine(orange, 8, y, navPanel.Width - 8, y);
                }
            }
        }

        private void NavPanel_MouseClick(object? sender, MouseEventArgs e) {
            if (grid1.RowCount == 0) return;
            int row = (int)(e.Y / (float)navPanel.Height * grid1.RowCount);
            if (row >= 0 && row < grid1.RowCount) grid1.FirstDisplayedScrollingRowIndex = row;
        }

        private void EnableDragDrop(Control ctrl, TextBox t, ComboBox c) {
            ctrl.AllowDrop = true;
            ctrl.DragEnter += (s, e) => { if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            ctrl.DragDrop += (s, e) => {
                if (e.Data == null) return;
                string[]? files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0) { t.Text = files[0]; LoadSheetsAfterDrop(t.Text, c); }
            };
        }

        private void LoadExcel(TextBox t, ComboBox c) {
            using OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls" };
            if (ofd.ShowDialog() == DialogResult.OK) {
                t.Text = ofd.FileName;
                LoadSheetsAfterDrop(t.Text, c);
            }
        }

        private void LoadSheetsAfterDrop(string path, ComboBox c) {
            try {
                c.Items.Clear();
                using (var wb = new XLWorkbook(path)) {
                    foreach (var s in wb.Worksheets) c.Items.Add(s.Name);
                    if (c.Items.Count > 0) c.SelectedIndex = 0;
                }
            } catch { }
        }

        private void ConfigGrid(DataGridView g) {
            g.Dock = DockStyle.Fill; g.BackgroundColor = Color.White; g.RowHeadersVisible = false;
            g.AllowUserToAddRows = false; g.EnableHeadersVisualStyles = false;
            g.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy; g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(g, true);
            g.Scroll += (s, e) => { if (isSyncEnabled && !isSyncing) navPanel.Invalidate(); };
        }

        private void DisplaySheet(string p, string s, DataGridView g) {
            if (string.IsNullOrEmpty(p) || string.IsNullOrEmpty(s)) return;
            DataTable dt = new DataTable();
            try {
                using (var wb = new XLWorkbook(p)) {
                    var ws = wb.Worksheet(s);
                    var range = ws.RangeUsed();
                    if (range == null) return;
//                    for (int i = 1; i <= range.LastColumn().ColumnNumber(); i++) dt.Columns.Add(ws.Cell(1, i).Value?.ToString() ?? "Col" + i);
                    for (int i = 1; i <= range.LastColumn().ColumnNumber(); i++) dt.Columns.Add(ws.Cell(1, i).Value.ToString() ?? "Col" + i);
                    foreach (var r in ws.RowsUsed().Skip(1)) {
                        DataRow dr = dt.NewRow();
//                        for (int i = 0; i < dt.Columns.Count; i++) dr[i] = r.Cell(i + 1).Value?.ToString() ?? "";
                        for (int i = 0; i < dt.Columns.Count; i++) dr[i] = r.Cell(i + 1).Value.ToString() ?? "";
                        dt.Rows.Add(dr);
                    }
                }
                g.DataSource = dt; g.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                diffRows.Clear(); navPanel.Invalidate();
            } catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }

        private void SaveExcel(string filePath, string sheetName, DataGridView g) {
            if (string.IsNullOrEmpty(filePath) || g.DataSource == null) return;
            g.EndEdit();
            DataTable dt = (DataTable)g.DataSource;
            try {
                using (var wb = new XLWorkbook(filePath)) {
                    var ws = wb.Worksheet(sheetName);
                    ws.Clear(XLClearOptions.Contents);
                    for (int i = 0; i < dt.Columns.Count; i++) ws.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                    for (int r = 0; r < dt.Rows.Count; r++) {
                        for (int c = 0; c < dt.Columns.Count; c++) ws.Cell(r + 2, c + 1).Value = dt.Rows[r][c]?.ToString() ?? "";
                    }
                    wb.Save(); MessageBox.Show("存檔成功。");
                    ResetGridStyle(g);
                }
            } catch (Exception ex) { MessageBox.Show("Save Error: " + ex.Message); }
        }

        private void SaveExcelAs(TextBox t, ComboBox c, DataGridView g) {
            if (g.DataSource == null) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel|*.xlsx|Excel 2007|*.xlsx|Excel 97-2003|*.xls" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    string newPath = sfd.FileName;
                    g.EndEdit();
                    DataTable dt = (DataTable)g.DataSource;
                    try {
                        using (var wb = new XLWorkbook()) {
                            var ws = wb.Worksheets.Add(c.Text);
                            for (int i = 0; i < dt.Columns.Count; i++) ws.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                            for (int r = 0; r < dt.Rows.Count; r++) {
                                for (int col = 0; col < dt.Columns.Count; col++) ws.Cell(r + 2, col + 1).Value = dt.Rows[r][col]?.ToString() ?? "";
                            }
                            wb.SaveAs(newPath); 
                            t.Text = newPath;
                            MessageBox.Show($"檔案已另存為:\n{newPath}");
                            ResetGridStyle(g);
                        }
                    } catch (Exception ex) { MessageBox.Show("Save As Error: " + ex.Message); }
                }
            }
        }

        private void ResetGridStyle(DataGridView g) {
            foreach (DataGridViewRow r in g.Rows) {
                foreach (DataGridViewCell c in r.Cells) {
                    c.Style.BackColor = Color.White;
                    c.Style.ForeColor = Color.Black;
                }
            }
        }

        private void HighlightDiffCell(DataGridView g, int col, int row) {
            if (row < g.RowCount && col < g.ColumnCount) {
                DataGridViewCell cell = g[col, row];
                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.BackColor = Color.MistyRose;
                style.ForeColor = Color.Black;
                cell.Style = style;
                g.InvalidateCell(col, row);
            }
        }

        private void SwapGrids(object? sender, EventArgs e) {
            isResizingColumn = true;
            string t = txtFile1.Text; txtFile1.Text = txtFile2.Text; txtFile2.Text = t;
            var d = grid1.DataSource; grid1.DataSource = grid2.DataSource; grid2.DataSource = d;
            isResizingColumn = false; ResetGridStyle(grid1); ResetGridStyle(grid2);
        }

        private void SyncColumnWidth(DataGridView s, DataGridView t, DataGridViewColumnEventArgs e) {
            if (!isSyncEnabled || isResizingColumn) return;
            isResizingColumn = true;
            try {
                if (e.Column != null && e.Column.Index >= 0 && e.Column.Index < t.ColumnCount) {
                    t.Columns[e.Column.Index].Width = e.Column.Width;
                }
            } finally {
                isResizingColumn = false;
            }
        }

        private void SyncGrids(DataGridView s, DataGridView t, ScrollEventArgs e) {
            if (!isSyncEnabled || isSyncing || t.RowCount == 0) return;
            isSyncing = true;
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) {
                if (s.FirstDisplayedScrollingRowIndex >= 0 && s.FirstDisplayedScrollingRowIndex < t.RowCount)
                    t.FirstDisplayedScrollingRowIndex = s.FirstDisplayedScrollingRowIndex;
            } else if (s.HorizontalScrollingOffset >= 0) {
                t.HorizontalScrollingOffset = s.HorizontalScrollingOffset;
            }
            isSyncing = false;
        }

        private Image? GetResourceImage(string name, int w, int h) {
            try {
                var assembly = Assembly.GetExecutingAssembly();
                string? resName = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith(name));
                if (resName != null) {
                    using (Stream? s = assembly.GetManifestResourceStream(resName)) {
                        if (s == null) return null;
                        using (Bitmap bmp = new Bitmap(s)) {
                            Bitmap res = new Bitmap(w, h);
                            using (Graphics g = Graphics.FromImage(res)) {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.DrawImage(bmp, 0, 0, w, h);
                            }
                            return res;
                        }
                    }
                }
            } catch { }
            return null;
        }
    }
}