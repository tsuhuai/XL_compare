using System;
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
        private TextBox txtFile1 = new(), txtFile2 = new();
        private ComboBox comboSheet1 = new(), comboSheet2 = new();
        private DataGridView grid1 = new(), grid2 = new();
        private Button btnCompare = new(), btnSyncToggle = new(), btnSwap = new();
        private Button btnBrowse1 = new(), btnBrowse2 = new();
        private SplitContainer splitContainer = new();

        private bool isSyncEnabled = true;
        private bool isSyncing = false;
        private bool isResizingColumn = false;

        public MainForm()
        {
            this.Width = (int)(Screen.PrimaryScreen!.Bounds.Width * 0.8);
            this.Height = (int)(Screen.PrimaryScreen!.Bounds.Height * 0.8);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel Professional Comparator (20260320)";
            SetupDynamicLayout();
        }

        private void SetupDynamicLayout()
        {
            Panel topPanel = new Panel {
                Dock = DockStyle.Top, Height = 180,
                BackColor = Color.FromArgb(245, 245, 245), BorderStyle = BorderStyle.FixedSingle
            };

            int margin = 20;

            Image? GetResourceImage(string fileName, int w, int h) {
                try {
                    var assembly = Assembly.GetExecutingAssembly();
                    var resNames = assembly.GetManifestResourceNames();
                    string? targetKey = resNames.FirstOrDefault(r => r.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase));
                    if (targetKey != null) {
                        using (Stream? stream = assembly.GetManifestResourceStream(targetKey)) {
                            if (stream != null) {
                                Image original = Image.FromStream(stream);
                                Bitmap resized = new Bitmap(w, h);
                                using (Graphics g = Graphics.FromImage(resized)) {
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                    g.DrawImage(original, 0, 0, w, h);
                                }
                                return resized;
                            }
                        }
                    }
                } catch { }
                return null;
            }

            void StyleIconButton(Button btn, string imgName, string text) {
                btn.Size = new Size(80, 80);
                btn.Text = text;
                btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                btn.TextAlign = ContentAlignment.BottomCenter;
                btn.TextImageRelation = TextImageRelation.ImageAboveText;
                btn.BackColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderColor = Color.LightGray;
                btn.Cursor = Cursors.Hand;
                btn.Image = GetResourceImage(imgName, 45, 45);
            }

            // 修改：純圖示 Browse 按鈕樣式
            void StyleBrowseIconOnly(Button btn, string imgName) {
                btn.Size = new Size(36, 24); // 正方形或接近正方形
                btn.Text = ""; // 不需要文字
                btn.BackColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderColor = Color.Silver;
                btn.Cursor = Cursors.Hand;
                btn.Image = GetResourceImage(imgName, 18, 18);
            }

            StyleIconButton(btnCompare, "compare.png", "COMPARE");
            btnCompare.Location = new Point(margin, 10);

            StyleIconButton(btnSwap, "swap.png", "SWAP");
            btnSwap.Location = new Point(margin + 90, 10);

            StyleIconButton(btnSyncToggle, "sync.png", "SYNC: ON");
            btnSyncToggle.Location = new Point(margin + 180, 10);
            btnSyncToggle.BackColor = Color.Honeydew;

            Label lbl1 = new Label { Text = "File 1:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            Label lblS1 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };
            StyleBrowseIconOnly(btnBrowse1, "load_file.png");

            Label lbl2 = new Label { Text = "File 2:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            Label lblS2 = new Label { Text = "Sheet:", AutoSize = true, ForeColor = Color.DimGray };
            StyleBrowseIconOnly(btnBrowse2, "load_file.png");

            txtFile1.ReadOnly = txtFile2.ReadOnly = true;
            comboSheet1.Width = 120; comboSheet1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboSheet2.Width = 120; comboSheet2.DropDownStyle = ComboBoxStyle.DropDownList;

            this.Resize += (s, e) => {
                if (this.WindowState == FormWindowState.Minimized) return;
                int midX = topPanel.ClientSize.Width / 2;
                // 因為按鈕變窄了，inputWidth 可以稍微加寬
                int inputWidth = midX - 120; 
                if (inputWidth < 100) inputWidth = 100;

                lbl1.Location = new Point(margin, 115);
                txtFile1.Bounds = new Rectangle(margin + 50, 112, inputWidth, 25);
                btnBrowse1.Location = new Point(txtFile1.Right + 3, 111);
                lblS1.Location = new Point(margin, 148);
                comboSheet1.Location = new Point(margin + 60, 145);

                lbl2.Location = new Point(midX + margin, 115);
                txtFile2.Bounds = new Rectangle(midX + margin + 50, 112, inputWidth, 25);
                btnBrowse2.Location = new Point(txtFile2.Right + 3, 111);
                lblS2.Location = new Point(midX + margin, 148);
                comboSheet2.Location = new Point(midX + margin + 60, 145);
                splitContainer.SplitterDistance = midX;
            };

            btnBrowse1.Click += (s, e) => LoadExcel(txtFile1, comboSheet1);
            btnBrowse2.Click += (s, e) => LoadExcel(txtFile2, comboSheet2);
            comboSheet1.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile1.Text, comboSheet1.Text, grid1);
            comboSheet2.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile2.Text, comboSheet2.Text, grid2);
            btnCompare.Click += RunComparison;

            btnSyncToggle.Click += (s, e) => {
                isSyncEnabled = !isSyncEnabled;
                btnSyncToggle.Text = isSyncEnabled ? "SYNC: ON" : "SYNC: OFF";
                btnSyncToggle.BackColor = isSyncEnabled ? Color.Honeydew : Color.LightGray;
            };

            btnSwap.Click += (s, e) => {
                int[] widths1 = grid1.Columns.Cast<DataGridViewColumn>().Select(c => c.Width).ToArray();
                int[] widths2 = grid2.Columns.Cast<DataGridViewColumn>().Select(c => c.Width).ToArray();
                string f1 = txtFile1.Text, f2 = txtFile2.Text;
                var items1 = comboSheet1.Items.Cast<object>().ToArray();
                var items2 = comboSheet2.Items.Cast<object>().ToArray();
                int idx1 = comboSheet1.SelectedIndex, idx2 = comboSheet2.SelectedIndex;
                object? ds1 = grid1.DataSource, ds2 = grid2.DataSource;

                isResizingColumn = true;
                grid1.SuspendLayout(); grid2.SuspendLayout();
                try {
                    txtFile1.Text = f2; txtFile2.Text = f1;
                    comboSheet1.Items.Clear(); comboSheet1.Items.AddRange(items2);
                    if (comboSheet1.Items.Count > 0) comboSheet1.SelectedIndex = idx2;
                    comboSheet2.Items.Clear(); comboSheet2.Items.AddRange(items1);
                    if (comboSheet2.Items.Count > 0) comboSheet2.SelectedIndex = idx1;
                    grid1.DataSource = ds2; grid2.DataSource = ds1;
                    for (int i = 0; i < grid1.ColumnCount && i < widths2.Length; i++) grid1.Columns[i].Width = widths2[i];
                    for (int i = 0; i < grid2.ColumnCount && i < widths1.Length; i++) grid2.Columns[i].Width = widths1[i];
                } finally {
                    isResizingColumn = false;
                    grid1.ResumeLayout(); grid2.ResumeLayout();
                }
                ResetGridStyle(grid1); ResetGridStyle(grid2);
            };

            topPanel.Controls.AddRange(new Control[] { btnCompare, btnSwap, btnSyncToggle, lbl1, txtFile1, btnBrowse1, lblS1, comboSheet1, lbl2, txtFile2, btnBrowse2, lblS2, comboSheet2 });
            splitContainer = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BorderStyle = BorderStyle.Fixed3D };
            ConfigGrid(grid1); ConfigGrid(grid2);
            
            grid1.ColumnWidthChanged += (s, e) => SyncColumnWidth(grid1, grid2, e);
            grid2.ColumnWidthChanged += (s, e) => SyncColumnWidth(grid2, grid1, e);
            grid1.Scroll += (s, e) => { if (isSyncEnabled) SyncGrids(grid1, grid2, e); };
            grid2.Scroll += (s, e) => { if (isSyncEnabled) SyncGrids(grid2, grid1, e); };

            splitContainer.Panel1.Controls.Add(grid1);
            splitContainer.Panel2.Controls.Add(grid2);
            this.Controls.Add(splitContainer);
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
                } else { 
                    tar.HorizontalScrollingOffset = src.HorizontalScrollingOffset; 
                } 
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
            } catch { } 
        }

        private void ResetGridStyle(DataGridView g) { 
            if (g.DataSource == null) return; 
            foreach (DataGridViewRow row in g.Rows) 
                foreach (DataGridViewCell cell in row.Cells) cell.Style.BackColor = Color.White; 
        }

        private async void RunComparison(object? sender, EventArgs e) { 
            if (grid1.DataSource == null || grid2.DataSource == null) return; 
            btnCompare.Enabled = false;
            int diffs = 0;
            int rCount = Math.Min(grid1.RowCount, grid2.RowCount); 
            int cCount = Math.Min(grid1.ColumnCount, grid2.ColumnCount); 
            await Task.Run(() => {
                for (int r = 0; r < rCount; r++) { 
                    for (int c = 0; c < cCount; c++) { 
                        string v1 = grid1[c, r].Value?.ToString() ?? "";
                        string v2 = grid2[c, r].Value?.ToString() ?? "";
                        if (v1 != v2) { 
                            diffs++; 
                            grid1[c, r].Style.BackColor = grid2[c, r].Style.BackColor = Color.Yellow; 
                        } else { 
                            grid1[c, r].Style.BackColor = grid2[c, r].Style.BackColor = Color.White; 
                        } 
                    } 
                }
            });
            btnCompare.Enabled = true;
            MessageBox.Show($"Found {diffs} differences.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information); 
        }
    }
}