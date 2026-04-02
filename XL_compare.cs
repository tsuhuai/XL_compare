using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
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
        private static bool IsChinese => CultureInfo.CurrentUICulture.Name.StartsWith("zh");
        private static class Lang {
            public static string AppTitle = IsChinese ? "Excel 智慧比對工具 v6.2" : "Excel Smart Comparator v6.2";
            public static string Compare = IsChinese ? "智慧比對" : "Smart Diff";
            public static string Swap = IsChinese ? "左右對調" : "Swap Side";
            public static string IgnoreBlank = IsChinese ? "忽略空列" : "Ignore Blk";
            public static string Sync = IsChinese ? "同步滾動" : "Sync Scroll";
            public static string File1 = IsChinese ? "檔案 1:" : "File 1:";
            public static string File2 = IsChinese ? "檔案 2:" : "File 2:";
            public static string DropHint = IsChinese ? "【 拖曳 Excel 檔案至此處讀取 】" : "【 Drag & Drop Excel File Here 】";
            public static string MsgSaveSuccess = IsChinese ? "存檔成功！" : "Saved Successfully!";
            public static string MsgError = IsChinese ? "錯誤：" : "Error:";
            public static string Aligning = IsChinese ? "比對中..." : "Aligning...";
            public static string SaveAs = IsChinese ? "另存新檔" : "Save As";
            public static string Ready = IsChinese ? "就緒" : "Ready";
            public static string DiffCount = IsChinese ? "偵測到 {0} 處差異" : "Found {0} differences";
            public static string NoDiff = IsChinese ? "檔案完全一致！" : "No differences found!";
            public static string TipLoad = IsChinese ? "選取並載入 Excel 檔案" : "Browse & Load Excel";
            public static string TipSave = IsChinese ? "儲存內容至原檔案" : "Save Changes to File";
            public static string TipSaveAs = IsChinese ? "將此區域內容另存新檔" : "Save As New Excel File";
        }

        private TextBox txtFile1 = new(), txtFile2 = new();
        private ComboBox comboSheet1 = new(), comboSheet2 = new();
        private DataGridView grid1 = new(), grid2 = new();
        private Button btnCompare = new(), btnSyncToggle = new(), btnSwap = new(), btnIgnoreBlank = new();
        private Button btnBrowse1 = new(), btnBrowse2 = new(), btnSave1 = new(), btnSave2 = new(), btnSaveAs1 = new(), btnSaveAs2 = new();
        private Label lbl1 = new(), lbl2 = new();
        private SplitContainer splitContainer = new();
        private Panel navPanel = new();
        private StatusStrip statusStrip = new();
        private ToolStripStatusLabel lblStatus = new();

        private bool isSyncEnabled = true, isIgnoreBlank = false, isSyncing = false;
        private bool isNavDragging = false; // 用於判斷導覽列是否正在被拖曳
        private List<int> diffRows = new();
        private const string EMPTY_MARKER = "[EMPTY_ROW]";

        public MainForm()
        {
            this.Width = (int)(Screen.PrimaryScreen!.Bounds.Width * 0.85);
            this.Height = (int)(Screen.PrimaryScreen!.Bounds.Height * 0.8);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = Lang.AppTitle;
            this.DoubleBuffered = true;
            try { if (File.Exists("XL_compare.ico")) this.Icon = new Icon("XL_compare.ico"); } catch { }
            SetupDynamicLayout();
        }

        private void SetupDynamicLayout()
        {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 170, BackColor = Color.FromArgb(245, 245, 245), BorderStyle = BorderStyle.FixedSingle };
            int margin = 20;

            ToolTip toolTip = new ToolTip { InitialDelay = 500, ReshowDelay = 100, ShowAlways = true };

            void StyleIconButton(Button btn, string imgName, string text) {
                btn.Size = new Size(100, 85); btn.Text = text; btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                btn.TextAlign = ContentAlignment.BottomCenter; btn.TextImageRelation = TextImageRelation.ImageAboveText;
                btn.BackColor = Color.White; btn.FlatStyle = FlatStyle.Flat; btn.Image = GetResourceImage(imgName, 40, 40);
            }

            void StyleSmallBtn(Button btn, string imgName, string tipText) {
                btn.Size = new Size(36, 26); btn.BackColor = Color.White; btn.FlatStyle = FlatStyle.Flat;
                btn.Image = GetResourceImage(imgName, 20, 20);
                toolTip.SetToolTip(btn, tipText);
            }

            StyleIconButton(btnCompare, "compare.png", Lang.Compare); btnCompare.Location = new Point(margin, 10);
            StyleIconButton(btnSwap, "swap.png", Lang.Swap); btnSwap.Location = new Point(margin + 110, 10);
            StyleIconButton(btnIgnoreBlank, "ignore.png", $"{Lang.IgnoreBlank}: OFF"); btnIgnoreBlank.Location = new Point(margin + 220, 10);
            StyleIconButton(btnSyncToggle, "sync.png", $"{Lang.Sync}: ON"); btnSyncToggle.Location = new Point(margin + 330, 10);
            btnSyncToggle.BackColor = Color.Honeydew;

            lbl1.Text = Lang.File1; lbl1.AutoSize = true; lbl1.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lbl2.Text = Lang.File2; lbl2.AutoSize = true; lbl2.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            StyleSmallBtn(btnBrowse1, "load_file.png", Lang.TipLoad); StyleSmallBtn(btnSave1, "save_file.png", Lang.TipSave); StyleSmallBtn(btnSaveAs1, "save_as.png", Lang.TipSaveAs);
            StyleSmallBtn(btnBrowse2, "load_file.png", Lang.TipLoad); StyleSmallBtn(btnSave2, "save_file.png", Lang.TipSave); StyleSmallBtn(btnSaveAs2, "save_as.png", Lang.TipSaveAs);

            lblStatus.Text = Lang.Ready;
            statusStrip.Items.Add(lblStatus);
            statusStrip.BackColor = Color.White;

            // --- 導覽列：增加主動拖曳與同步功能 ---
            navPanel.Dock = DockStyle.Right; navPanel.Width = 40; navPanel.BackColor = Color.FromArgb(45, 45, 48);
            navPanel.Cursor = Cursors.Hand;
            
            navPanel.Paint += (s, e) => {
                if (grid1.RowCount <= 0) return;
                float ratio = (float)navPanel.Height / grid1.RowCount;
                using var p = new Pen(Color.Orange, 2);
                foreach (int row in diffRows) e.Graphics.DrawLine(p, 5, row * ratio, navPanel.Width - 5, row * ratio);

                // 繪製目前的 Focus 區域 (滑塊)
                int visibleRows = grid1.DisplayedRowCount(false);
                int firstIdx = grid1.FirstDisplayedScrollingRowIndex;
                if (firstIdx >= 0) {
                    float rectY = firstIdx * ratio;
                    float rectH = Math.Max(10, visibleRows * ratio);
                    using var brush = new SolidBrush(Color.FromArgb(100, Color.White)); 
                    e.Graphics.FillRectangle(brush, 4, rectY, navPanel.Width - 8, rectH);
                    e.Graphics.DrawRectangle(Pens.White, 4, rectY, navPanel.Width - 8, rectH);
                }
            };

            // 處理導覽列的滑鼠連動
            Action<int> handleNavScroll = (mouseY) => {
                if (grid1.RowCount <= 0) return;
                float ratio = (float)navPanel.Height / grid1.RowCount;
                int targetRow = (int)(mouseY / ratio);
                
                // 限制邊界避免報錯
                targetRow = Math.Max(0, Math.Min(targetRow, grid1.RowCount - 1));
                
                if (grid1.FirstDisplayedScrollingRowIndex != targetRow) {
                    grid1.FirstDisplayedScrollingRowIndex = targetRow;
                    // 如果開啟同步滾動，則強制另一邊也要跳轉
                    if (isSyncEnabled && grid2.RowCount > targetRow)
                        grid2.FirstDisplayedScrollingRowIndex = targetRow;
                    
                    navPanel.Invalidate(); // 更新導覽列視覺
                }
            };

            navPanel.MouseDown += (s, e) => { isNavDragging = true; handleNavScroll(e.Y); };
            navPanel.MouseMove += (s, e) => { if (isNavDragging) handleNavScroll(e.Y); };
            navPanel.MouseUp += (s, e) => { isNavDragging = false; };

            splitContainer = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BorderStyle = BorderStyle.Fixed3D };
            ConfigGrid(grid1); ConfigGrid(grid2);
            splitContainer.Panel1.Controls.Add(grid1); splitContainer.Panel2.Controls.Add(grid2);

            topPanel.Controls.AddRange(new Control[] { btnCompare, btnSwap, btnIgnoreBlank, btnSyncToggle, lbl1, txtFile1, btnBrowse1, btnSave1, btnSaveAs1, comboSheet1, lbl2, txtFile2, btnBrowse2, btnSave2, btnSaveAs2, comboSheet2 });
            this.Controls.Add(splitContainer); this.Controls.Add(navPanel); this.Controls.Add(statusStrip); this.Controls.Add(topPanel);

            // 事件繫結
            btnCompare.Click += RunSmartComparison;
            btnSwap.Click += SwapGrids;
            btnIgnoreBlank.Click += (s, e) => { isIgnoreBlank = !isIgnoreBlank; btnIgnoreBlank.Text = $"{Lang.IgnoreBlank}: {(isIgnoreBlank ? "ON" : "OFF")}"; btnIgnoreBlank.BackColor = isIgnoreBlank ? Color.NavajoWhite : Color.White; };
            btnSyncToggle.Click += (s, e) => { isSyncEnabled = !isSyncEnabled; btnSyncToggle.Text = $"{Lang.Sync}: {(isSyncEnabled ? "ON" : "OFF")}"; btnSyncToggle.BackColor = isSyncEnabled ? Color.Honeydew : Color.LightGray; };
            btnBrowse1.Click += (s, e) => LoadExcel(txtFile1, comboSheet1);
            btnBrowse2.Click += (s, e) => LoadExcel(txtFile2, comboSheet2);
            btnSave1.Click += (s, e) => SaveExcel(txtFile1.Text, comboSheet1.Text, grid1);
            btnSave2.Click += (s, e) => SaveExcel(txtFile2.Text, comboSheet2.Text, grid2);
            btnSaveAs1.Click += (s, e) => SaveExcelAs(txtFile1, comboSheet1, grid1);
            btnSaveAs2.Click += (s, e) => SaveExcelAs(txtFile2, comboSheet2, grid2);
            comboSheet1.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile1.Text, comboSheet1.Text, grid1);
            comboSheet2.SelectedIndexChanged += (s, e) => DisplaySheet(txtFile2.Text, comboSheet2.Text, grid2);

            EnableDragDrop(grid1, txtFile1, comboSheet1);
            EnableDragDrop(grid2, txtFile2, comboSheet2);

            Action doLayout = () => {
                if (this.WindowState == FormWindowState.Minimized) return;
                int availW = this.ClientSize.Width - navPanel.Width;
                int midX = availW / 2;
                lbl1.Location = new Point(margin, 105); txtFile1.Bounds = new Rectangle(margin + 50, 105, Math.Max(50, midX - 250), 25);
                btnBrowse1.Location = new Point(txtFile1.Right + 3, 104); btnSave1.Location = new Point(btnBrowse1.Right + 3, 104); btnSaveAs1.Location = new Point(btnSave1.Right + 3, 104);
                comboSheet1.Bounds = new Rectangle(margin + 50, 135, 120, 25);
                lbl2.Location = new Point(midX + margin, 105); txtFile2.Bounds = new Rectangle(midX + margin + 50, 105, Math.Max(50, midX - 250), 25);
                btnBrowse2.Location = new Point(txtFile2.Right + 3, 104); btnSave2.Location = new Point(btnBrowse2.Right + 3, 104); btnSaveAs2.Location = new Point(btnSave2.Right + 3, 104);
                comboSheet2.Bounds = new Rectangle(midX + margin + 50, 135, 120, 25);
            };

            this.Resize += (s, e) => doLayout();
            this.Load += (s, e) => {
                doLayout();
                int initialMid = (this.ClientSize.Width - navPanel.Width) / 2;
                if (initialMid > 100) splitContainer.SplitterDistance = initialMid;
            };
        }

        private void ConfigGrid(DataGridView g)
        {
            g.Dock = DockStyle.Fill; 
            g.BackgroundColor = Color.White; 
        
            // 顯示行號區域
            g.RowHeadersVisible = true; 
            g.RowHeadersWidth = 60; 
        
            g.AllowUserToAddRows = false; 
            g.EnableHeadersVisualStyles = false;
            g.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy; 
            g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        
            // --- 【關鍵設定 1】允許使用者手動調整「標題列」高度 ---
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
        
            // --- 【關鍵設定 2】禁止使用者手動調整「資料列」高度 ---
            g.AllowUserToResizeRows = false; 
        
            // 1. 連動：欄位寬度 (在 Header 調整時，左右同步，資料區會自動跟隨)
            g.ColumnWidthChanged += (s, e) => {
                if (isSyncing) return;
                isSyncing = true;
                DataGridView source = (DataGridView)s;
                DataGridView target = (source == grid1) ? grid2 : grid1;
                try {
                    int colIdx = e.Column.Index;
                    if (target.ColumnCount > colIdx) {
                        target.Columns[colIdx].Width = source.Columns[colIdx].Width;
                    }
                } finally { isSyncing = false; }
            };
        
            // 2. 連動：Header Row 的高度 (當你拉動標題列下緣時觸發)
            g.ColumnHeadersHeightChanged += (s, e) => {
                if (isSyncing) return;
                isSyncing = true;
                DataGridView source = (DataGridView)s;
                DataGridView target = (source == grid1) ? grid2 : grid1;
                try {
                    target.ColumnHeadersHeight = source.ColumnHeadersHeight;
                } finally { isSyncing = false; }
            };
        
            // --- 繪製行號 (1, 2, 3...) ---
            g.RowPostPaint += (s, e) => {
                var grid = (DataGridView)s;
                string rowIdx = (e.RowIndex + 1).ToString();
                var centerFormat = new StringFormat() {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                Rectangle headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
                e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
            };
        
            // --- 繪製拖曳提示文字 ---
            g.Paint += (s, e) => {
                if (g.Rows.Count == 0) {
                    using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    e.Graphics.DrawString(Lang.DropHint, new Font(IsChinese ? "Microsoft JhengHei" : "Segoe UI", 11, FontStyle.Bold), Brushes.LightGray, new RectangleF(0, 0, g.Width, g.Height), sf);
                }
            };
        
            // --- 處理 [EMPTY_ROW] 的斜線背景 ---
            g.CellPainting += (s, e) => {
                if (e.RowIndex >= 0 && e.Value?.ToString() == EMPTY_MARKER) {
                    e.PaintBackground(e.CellBounds, true);
                    using var hb = new HatchBrush(HatchStyle.BackwardDiagonal, Color.FromArgb(200, 200, 200), Color.Transparent);
                    e.Graphics.FillRectangle(hb, e.CellBounds); e.Handled = true;
                }
            };
        
            // --- 同步捲動 (含導覽列重繪) ---
            g.Scroll += (s, e) => { 
                if (isSyncEnabled && !isSyncing) SyncGrids(g, (g == grid1 ? grid2 : grid1), e); 
                navPanel.Invalidate(); 
            };
        
            // 啟動雙緩衝，防止 14,000 列捲動時閃爍
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(g, true);
        }

        //private void ConfigGrid(DataGridView g)
        //{
        //    g.Dock = DockStyle.Fill; 
		//	g.BackgroundColor = Color.White; 
		//
        //    g.RowHeadersVisible = true; 
        //    g.RowHeadersWidth = 60; // 14,000 列建議設定 60 左右才不會被遮住
		//
        //    g.AllowUserToAddRows = false; 
		//	g.EnableHeadersVisualStyles = false;
        //    g.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy; 
		//	g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
		//
        //    g.RowPostPaint += (s, e) => {
        //        var grid = (DataGridView)s;
        //        string rowIdx = (e.RowIndex + 1).ToString();
        //        var centerFormat = new StringFormat() {
        //            Alignment = StringAlignment.Center,
        //            LineAlignment = StringAlignment.Center
        //        };
        //        Rectangle headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
        //        e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        //    };
		//
        //    g.Paint += (s, e) => {
        //        if (g.Rows.Count == 0) {
        //            using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        //            e.Graphics.DrawString(Lang.DropHint, new Font(IsChinese ? "Microsoft JhengHei" : "Segoe UI", 11, FontStyle.Bold), Brushes.LightGray, new RectangleF(0, 0, g.Width, g.Height), sf);
        //        }
        //    };
        //    g.CellPainting += (s, e) => {
        //        if (e.RowIndex >= 0 && e.Value?.ToString() == EMPTY_MARKER) {
        //            e.PaintBackground(e.CellBounds, true);
        //            using var hb = new HatchBrush(HatchStyle.BackwardDiagonal, Color.FromArgb(200, 200, 200), Color.Transparent);
        //            e.Graphics.FillRectangle(hb, e.CellBounds); e.Handled = true;
        //        }
        //    };
        //    g.Scroll += (s, e) => { 
        //        if (isSyncEnabled && !isSyncing) SyncGrids(g, (g == grid1 ? grid2 : grid1), e); 
        //        navPanel.Invalidate(); 
        //    };
        //    typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(g, true);
        //}

        private async void RunSmartComparison(object? sender, EventArgs e)
        {
            if (grid1.DataSource == null || grid2.DataSource == null) return;
            btnCompare.Enabled = false; btnCompare.Text = Lang.Aligning; lblStatus.Text = Lang.Aligning;
            DataTable dt1 = (DataTable)grid1.DataSource; DataTable dt2 = (DataTable)grid2.DataSource;
            var result = await Task.Run(() => {
                var list1 = dt1.AsEnumerable().Select(r => string.Join("|", r.ItemArray)).ToList();
                var list2 = dt2.AsEnumerable().Select(r => string.Join("|", r.ItemArray)).ToList();
                DataTable v1 = dt1.Clone(), v2 = dt2.Clone();
                int i = 0, j = 0, cur = 0; HashSet<int> diffs = new();
                while (i < list1.Count || j < list2.Count) {
                    if (i < list1.Count && j < list2.Count && list1[i] == list2[j]) {
                        v1.ImportRow(dt1.Rows[i]); v2.ImportRow(dt2.Rows[j]); i++; j++;
                    } else {
                        bool found = false;
                        for (int k = 1; k <= 15; k++) {
                            if (i + k < list1.Count && j < list2.Count && list1[i + k] == list2[j]) {
                                v1.ImportRow(dt1.Rows[i]); AddEmpty(v2); diffs.Add(cur); i++; found = true; break;
                            }
                            if (j + k < list2.Count && i < list1.Count && list2[j + k] == list1[i]) {
                                v2.ImportRow(dt2.Rows[j]); AddEmpty(v1); diffs.Add(cur); j++; found = true; break;
                            }
                        }
                        if (!found && i < list1.Count && j < list2.Count) {
                            v1.ImportRow(dt1.Rows[i]); v2.ImportRow(dt2.Rows[j]); diffs.Add(cur); i++; j++;
                        } else if (!found) {
                            if (i < list1.Count) { v1.ImportRow(dt1.Rows[i]); AddEmpty(v2); i++; }
                            else { v2.ImportRow(dt2.Rows[j]); AddEmpty(v1); j++; }
                            diffs.Add(cur);
                        }
                    }
                    cur++;
                }
                return new { V1 = v1, V2 = v2, Rows = diffs.OrderBy(x => x).ToList() };
            });

            grid1.DataSource = result.V1; grid2.DataSource = result.V2; diffRows = result.Rows;

            // --- 完整差異著色邏輯 (v6.3) ---
            foreach (int r in diffRows)
            {
                bool isR1Valid = r < grid1.RowCount;
                bool isR2Valid = r < grid2.RowCount;
            
                // --- (1) 缺少的 Row (其中一邊是 [EMPTY_ROW] 佔位符) ---
                if ((isR1Valid && grid1[0, r].Value?.ToString() == EMPTY_MARKER) || 
                    (isR2Valid && grid2[0, r].Value?.ToString() == EMPTY_MARKER))
                {
                    // 建議 R/G/B: (240, 240, 240) - 極淺灰色，代表此處無資料
                    int missingR = 240, missingG = 240, missingB = 240;
                    Color missingColor = Color.FromArgb(missingR, missingG, missingB);
            
                    if (isR1Valid) {
                        for (int c = 0; c < grid1.ColumnCount; c++) grid1[c, r].Style.BackColor = missingColor;
                    }
                    if (isR2Valid) {
                        for (int c = 0; c < grid2.ColumnCount; c++) grid2[c, r].Style.BackColor = missingColor;
                    }
                }
                // --- (2) 不一樣的 Row (兩邊都有資料，但內容不同) ---
                else
                {
                    // 建議 R/G/B: (255, 240, 240) - 極淺粉色，作為整列差異的底色
                    int rowDiffR = 255, rowDiffG = 240, rowDiffB = 240;
                    Color rowDiffColor = Color.FromArgb(rowDiffR, rowDiffG, rowDiffB);
            
                    if (isR1Valid) {
                        for (int c = 0; c < grid1.ColumnCount; c++) grid1[c, r].Style.BackColor = rowDiffColor;
                    }
                    if (isR2Valid) {
                        for (int c = 0; c < grid2.ColumnCount; c++) grid2[c, r].Style.BackColor = rowDiffColor;
                    }
            
                    // --- (3) 不一樣的 Row 裡面，「具體差異」的 Cell ---
                    if (isR1Valid && isR2Valid)
                    {
                        for (int c = 0; c < grid1.ColumnCount; c++)
                        {
                            string val1 = grid1[c, r].Value?.ToString() ?? "";
                            string val2 = grid2[c, r].Value?.ToString() ?? "";
            
                            if (val1 != val2)
                            {
                                // 左邊 Cell 差異：建議淡薄荷綠 (180, 230, 180)
                                int leftCellR = 180, leftCellG = 230, leftCellB = 180;
                                grid1[c, r].Style.BackColor = Color.FromArgb(leftCellR, leftCellG, leftCellB);
                                grid1[c, r].Style.ForeColor = Color.Black; // 確保文字清楚
                                grid1[c, r].Style.Font = new Font(grid1.Font, FontStyle.Bold);
            
                                // 右邊 Cell 差異：建議淡天空藍 (180, 210, 240)
                                int rightCellR = 180, rightCellG = 210, rightCellB = 240;
                                grid2[c, r].Style.BackColor = Color.FromArgb(rightCellR, rightCellG, rightCellB);
                                grid2[c, r].Style.ForeColor = Color.Black; // 確保文字清楚
                                grid2[c, r].Style.Font = new Font(grid2.Font, FontStyle.Bold);
                            }
                        }
                    }
                }
            }
			
            lblStatus.Text = diffRows.Count > 0 ? string.Format(Lang.DiffCount, diffRows.Count) : Lang.NoDiff;
            lblStatus.ForeColor = diffRows.Count > 0 ? Color.Crimson : Color.DarkGreen;
            btnCompare.Enabled = true; btnCompare.Text = Lang.Compare; navPanel.Invalidate();
        }

        private void AddEmpty(DataTable dt) {
            DataRow dr = dt.NewRow(); for (int i = 0; i < dt.Columns.Count; i++) dr[i] = EMPTY_MARKER;
            dt.Rows.Add(dr);
        }

        private void SaveExcel(string p, string s, DataGridView g) {
            try {
                using var wb = new XLWorkbook(p); var ws = wb.Worksheet(s);
                ws.Clear(XLClearOptions.Contents); DataTable dt = (DataTable)g.DataSource;
                for (int i = 0; i < dt.Columns.Count; i++) ws.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                int off = 2;
                for (int r = 0; r < dt.Rows.Count; r++) {
                    if (dt.Rows[r][0].ToString() == EMPTY_MARKER) continue;
                    for (int c = 0; c < dt.Columns.Count; c++) ws.Cell(off, c + 1).Value = dt.Rows[r][c]?.ToString();
                    off++;
                }
                wb.Save(); MessageBox.Show(Lang.MsgSaveSuccess);
            } catch (Exception ex) { MessageBox.Show($"{Lang.MsgError} {ex.Message}"); }
        }

        private void SaveExcelAs(TextBox t, ComboBox c, DataGridView g) {
            using SaveFileDialog sfd = new() { Filter = "Excel|*.xlsx", Title = Lang.SaveAs };
            if (sfd.ShowDialog() == DialogResult.OK) {
                try {
                    using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add(c.Text);
                    DataTable dt = (DataTable)g.DataSource;
                    for (int i = 0; i < dt.Columns.Count; i++) ws.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                    int off = 2;
                    for (int r = 0; r < dt.Rows.Count; r++) {
                        if (dt.Rows[r][0].ToString() == EMPTY_MARKER) continue;
                        for (int col = 0; col < dt.Columns.Count; col++) ws.Cell(off, col + 1).Value = dt.Rows[r][col]?.ToString();
                        off++;
                    }
                    wb.SaveAs(sfd.FileName); t.Text = sfd.FileName; MessageBox.Show(Lang.MsgSaveSuccess);
                } catch (Exception ex) { MessageBox.Show($"{Lang.MsgError} {ex.Message}"); }
            }
        }

        private void SyncGrids(DataGridView s, DataGridView t, ScrollEventArgs e) {
            isSyncing = true;
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) {
                if (s.FirstDisplayedScrollingRowIndex >= 0 && s.FirstDisplayedScrollingRowIndex < t.RowCount)
                    t.FirstDisplayedScrollingRowIndex = s.FirstDisplayedScrollingRowIndex;
            } else t.HorizontalScrollingOffset = s.HorizontalScrollingOffset;
            isSyncing = false;
        }

        private void EnableDragDrop(Control ctrl, TextBox t, ComboBox c) {
            ctrl.AllowDrop = true;
            ctrl.DragEnter += (s, e) => { if (e.Data!.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            ctrl.DragDrop += (s, e) => {
                string[] files = (string[])e.Data!.GetData(DataFormats.FileDrop);
                if (files.Length > 0) { t.Text = files[0]; LoadSheetsAfterDrop(t.Text, c); }
            };
        }

        private void LoadExcel(TextBox t, ComboBox c) {
            OpenFileDialog ofd = new() { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() == DialogResult.OK) { t.Text = ofd.FileName; LoadSheetsAfterDrop(t.Text, c); }
        }

        private void LoadSheetsAfterDrop(string p, ComboBox c) {
            try { c.Items.Clear(); using var wb = new XLWorkbook(p); foreach (var s in wb.Worksheets) c.Items.Add(s.Name); if (c.Items.Count > 0) c.SelectedIndex = 0; } catch { }
        }

        private void DisplaySheet(string p, string s, DataGridView g) {
            if (string.IsNullOrEmpty(p) || string.IsNullOrEmpty(s)) return;
            try {
                DataTable dt = new(); using var wb = new XLWorkbook(p); var ws = wb.Worksheet(s);
                var range = ws.RangeUsed(); if (range == null) return;
                for (int i = 1; i <= range.LastColumn().ColumnNumber(); i++) dt.Columns.Add(ws.Cell(1, i).Value.ToString());
                foreach (var r in ws.RowsUsed().Skip(1)) {
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count; i++) dr[i] = r.Cell(i + 1).Value.ToString();
                    dt.Rows.Add(dr);
                }
                g.DataSource = dt; g.AutoResizeColumns();
            } catch { }
        }

        private Image? GetResourceImage(string n, int w, int h) {
            try {
                var asm = Assembly.GetExecutingAssembly();
                var res = asm.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith(n));
                if (res != null) {
                    using var s = asm.GetManifestResourceStream(res);
                    using Bitmap b = new(s!); Bitmap r = new(w, h);
                    using Graphics g = Graphics.FromImage(r);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(b, 0, 0, w, h); return r;
                }
            } catch { } return null;
        }

        private void SwapGrids(object? sender, EventArgs e) {
            var temp = grid1.DataSource; grid1.DataSource = grid2.DataSource; grid2.DataSource = temp;
            var t = txtFile1.Text; txtFile1.Text = txtFile2.Text; txtFile2.Text = t;
        }
    }
}