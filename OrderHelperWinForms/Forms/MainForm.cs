using System.Diagnostics;
using OrderHelperWinForms.Models;
using OrderHelperWinForms.Services;

namespace OrderHelperWinForms.Forms;

public class MainForm : Form
{
    // ---- Tab 1 controls ----
    readonly Button         _btnSelectExcel  = new();
    readonly Button         _btnExportSample = new();
    readonly Label          _lblExcelPath    = new();
    readonly DateTimePicker _dtpOrderDate    = new();
    readonly CheckBox       _chkAutoDate     = new();
    readonly Button         _btnGenerate     = new();
    readonly ProgressBar    _progress        = new();
    readonly Label          _lblStatus       = new();
    readonly Label          _lblValidation   = new();
    readonly DataGridView   _dgvErrors       = new();

    // ---- Tab 2 controls ----
    readonly DataGridView   _dgvRules        = new();
    readonly Button         _btnSaveRules    = new();
    readonly Button         _btnResetRules   = new();

    // ---- Tab 3 controls ----
    readonly Dictionary<string, TextBox> _hsFields = new();
    readonly Button         _btnSaveHospital  = new();
    readonly Button         _btnResetHospital = new();

    // ---- State ----
    string?           _excelPath;
    string?           _selectedSheet;
    ValidationConfig  _validationConfig = AppSettings.LoadValidation();
    HospitalSettings  _hospitalSettings = AppSettings.LoadHospital();

    public MainForm()
    {
        Text            = "義大醫院 藥品訂購單 PDF 產生器";
        Size            = new Size(740, 660);
        MinimumSize     = new Size(660, 580);
        StartPosition   = FormStartPosition.CenterScreen;

        var menu = BuildMenu();
        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildTab1());
        tabs.TabPages.Add(BuildTab2());
        tabs.TabPages.Add(BuildTab3());

        Controls.Add(menu);   // Dock=Top by default for MenuStrip
        Controls.Add(tabs);   // Fill
        MainMenuStrip = menu;

        BindRulesGrid();
        BindHospitalFields();
    }

    // ============================================================
    // Menu
    // ============================================================
    MenuStrip BuildMenu()
    {
        var menu     = new MenuStrip();
        var miTools  = new ToolStripMenuItem("工具(&T)");
        var miLog    = new ToolStripMenuItem("檢視操作記錄…");
        miLog.Click += (_, _) => new LogViewerForm().ShowDialog(this);
        miTools.DropDownItems.Add(miLog);
        menu.Items.Add(miTools);
        return menu;
    }

    // ============================================================
    // Tab 1 — 訂購單產生
    // ============================================================
    TabPage BuildTab1()
    {
        var page = new TabPage("訂購單產生") { Padding = new Padding(10) };
        const int P = 12;
        var y = P;

        // Row 1: Excel file selection
        var lblTitle = new Label
        {
            Text = "Excel 訂購檔：", Left = P, Top = y, Width = 100, Height = 26,
            TextAlign = ContentAlignment.MiddleLeft,
        };

        _btnSelectExcel.Text = "選擇檔案…";
        _btnSelectExcel.SetBounds(P + 100, y, 100, 26);
        _btnSelectExcel.Click += BtnSelectExcel_Click;

        _lblExcelPath.SetBounds(P + 208, y, 0, 26); // width set in resize
        _lblExcelPath.Text      = "（尚未選擇）";
        _lblExcelPath.ForeColor = Color.Gray;
        _lblExcelPath.TextAlign = ContentAlignment.MiddleLeft;
        _lblExcelPath.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 34;

        // Row 2: Export sample button
        _btnExportSample.Text      = "匯出範例 Excel…";
        _btnExportSample.SetBounds(P + 100, y, 130, 24);
        _btnExportSample.ForeColor = Color.DarkGreen;
        _btnExportSample.FlatStyle = FlatStyle.Flat;
        _btnExportSample.Click    += BtnExportSample_Click;
        var lblSampleHint = new Label
        {
            Text      = "下載填寫範本",
            Left      = P + 238, Top = y + 3,
            Width     = 200, Height = 18,
            ForeColor = Color.DimGray,
        };
        y += 32;

        // Row 3: Date picker
        var lblDate = new Label
        {
            Text = "訂貨日期：", Left = P, Top = y, Width = 100, Height = 26,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        _dtpOrderDate.SetBounds(P + 100, y, 150, 26);
        _dtpOrderDate.Format = DateTimePickerFormat.Short;
        _dtpOrderDate.Value  = DateTime.Today;

        _chkAutoDate.Text    = "自動從檔名/單號推算";
        _chkAutoDate.SetBounds(P + 258, y + 3, 180, 22);
        _chkAutoDate.Checked = true;
        _chkAutoDate.CheckedChanged += (_, _) => _dtpOrderDate.Enabled = !_chkAutoDate.Checked;
        _dtpOrderDate.Enabled = false;
        y += 38;

        // Separator
        var sep = new Panel
        {
            Left = P, Top = y, Height = 1,
            BackColor = Color.LightGray,
            Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        y += 10;

        // Generate button
        _btnGenerate.Text      = "產生 PDF";
        _btnGenerate.SetBounds(P, y, 120, 36);
        _btnGenerate.Font      = new Font(_btnGenerate.Font, FontStyle.Bold);
        _btnGenerate.BackColor = Color.FromArgb(32, 84, 147);
        _btnGenerate.ForeColor = Color.White;
        _btnGenerate.FlatStyle = FlatStyle.Flat;
        _btnGenerate.Enabled   = false;
        _btnGenerate.Click    += BtnGenerate_Click;
        y += 48;

        // Progress bar
        _progress.SetBounds(P, y, 0, 6);
        _progress.Visible = false;
        _progress.Anchor  = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 14;

        // Status label
        _lblStatus.SetBounds(P, y, 0, 40);
        _lblStatus.Text      = "請先選擇 Excel 檔案。";
        _lblStatus.ForeColor = Color.DimGray;
        _lblStatus.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 44;

        // Validation header
        _lblValidation.SetBounds(P, y, 0, 20);
        _lblValidation.Visible   = false;
        _lblValidation.ForeColor = Color.Crimson;
        _lblValidation.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 22;

        // Validation error grid
        _dgvErrors.SetBounds(P, y, 0, 160);
        _dgvErrors.Visible              = false;
        _dgvErrors.ReadOnly             = true;
        _dgvErrors.AllowUserToAddRows   = false;
        _dgvErrors.RowHeadersVisible    = false;
        _dgvErrors.SelectionMode        = DataGridViewSelectionMode.FullRowSelect;
        _dgvErrors.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvErrors.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Row",     HeaderText = "列",     Width = 42 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Field",   HeaderText = "欄位",   Width = 110 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Message", HeaderText = "錯誤訊息",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

        page.Controls.AddRange(new Control[]
        {
            lblTitle, _btnSelectExcel, _lblExcelPath,
            _btnExportSample, lblSampleHint,
            lblDate, _dtpOrderDate, _chkAutoDate,
            sep, _btnGenerate, _progress, _lblStatus,
            _lblValidation, _dgvErrors,
        });

        // Resize handler for anchored-width controls
        page.Resize += (_, _) =>
        {
            int w = page.ClientSize.Width - P * 2;
            _lblExcelPath.Width    = page.ClientSize.Width - P - 208 - P;
            sep.Width              = w;
            _progress.Width        = w;
            _lblStatus.Width       = w;
            _lblValidation.Width   = w;
            _dgvErrors.Width       = w;
        };

        return page;
    }

    // ============================================================
    // Tab 2 — 檢核設定
    // ============================================================
    TabPage BuildTab2()
    {
        var page = new TabPage("檢核設定") { Padding = new Padding(10) };
        const int P = 12;

        var lbl = new Label
        {
            Text   = "設定訂單資料的檢核規則（可新增/刪除列、勾選啟用）：",
            Left   = P, Top = P, Width = 650, Height = 22,
        };

        _dgvRules.SetBounds(P, P + 26, 0, 0);
        _dgvRules.AllowUserToAddRows    = true;
        _dgvRules.AllowUserToDeleteRows = true;
        _dgvRules.EditMode              = DataGridViewEditMode.EditOnEnter;
        _dgvRules.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvRules.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;

        var colEnabled = new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "啟用", Width = 50 };
        _dgvRules.Columns.Add(colEnabled);

        var colField = new DataGridViewComboBoxColumn { Name = "Field", HeaderText = "欄位", Width = 110, FlatStyle = FlatStyle.Flat };
        foreach (var f in ValidationService.KnownFields) colField.Items.Add(f);
        _dgvRules.Columns.Add(colField);

        var colType = new DataGridViewComboBoxColumn { Name = "RuleType", HeaderText = "規則類型", Width = 105, FlatStyle = FlatStyle.Flat };
        colType.Items.AddRange("Required", "Regex", "MaxLength");
        _dgvRules.Columns.Add(colType);

        _dgvRules.Columns.Add(new DataGridViewTextBoxColumn { Name = "Parameter", HeaderText = "參數（正則/長度）", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        _dgvRules.Columns.Add(new DataGridViewTextBoxColumn { Name = "Message",   HeaderText = "錯誤訊息", Width = 180 });

        _btnSaveRules.Text  = "儲存設定";
        _btnResetRules.Text = "還原預設值";
        _btnSaveRules.Size  = _btnResetRules.Size = new Size(110, 30);
        _btnSaveRules.Click  += BtnSaveRules_Click;
        _btnResetRules.Click += BtnResetRules_Click;

        page.Controls.AddRange(new Control[] { lbl, _dgvRules, _btnSaveRules, _btnResetRules });

        page.Resize += (_, _) =>
        {
            int w = page.ClientSize.Width - P * 2;
            int h = page.ClientSize.Height - P * 2 - 26 - 44;
            _dgvRules.Width  = w;
            _dgvRules.Height = Math.Max(100, h);
            int btnY = _dgvRules.Bottom + 8;
            _btnSaveRules.SetBounds(P, btnY, 110, 30);
            _btnResetRules.SetBounds(P + 118, btnY, 120, 30);
        };

        return page;
    }

    // ============================================================
    // Tab 3 — PDF 樣式設定
    // ============================================================
    TabPage BuildTab3()
    {
        var page = new TabPage("PDF 樣式設定") { Padding = new Padding(10) };
        const int P = 12;

        var scrollPanel = new Panel { Left = P, Top = P, AutoScroll = true };
        scrollPanel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;

        (string key, string label)[] fields =
        {
            ("HospitalName",    "醫院名稱（標題大字）"),
            ("FormTitle",       "表單標題"),
            ("InvoiceHeader",   "發票抬頭"),
            ("InvoiceAddress",  "發票地址"),
            ("TaxId",           "統一編號"),
            ("MedicalCode",     "醫療機構代碼"),
            ("DrugLicenseNo",   "管證字號"),
            ("DeliveryAddress", "交貨地址"),
            ("DeliveryNote",    "交貨備注"),
            ("ContactPhone",    "聯絡電話"),
            ("ContactFax",      "傳真"),
            ("Note1",           "備註1"),
            ("Note2",           "備註2"),
            ("Note3",           "備註3"),
            ("Note4",           "備註4"),
        };

        int y = 0;
        const int LblW = 160;
        foreach (var (key, label) in fields)
        {
            var lbl = new Label
            {
                Text = label + "：", Left = 0, Top = y, Width = LblW, Height = 26,
                TextAlign = ContentAlignment.MiddleRight,
            };
            var tb = new TextBox
            {
                Left = LblW + 4, Top = y, Height = 26,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
            };
            _hsFields[key] = tb;
            scrollPanel.Controls.Add(lbl);
            scrollPanel.Controls.Add(tb);
            y += 32;
        }

        _btnSaveHospital.Text   = "儲存設定";
        _btnResetHospital.Text  = "還原預設值";
        _btnSaveHospital.SetBounds(LblW + 4, y + 8, 110, 30);
        _btnResetHospital.SetBounds(LblW + 122, y + 8, 120, 30);
        _btnSaveHospital.Click  += BtnSaveHospital_Click;
        _btnResetHospital.Click += BtnResetHospital_Click;
        scrollPanel.Controls.Add(_btnSaveHospital);
        scrollPanel.Controls.Add(_btnResetHospital);

        page.Controls.Add(scrollPanel);

        page.Resize += (_, _) =>
        {
            int w = page.ClientSize.Width - P * 2;
            int h = page.ClientSize.Height - P * 2;
            scrollPanel.Width  = w;
            scrollPanel.Height = h;
            int tbW = w - LblW - 4 - 20; // -20 for scrollbar
            foreach (var tb in _hsFields.Values) tb.Width = Math.Max(60, tbW);
        };

        return page;
    }

    // ============================================================
    // Data binding
    // ============================================================
    void BindRulesGrid()
    {
        _dgvRules.Rows.Clear();
        foreach (var r in _validationConfig.Rules)
            _dgvRules.Rows.Add(r.Enabled, r.Field, r.RuleType.ToString(), r.Parameter, r.Message);
    }

    void BindHospitalFields()
    {
        var hs = _hospitalSettings;
        Set("HospitalName",    hs.HospitalName);    Set("FormTitle",       hs.FormTitle);
        Set("InvoiceHeader",   hs.InvoiceHeader);   Set("InvoiceAddress",  hs.InvoiceAddress);
        Set("TaxId",           hs.TaxId);           Set("MedicalCode",     hs.MedicalCode);
        Set("DrugLicenseNo",   hs.DrugLicenseNo);   Set("DeliveryAddress", hs.DeliveryAddress);
        Set("DeliveryNote",    hs.DeliveryNote);    Set("ContactPhone",    hs.ContactPhone);
        Set("ContactFax",      hs.ContactFax);
        Set("Note1", hs.Note1); Set("Note2", hs.Note2); Set("Note3", hs.Note3); Set("Note4", hs.Note4);
    }

    void Set(string key, string value) { if (_hsFields.TryGetValue(key, out var tb)) tb.Text = value; }
    string Get(string key) => _hsFields.TryGetValue(key, out var tb) ? tb.Text : "";

    HospitalSettings ReadHospitalFromUI() => new()
    {
        HospitalName    = Get("HospitalName"),    FormTitle       = Get("FormTitle"),
        InvoiceHeader   = Get("InvoiceHeader"),   InvoiceAddress  = Get("InvoiceAddress"),
        TaxId           = Get("TaxId"),           MedicalCode     = Get("MedicalCode"),
        DrugLicenseNo   = Get("DrugLicenseNo"),   DeliveryAddress = Get("DeliveryAddress"),
        DeliveryNote    = Get("DeliveryNote"),     ContactPhone    = Get("ContactPhone"),
        ContactFax      = Get("ContactFax"),
        Note1 = Get("Note1"), Note2 = Get("Note2"), Note3 = Get("Note3"), Note4 = Get("Note4"),
    };

    // ============================================================
    // Tab 1 event handlers
    // ============================================================
    void BtnSelectExcel_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title  = "選擇訂購 Excel 檔",
            Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        string path = dlg.FileName;

        // Sheet detection
        List<string> sheets;
        try { sheets = ExcelReader.ListSheets(path); }
        catch (Exception ex)
        {
            MessageBox.Show("無法讀取 Excel 檔案：" + ex.Message, "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (sheets.Count == 0)
        {
            MessageBox.Show("Excel 沒有工作表。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (sheets.Count == 1)
        {
            _selectedSheet = sheets[0];
        }
        else
        {
            using var sel = new SheetSelectForm(sheets);
            if (sel.ShowDialog(this) != DialogResult.OK) return;
            _selectedSheet = sel.SelectedSheet;
        }

        _excelPath = path;
        _lblExcelPath.Text      = Path.GetFileName(path) +
                                  (_selectedSheet != null ? $"  [{_selectedSheet}]" : "");
        _lblExcelPath.ForeColor = Color.Black;
        _btnGenerate.Enabled    = true;
        _dgvErrors.Visible      = false;
        _lblValidation.Visible  = false;

        SetStatus("已選擇：" + path);
        ActivityLogger.Log("載入Excel", source: Path.GetFileName(path));
    }

    void BtnExportSample_Click(object? sender, EventArgs e)
    {
        using var dlg = new SaveFileDialog
        {
            Title      = "儲存範例 Excel",
            Filter     = "Excel 活頁簿 (*.xlsx)|*.xlsx",
            FileName   = "訂購單範本.xlsx",
            DefaultExt = "xlsx",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            ExcelExporter.ExportSample(dlg.FileName);
            ActivityLogger.Log("匯出範例Excel", output: Path.GetFileName(dlg.FileName));
            if (MessageBox.Show("範例 Excel 已匯出，是否立即開啟？", "完成",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            ActivityLogger.Log("匯出範例Excel", output: Path.GetFileName(dlg.FileName),
                success: false, error: ex.Message);
            MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    async void BtnGenerate_Click(object? sender, EventArgs e)
    {
        if (_excelPath is null) return;
        string excelPath   = _excelPath;
        string? sheetName  = _selectedSheet;

        // ---- Validation ----
        List<ValidationError>? validationErrors = null;
        try
        {
            using var vs = File.OpenRead(excelPath);
            var orders = ExcelReader.ReadOrders(vs, sheetName);
            validationErrors = ValidationService.Validate(orders, _validationConfig);
            ActivityLogger.Log("驗證Excel",
                source: Path.GetFileName(excelPath),
                success: validationErrors.Count == 0,
                error: validationErrors.Count > 0
                    ? $"{validationErrors.Count} 筆問題" : null);
        }
        catch (Exception ex)
        {
            ActivityLogger.Log("驗證Excel", source: Path.GetFileName(excelPath),
                success: false, error: ex.Message);
        }

        if (validationErrors?.Count > 0)
        {
            _lblValidation.Text    = $"發現 {validationErrors.Count} 筆檢核問題：";
            _lblValidation.Visible = true;
            _dgvErrors.Rows.Clear();
            foreach (var err in validationErrors)
                _dgvErrors.Rows.Add((err.RowIndex + 1).ToString(), err.Field, err.Message);
            _dgvErrors.Visible = true;

            var res = MessageBox.Show(
                $"資料中有 {validationErrors.Count} 筆檢核問題（詳見下方列表）。\n\n是否忽略警告繼續產生 PDF？",
                "檢核警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (res == DialogResult.No) return;
        }
        else
        {
            _dgvErrors.Visible     = false;
            _lblValidation.Visible = false;
        }

        // ---- Save dialog ----
        using var saveDlg = new SaveFileDialog
        {
            Title           = "儲存 PDF",
            Filter          = "PDF 文件 (*.pdf)|*.pdf",
            FileName        = Path.GetFileNameWithoutExtension(excelPath) + "_訂購單.pdf",
            DefaultExt      = "pdf",
            OverwritePrompt = true,
        };
        if (saveDlg.ShowDialog(this) != DialogResult.OK) return;

        string savePath  = saveDlg.FileName;
        string? orderDate = _chkAutoDate.Checked ? null : _dtpOrderDate.Value.ToString("yyyy-MM-dd");
        var hospitalSnap  = _hospitalSettings;

        _btnGenerate.Enabled    = false;
        _btnSelectExcel.Enabled = false;
        _progress.Visible       = true;
        _progress.Style         = ProgressBarStyle.Marquee;
        SetStatus("正在處理，請稍候…");

        int    rowCount   = 0;
        string finalDate  = "";
        int    vendorCount = 0;
        bool   pdfOk      = false;
        string? pdfError  = null;

        try
        {
            (rowCount, finalDate, vendorCount) = await Task.Run(() =>
            {
                using var excelStream = File.OpenRead(excelPath);
                using var pdfStream   = File.Create(savePath);
                return PdfGenerator.BuildPdf(
                    excelStream, pdfStream,
                    Path.GetFileName(excelPath),
                    orderDate, hospitalSnap, sheetName);
            });
            pdfOk = true;
        }
        catch (Exception ex)
        {
            pdfError = ex.Message;
        }
        finally
        {
            _btnGenerate.Enabled    = true;
            _btnSelectExcel.Enabled = true;
            _progress.Visible       = false;
        }

        ActivityLogger.Log("產生PDF",
            source: Path.GetFileName(excelPath),
            output: Path.GetFileName(savePath),
            success: pdfOk,
            error: pdfError);

        if (!pdfOk)
        {
            SetStatus("錯誤：" + pdfError, error: true);
            MessageBox.Show(pdfError, "產生失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        SetStatus($"完成！共 {rowCount} 筆、{vendorCount} 家廠商、訂貨日期 {finalDate}。\n已儲存：{savePath}",
                  success: true);

        // PDF preview
        try
        {
            var preview = new PreviewForm(savePath);
            preview.Show(this);
        }
        catch (Exception ex)
        {
            // WebView2 not available — fall back to external viewer prompt
            if (MessageBox.Show("PDF 已產生，是否立即開啟？\n(WebView2 預覽不可用：" + ex.Message + ")",
                    "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }

    void SetStatus(string message, bool success = false, bool error = false)
    {
        _lblStatus.ForeColor = error ? Color.Crimson : success ? Color.DarkGreen : Color.DimGray;
        _lblStatus.Text = message;
    }

    // ============================================================
    // Tab 2 event handlers
    // ============================================================
    void BtnSaveRules_Click(object? sender, EventArgs e)
    {
        var rules = new List<ValidationRule>();
        foreach (DataGridViewRow row in _dgvRules.Rows)
        {
            if (row.IsNewRow) continue;
            var enabled = row.Cells["Enabled"].Value is true;
            var field   = row.Cells["Field"].Value?.ToString()   ?? "";
            var typeStr = row.Cells["RuleType"].Value?.ToString() ?? "Required";
            var param   = row.Cells["Parameter"].Value?.ToString() ?? "";
            var msg     = row.Cells["Message"].Value?.ToString()   ?? "";
            if (string.IsNullOrWhiteSpace(field)) continue;
            if (!Enum.TryParse<RuleType>(typeStr, out var rt)) rt = RuleType.Required;
            rules.Add(new ValidationRule { Enabled = enabled, Field = field, RuleType = rt, Parameter = param, Message = msg });
        }
        _validationConfig = new ValidationConfig { Rules = rules };
        AppSettings.SaveValidation(_validationConfig);
        MessageBox.Show("檢核設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetRules_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設檢核規則嗎？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _validationConfig = ValidationConfig.Default();
        AppSettings.SaveValidation(_validationConfig);
        BindRulesGrid();
    }

    // ============================================================
    // Tab 3 event handlers
    // ============================================================
    void BtnSaveHospital_Click(object? sender, EventArgs e)
    {
        _hospitalSettings = ReadHospitalFromUI();
        AppSettings.SaveHospital(_hospitalSettings);
        MessageBox.Show("PDF 樣式設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetHospital_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設值嗎？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _hospitalSettings = HospitalSettings.Default();
        AppSettings.SaveHospital(_hospitalSettings);
        BindHospitalFields();
    }
}
