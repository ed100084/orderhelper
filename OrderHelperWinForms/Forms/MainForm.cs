using System.Diagnostics;
using OrderHelperWinForms.Models;
using OrderHelperWinForms.Services;

namespace OrderHelperWinForms.Forms;

public class MainForm : Form
{
    // ---- Tab 1: 訂購單產生 ----
    readonly Button          _btnSelectExcel  = new();
    readonly Label           _lblExcelPath    = new();
    readonly DateTimePicker  _dtpOrderDate    = new();
    readonly CheckBox        _chkAutoDate     = new();
    readonly Button          _btnGenerate     = new();
    readonly ProgressBar     _progress        = new();
    readonly Label           _lblStatus       = new();
    readonly DataGridView    _dgvErrors       = new();
    readonly Label           _lblValidation   = new();

    // ---- Tab 2: 檢核設定 ----
    readonly DataGridView    _dgvRules        = new();
    readonly Button          _btnSaveRules    = new();
    readonly Button          _btnResetRules   = new();

    // ---- Tab 3: PDF 樣式設定 ----
    readonly Dictionary<string, TextBox> _hsFields = new();
    readonly Button          _btnSaveHospital = new();
    readonly Button          _btnResetHospital= new();

    string? _excelPath;
    ValidationConfig  _validationConfig  = AppSettings.LoadValidation();
    HospitalSettings  _hospitalSettings  = AppSettings.LoadHospital();

    public MainForm()
    {
        Text            = "義大醫院 藥品訂購單 PDF 產生器";
        Size            = new Size(700, 600);
        MinimumSize     = new Size(640, 540);
        StartPosition   = FormStartPosition.CenterScreen;

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildTab1());
        tabs.TabPages.Add(BuildTab2());
        tabs.TabPages.Add(BuildTab3());
        Controls.Add(tabs);

        BindRulesGrid();
        BindHospitalFields();
    }

    // ============================================================
    // Tab 1 — 訂購單產生
    // ============================================================
    TabPage BuildTab1()
    {
        var page = new TabPage("訂購單產生") { Padding = new Padding(10) };
        const int P = 10;
        var y = P;

        // Excel selection
        var lblTitle = new Label { Text = "Excel 訂購檔：", Left = P, Top = y, Width = 100, Height = 24, TextAlign = ContentAlignment.MiddleLeft };
        _btnSelectExcel.Text = "選擇檔案…"; _btnSelectExcel.Left = P + 100; _btnSelectExcel.Top = y;
        _btnSelectExcel.Width = 100; _btnSelectExcel.Height = 26; _btnSelectExcel.Click += BtnSelectExcel_Click;
        _lblExcelPath.Left = P + 210; _lblExcelPath.Top = y; _lblExcelPath.Width = 440; _lblExcelPath.Height = 26;
        _lblExcelPath.Text = "（尚未選擇）"; _lblExcelPath.ForeColor = Color.Gray; _lblExcelPath.TextAlign = ContentAlignment.MiddleLeft;
        _lblExcelPath.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 38;

        // Date picker
        var lblDate = new Label { Text = "訂貨日期：", Left = P, Top = y, Width = 100, Height = 26, TextAlign = ContentAlignment.MiddleLeft };
        _dtpOrderDate.Left = P + 100; _dtpOrderDate.Top = y; _dtpOrderDate.Width = 140; _dtpOrderDate.Height = 26;
        _dtpOrderDate.Format = DateTimePickerFormat.Short; _dtpOrderDate.Value = DateTime.Today;
        _chkAutoDate.Text = "自動從檔名/單號推算"; _chkAutoDate.Left = P + 250; _chkAutoDate.Top = y + 3;
        _chkAutoDate.Width = 180; _chkAutoDate.Height = 22; _chkAutoDate.Checked = true;
        _chkAutoDate.CheckedChanged += (_, _) => _dtpOrderDate.Enabled = !_chkAutoDate.Checked;
        _dtpOrderDate.Enabled = false;
        y += 40;

        // Separator
        var sep = new Panel { Left = P, Top = y, Height = 1, BackColor = Color.LightGray, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
        sep.Width = 640;
        y += 10;

        // Generate button
        _btnGenerate.Text = "產生 PDF"; _btnGenerate.Left = P; _btnGenerate.Top = y;
        _btnGenerate.Width = 110; _btnGenerate.Height = 34;
        _btnGenerate.Font = new Font(_btnGenerate.Font, FontStyle.Bold);
        _btnGenerate.BackColor = Color.FromArgb(32, 84, 147); _btnGenerate.ForeColor = Color.White;
        _btnGenerate.FlatStyle = FlatStyle.Flat; _btnGenerate.Enabled = false;
        _btnGenerate.Click += BtnGenerate_Click;
        y += 46;

        // Progress
        _progress.Left = P; _progress.Top = y; _progress.Height = 6; _progress.Visible = false;
        _progress.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        _progress.Width = 650;
        y += 14;

        // Status
        _lblStatus.Left = P; _lblStatus.Top = y; _lblStatus.Height = 40;
        _lblStatus.Text = "請先選擇 Excel 檔案。"; _lblStatus.ForeColor = Color.DimGray;
        _lblStatus.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        _lblStatus.Width = 650;
        y += 44;

        // Validation result area
        _lblValidation.Left = P; _lblValidation.Top = y; _lblValidation.Height = 20;
        _lblValidation.Width = 650; _lblValidation.Visible = false;
        _lblValidation.ForeColor = Color.Crimson;
        _lblValidation.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 22;

        _dgvErrors.Left = P; _dgvErrors.Top = y; _dgvErrors.Height = 160; _dgvErrors.Width = 650;
        _dgvErrors.Visible = false; _dgvErrors.ReadOnly = true; _dgvErrors.AllowUserToAddRows = false;
        _dgvErrors.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvErrors.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _dgvErrors.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Row",     HeaderText = "列",     Width = 40 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Field",   HeaderText = "欄位",   Width = 100 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Message", HeaderText = "錯誤訊息", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

        page.Controls.AddRange(new Control[]
        {
            lblTitle, _btnSelectExcel, _lblExcelPath,
            lblDate, _dtpOrderDate, _chkAutoDate,
            sep, _btnGenerate, _progress, _lblStatus,
            _lblValidation, _dgvErrors,
        });
        return page;
    }

    // ============================================================
    // Tab 2 — 檢核設定
    // ============================================================
    TabPage BuildTab2()
    {
        var page = new TabPage("檢核設定") { Padding = new Padding(10) };
        const int P = 10;

        var lbl = new Label
        {
            Text = "設定訂單資料的檢核規則（可新增/刪除/勾選啟用）：",
            Left = P, Top = P, Width = 600, Height = 20,
        };

        _dgvRules.Left = P; _dgvRules.Top = P + 24; _dgvRules.Width = 640; _dgvRules.Height = 380;
        _dgvRules.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        _dgvRules.AllowUserToAddRows = true; _dgvRules.AllowUserToDeleteRows = true;
        _dgvRules.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvRules.EditMode = DataGridViewEditMode.EditOnEnter;

        // Enabled checkbox column
        var colEnabled = new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "啟用", Width = 50, ReadOnly = false };
        _dgvRules.Columns.Add(colEnabled);

        // Field dropdown
        var colField = new DataGridViewComboBoxColumn { Name = "Field", HeaderText = "欄位", Width = 100, ReadOnly = false };
        foreach (var f in ValidationService.KnownFields) colField.Items.Add(f);
        colField.FlatStyle = FlatStyle.Flat;
        _dgvRules.Columns.Add(colField);

        // RuleType dropdown
        var colType = new DataGridViewComboBoxColumn { Name = "RuleType", HeaderText = "規則類型", Width = 100, ReadOnly = false };
        colType.Items.AddRange("Required", "Regex", "MaxLength");
        colType.FlatStyle = FlatStyle.Flat;
        _dgvRules.Columns.Add(colType);

        _dgvRules.Columns.Add(new DataGridViewTextBoxColumn { Name = "Parameter", HeaderText = "參數（正則/長度）", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        _dgvRules.Columns.Add(new DataGridViewTextBoxColumn { Name = "Message",   HeaderText = "錯誤訊息", Width = 160 });

        int btnY = 420;
        _btnSaveRules.Text = "儲存設定"; _btnSaveRules.Left = P; _btnSaveRules.Top = btnY;
        _btnSaveRules.Width = 100; _btnSaveRules.Height = 30; _btnSaveRules.Click += BtnSaveRules_Click;

        _btnResetRules.Text = "還原預設值"; _btnResetRules.Left = P + 110; _btnResetRules.Top = btnY;
        _btnResetRules.Width = 110; _btnResetRules.Height = 30; _btnResetRules.Click += BtnResetRules_Click;

        page.Controls.AddRange(new Control[] { lbl, _dgvRules, _btnSaveRules, _btnResetRules });
        return page;
    }

    // ============================================================
    // Tab 3 — PDF 樣式設定
    // ============================================================
    TabPage BuildTab3()
    {
        var page = new TabPage("PDF 樣式設定") { Padding = new Padding(10) };
        const int P = 10;

        var panel = new Panel { Left = P, Top = P, Width = 650, Height = 430, AutoScroll = true };
        panel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;

        (string key, string label)[] fields =
        {
            ("HospitalName",    "醫院名稱（標題大字）"),
            ("FormTitle",       "表單標題（藥品訂購單）"),
            ("InvoiceHeader",   "發票抬頭"),
            ("InvoiceAddress",  "發票地址"),
            ("TaxId",           "統一編號"),
            ("MedicalCode",     "醫療機構代碼"),
            ("DrugLicenseNo",   "管證字號"),
            ("DeliveryAddress", "交貨地址"),
            ("DeliveryNote",    "交貨備注（□藥庫…）"),
            ("ContactPhone",    "聯絡電話"),
            ("ContactFax",      "傳真"),
            ("Note1",           "備註1"),
            ("Note2",           "備註2"),
            ("Note3",           "備註3"),
            ("Note4",           "備註4"),
        };

        int y = 0;
        foreach (var (key, label) in fields)
        {
            var lbl = new Label { Text = label + "：", Left = 0, Top = y, Width = 150, Height = 26, TextAlign = ContentAlignment.MiddleRight };
            var tb  = new TextBox { Left = 155, Top = y, Width = 460, Height = 26, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
            _hsFields[key] = tb;
            panel.Controls.Add(lbl);
            panel.Controls.Add(tb);
            y += 32;
        }

        int btnY = y + 8;
        _btnSaveHospital.Text = "儲存設定";     _btnSaveHospital.Left = 155; _btnSaveHospital.Top = btnY;
        _btnSaveHospital.Width = 100; _btnSaveHospital.Height = 30; _btnSaveHospital.Click += BtnSaveHospital_Click;

        _btnResetHospital.Text = "還原預設值"; _btnResetHospital.Left = 265; _btnResetHospital.Top = btnY;
        _btnResetHospital.Width = 110; _btnResetHospital.Height = 30; _btnResetHospital.Click += BtnResetHospital_Click;

        panel.Controls.Add(_btnSaveHospital);
        panel.Controls.Add(_btnResetHospital);
        page.Controls.Add(panel);
        return page;
    }

    // ============================================================
    // Data binding helpers
    // ============================================================
    void BindRulesGrid()
    {
        _dgvRules.Rows.Clear();
        foreach (var r in _validationConfig.Rules)
        {
            _dgvRules.Rows.Add(r.Enabled, r.Field, r.RuleType.ToString(), r.Parameter, r.Message);
        }
    }

    void BindHospitalFields()
    {
        var hs = _hospitalSettings;
        SetField("HospitalName",    hs.HospitalName);
        SetField("FormTitle",       hs.FormTitle);
        SetField("InvoiceHeader",   hs.InvoiceHeader);
        SetField("InvoiceAddress",  hs.InvoiceAddress);
        SetField("TaxId",           hs.TaxId);
        SetField("MedicalCode",     hs.MedicalCode);
        SetField("DrugLicenseNo",   hs.DrugLicenseNo);
        SetField("DeliveryAddress", hs.DeliveryAddress);
        SetField("DeliveryNote",    hs.DeliveryNote);
        SetField("ContactPhone",    hs.ContactPhone);
        SetField("ContactFax",      hs.ContactFax);
        SetField("Note1",           hs.Note1);
        SetField("Note2",           hs.Note2);
        SetField("Note3",           hs.Note3);
        SetField("Note4",           hs.Note4);
    }

    void SetField(string key, string value)
    {
        if (_hsFields.TryGetValue(key, out var tb)) tb.Text = value;
    }

    HospitalSettings ReadHospitalFields() => new()
    {
        HospitalName    = GetField("HospitalName"),
        FormTitle       = GetField("FormTitle"),
        InvoiceHeader   = GetField("InvoiceHeader"),
        InvoiceAddress  = GetField("InvoiceAddress"),
        TaxId           = GetField("TaxId"),
        MedicalCode     = GetField("MedicalCode"),
        DrugLicenseNo   = GetField("DrugLicenseNo"),
        DeliveryAddress = GetField("DeliveryAddress"),
        DeliveryNote    = GetField("DeliveryNote"),
        ContactPhone    = GetField("ContactPhone"),
        ContactFax      = GetField("ContactFax"),
        Note1           = GetField("Note1"),
        Note2           = GetField("Note2"),
        Note3           = GetField("Note3"),
        Note4           = GetField("Note4"),
    };

    string GetField(string key) => _hsFields.TryGetValue(key, out var tb) ? tb.Text : "";

    // ============================================================
    // Event handlers — Tab 1
    // ============================================================
    void BtnSelectExcel_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title  = "選擇訂購 Excel 檔",
            Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        _excelPath = dlg.FileName;
        _lblExcelPath.Text      = Path.GetFileName(_excelPath);
        _lblExcelPath.ForeColor = Color.Black;
        _btnGenerate.Enabled    = true;
        SetStatus("已選擇：" + _excelPath);
        _dgvErrors.Visible = false;
        _lblValidation.Visible = false;
    }

    async void BtnGenerate_Click(object? sender, EventArgs e)
    {
        if (_excelPath is null) return;

        // ---- Validation ----
        List<ValidationError>? validationErrors = null;
        try
        {
            using var vs = File.OpenRead(_excelPath);
            var orders = ExcelReader.ReadOrders(vs);
            validationErrors = ValidationService.Validate(orders, _validationConfig);
        }
        catch { /* swallow; PDF generation will surface the real error */ }

        if (validationErrors?.Count > 0)
        {
            _lblValidation.Text = $"發現 {validationErrors.Count} 筆檢核問題：";
            _lblValidation.Visible = true;
            _dgvErrors.Rows.Clear();
            foreach (var err in validationErrors)
                _dgvErrors.Rows.Add((err.RowIndex + 1).ToString(), err.Field, err.Message);
            _dgvErrors.Visible = true;

            var result = MessageBox.Show(
                $"資料中有 {validationErrors.Count} 筆檢核問題（詳見下方列表）。\n\n是否忽略警告繼續產生 PDF？",
                "檢核警告",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result == DialogResult.No) return;
        }
        else
        {
            _dgvErrors.Visible = false;
            _lblValidation.Visible = false;
        }

        // ---- Save dialog ----
        using var saveDlg = new SaveFileDialog
        {
            Title           = "儲存 PDF",
            Filter          = "PDF 文件 (*.pdf)|*.pdf",
            FileName        = Path.GetFileNameWithoutExtension(_excelPath) + "_訂購單.pdf",
            DefaultExt      = "pdf",
            OverwritePrompt = true,
        };
        if (saveDlg.ShowDialog(this) != DialogResult.OK) return;

        string savePath  = saveDlg.FileName;
        string? orderDate = _chkAutoDate.Checked ? null : _dtpOrderDate.Value.ToString("yyyy-MM-dd");
        var hospitalSnap = _hospitalSettings;

        _btnGenerate.Enabled    = false;
        _btnSelectExcel.Enabled = false;
        _progress.Visible       = true;
        _progress.Style         = ProgressBarStyle.Marquee;
        SetStatus("正在處理，請稍候…");

        try
        {
            var (rowCount, finalDate, vendorCount) = await Task.Run(() =>
            {
                using var excelStream = File.OpenRead(_excelPath!);
                using var pdfStream   = File.Create(savePath);
                return PdfGenerator.BuildPdf(
                    excelStream, pdfStream,
                    Path.GetFileName(_excelPath!),
                    orderDate,
                    hospitalSnap);
            });

            SetStatus($"完成！共 {rowCount} 筆、{vendorCount} 家廠商、訂貨日期 {finalDate}。\n已儲存：{savePath}", success: true);

            if (MessageBox.Show("PDF 已產生，是否立即開啟？",
                    "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(savePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            SetStatus("錯誤：" + ex.Message, error: true);
            MessageBox.Show(ex.Message, "產生失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _btnGenerate.Enabled    = true;
            _btnSelectExcel.Enabled = true;
            _progress.Visible       = false;
        }
    }

    void SetStatus(string message, bool success = false, bool error = false)
    {
        _lblStatus.ForeColor = error ? Color.Crimson : success ? Color.DarkGreen : Color.DimGray;
        _lblStatus.Text = message;
    }

    // ============================================================
    // Event handlers — Tab 2
    // ============================================================
    void BtnSaveRules_Click(object? sender, EventArgs e)
    {
        var rules = new List<ValidationRule>();
        foreach (DataGridViewRow row in _dgvRules.Rows)
        {
            if (row.IsNewRow) continue;
            var enabled = row.Cells["Enabled"].Value is true;
            var field   = row.Cells["Field"].Value?.ToString() ?? "";
            var typeStr = row.Cells["RuleType"].Value?.ToString() ?? "Required";
            var param   = row.Cells["Parameter"].Value?.ToString() ?? "";
            var msg     = row.Cells["Message"].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(field)) continue;
            if (!Enum.TryParse<RuleType>(typeStr, out var ruleType)) ruleType = RuleType.Required;
            rules.Add(new ValidationRule { Enabled = enabled, Field = field, RuleType = ruleType, Parameter = param, Message = msg });
        }
        _validationConfig = new ValidationConfig { Rules = rules };
        AppSettings.SaveValidation(_validationConfig);
        MessageBox.Show("檢核設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetRules_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設檢核規則嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _validationConfig = ValidationConfig.Default();
        AppSettings.SaveValidation(_validationConfig);
        BindRulesGrid();
    }

    // ============================================================
    // Event handlers — Tab 3
    // ============================================================
    void BtnSaveHospital_Click(object? sender, EventArgs e)
    {
        _hospitalSettings = ReadHospitalFields();
        AppSettings.SaveHospital(_hospitalSettings);
        MessageBox.Show("PDF 樣式設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetHospital_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設值嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _hospitalSettings = HospitalSettings.Default();
        AppSettings.SaveHospital(_hospitalSettings);
        BindHospitalFields();
    }
}
