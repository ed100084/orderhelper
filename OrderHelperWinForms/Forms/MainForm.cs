using System.Diagnostics;
using System.Text;
using OrderHelperWinForms.Models;
using OrderHelperWinForms.Services;

namespace OrderHelperWinForms.Forms;

public class MainForm : Form
{
    static readonly Color AppBack       = Color.FromArgb(246, 248, 251);
    static readonly Color PanelBack     = Color.White;
    static readonly Color PrimaryBlue   = Color.FromArgb(37, 99, 235);
    static readonly Color PrimaryHover  = Color.FromArgb(29, 78, 216);
    static readonly Color BorderGray    = Color.FromArgb(214, 220, 229);
    static readonly Color TextMuted     = Color.FromArgb(100, 116, 139);
    static readonly Color TextStrong    = Color.FromArgb(15, 23, 42);

    static readonly string[] RuleChoices =
    {
        "必填",
        "必須是數字",
        "必須大於 0",
        "必須是整數",
        "必須是正整數",
        "不可為 0",
        "電話/傳真格式",
        "統一編號 8 碼",
        "長度不超過 20 字",
        "長度不超過 50 字",
        "長度不超過 100 字",
    };

    // ---- Tab 1 controls ----
    readonly Button         _btnSelectExcel   = new();
    readonly Label          _lblExcelPath     = new();
    readonly DateTimePicker _dtpOrderDate     = new();
    readonly CheckBox       _chkAutoDate      = new();
    readonly Button         _btnGenerate      = new();
    readonly CheckBox       _chkSplitByVendor = new();
    readonly CheckBox       _chkNextMonthInvoice = new();
    readonly ProgressBar    _progress        = new();
    readonly Label          _lblStatus       = new();
    readonly Label          _lblValidation   = new();
    readonly DataGridView   _dgvErrors       = new();
    readonly ToolTip        _toolTip         = new();

    // ---- Tab 2 controls ----
    readonly DataGridView   _dgvRules        = new();
    readonly Button         _btnSaveRules    = new();
    readonly Button         _btnResetRules   = new();

    // ---- Tab 3 controls ----
    readonly Dictionary<string, TextBox> _hsFields     = new();
    readonly Button         _btnSaveHospital  = new();
    readonly Button         _btnResetHospital = new();
    readonly RadioButton     _rdoOutputExcelDir = new();
    readonly RadioButton     _rdoOutputDefaultDir = new();
    readonly RadioButton     _rdoOutputAsk = new();
    readonly TextBox         _txtDefaultPdfDir = new();
    readonly Button          _btnBrowsePdfDir  = new();
    readonly NumericUpDown   _nudMaxRows       = new();

    // ---- State ----
    string?          _excelPath;
    string?          _selectedSheet;
    ValidationConfig _validationConfig  = AppSettings.LoadValidation();
    HospitalSettings _hospitalSettings  = AppSettings.LoadHospital();
    GeneralSettings  _generalSettings   = AppSettings.LoadGeneral();
    PreviewForm?     _previewForm;
    bool             _rulesDirty;
    bool             _settingsDirty;
    bool             _suspendDirtyTracking;

    // Captured for Tab 3 resize
    List<GroupBox>?  _tab3GroupBoxes;
    Panel?           _tab3ScrollPanel;
    const int        Tab3LblW = 140;
    const int        Tab3Pad  = 8;

    public MainForm()
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        string verSuffix = ver != null ? $" v{ver.Major}.{ver.Minor}.{ver.Build}" : "";
        Text          = "義大醫院 藥品訂購單 PDF 產生器" + verSuffix;
        Size          = new Size(740, 680);
        MinimumSize   = new Size(660, 600);
        StartPosition = FormStartPosition.CenterScreen;
        AllowDrop     = true;
        Font          = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular);
        BackColor     = AppBack;

        var menu = BuildMenu();
        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildTab1());
        tabs.TabPages.Add(BuildTab2());
        tabs.TabPages.Add(BuildTab3());

        // Fill control must be added BEFORE Top/Bottom controls; otherwise the
        // MenuStrip (z-order 0) renders on top of the TabControl's tab-header strip.
        Controls.Add(tabs);   // Fill — added first, lower visual z-order
        Controls.Add(menu);   // Top  — added second, higher visual z-order
        MainMenuStrip = menu;

        // Dirty-flag tracking
        _dgvRules.CellValueChanged += (_, e) =>
        {
            if (!_suspendDirtyTracking)
            {
                _rulesDirty = true;
                ApplyDefaultRuleMessage(e.RowIndex);
            }
        };
        _dgvRules.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_dgvRules.IsCurrentCellDirty)
                _dgvRules.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        _dgvRules.RowsAdded        += (_, _) => { if (!_suspendDirtyTracking) _rulesDirty = true; };
        _dgvRules.RowsRemoved      += (_, _) => { if (!_suspendDirtyTracking) _rulesDirty = true; };
        _dgvRules.DataError        += (_, e) => e.ThrowException = false;

        foreach (var tb in _hsFields.Values)
            tb.TextChanged += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; };
        _rdoOutputExcelDir.CheckedChanged += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; UpdateOutputDirectoryControls(); };
        _rdoOutputDefaultDir.CheckedChanged += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; UpdateOutputDirectoryControls(); };
        _rdoOutputAsk.CheckedChanged += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; UpdateOutputDirectoryControls(); };
        _txtDefaultPdfDir.TextChanged  += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; };
        _nudMaxRows.ValueChanged       += (_, _) => { if (!_suspendDirtyTracking) _settingsDirty = true; };

        _suspendDirtyTracking = true;
        BindRulesGrid();
        BindHospitalFields();
        BindGeneralSettings();
        _suspendDirtyTracking = false;

        DragEnter += MainForm_DragEnter;
        DragDrop  += MainForm_DragDrop;
    }

    static void StylePage(TabPage page)
    {
        page.BackColor = AppBack;
        page.ForeColor = TextStrong;
    }

    static void StyleButton(Button button, bool primary = false)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.FlatAppearance.BorderColor = BorderGray;
        button.BackColor = primary ? PrimaryBlue : PanelBack;
        button.ForeColor = primary ? Color.White : TextStrong;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
        if (primary)
        {
            button.FlatAppearance.MouseOverBackColor = PrimaryHover;
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(30, 64, 175);
        }
    }

    static void StyleGrid(DataGridView grid)
    {
        grid.BackgroundColor = PanelBack;
        grid.BorderStyle = BorderStyle.None;
        grid.GridColor = BorderGray;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = TextStrong;
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(241, 245, 249);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
        grid.DefaultCellStyle.SelectionForeColor = TextStrong;
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
        grid.RowTemplate.Height = 28;
    }

    static void StyleGroupBox(GroupBox box)
    {
        box.BackColor = PanelBack;
        box.ForeColor = TextStrong;
        box.Padding = new Padding(8);
    }

    // ============================================================
    // Menu
    // ============================================================
    MenuStrip BuildMenu()
    {
        var menu     = new MenuStrip();
        menu.BackColor = PanelBack;
        menu.ForeColor = TextStrong;
        var miTools  = new ToolStripMenuItem("工具(&T)");
        var miLog    = new ToolStripMenuItem("檢視操作記錄…");
        var miSep    = new ToolStripSeparator();
        var miAbout  = new ToolStripMenuItem("關於 OrderHelper…");
        miLog.Click   += (_, _) => new LogViewerForm().ShowDialog(this);
        miAbout.Click += (_, _) => new AboutForm().ShowDialog(this);
        miTools.DropDownItems.Add(miLog);
        miTools.DropDownItems.Add(miSep);
        miTools.DropDownItems.Add(miAbout);
        menu.Items.Add(miTools);
        return menu;
    }

    // ============================================================
    // Tab 1 — 訂購單產生
    // ============================================================
    TabPage BuildTab1()
    {
        var page = new TabPage("訂購單產生") { Padding = new Padding(10) };
        StylePage(page);
        const int P = 12;
        var y = P;

        var lblTitle = new Label
        {
            Text = "Excel 訂購檔：", Left = P, Top = y, Width = 100, Height = 26,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        _btnSelectExcel.Text = "選擇檔案…";
        _btnSelectExcel.SetBounds(P + 100, y, 100, 26);
        StyleButton(_btnSelectExcel);
        _btnSelectExcel.Click += BtnSelectExcel_Click;

        _lblExcelPath.SetBounds(P + 208, y, 0, 26);
        _lblExcelPath.Text      = "（尚未選擇，可拖放 .xlsx 至視窗）";
        _lblExcelPath.ForeColor = TextMuted;
        _lblExcelPath.TextAlign = ContentAlignment.MiddleLeft;
        _lblExcelPath.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 34;

        var lblDate = new Label
        {
            Text = "訂貨日期：", Left = P, Top = y, Width = 100, Height = 26,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        _dtpOrderDate.SetBounds(P + 100, y, 150, 26);
        _dtpOrderDate.Format  = DateTimePickerFormat.Short;
        _dtpOrderDate.Value   = DateTime.Today;
        _dtpOrderDate.Enabled = false;

        _chkAutoDate.Text    = "自動從檔名/單號推算";
        _chkAutoDate.SetBounds(P + 258, y + 1, 180, 24);
        _chkAutoDate.Checked = true;
        _chkAutoDate.CheckedChanged += (_, _) => _dtpOrderDate.Enabled = !_chkAutoDate.Checked && !_progress.Visible;
        y += 38;

        var sep = new Panel
        {
            Left = P, Top = y, Height = 1, BackColor = BorderGray,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        y += 10;

        _btnGenerate.Text      = "產生 PDF";
        _btnGenerate.SetBounds(P, y, 120, 36);
        _btnGenerate.Font      = new Font(_btnGenerate.Font, FontStyle.Bold);
        StyleButton(_btnGenerate, primary: true);
        _btnGenerate.Enabled   = false;
        _btnGenerate.Click    += BtnGenerate_Click;

        _chkSplitByVendor.Text      = "分廠商輸出（每家廠商一個 PDF）";
        _chkSplitByVendor.SetBounds(P + 128, y + 6, 285, 24);
        _chkSplitByVendor.Checked   = false;

        _chkNextMonthInvoice.Text      = "請開立下個月發票";
        _chkNextMonthInvoice.SetBounds(P + 420, y + 3, 250, 30);
        _chkNextMonthInvoice.Checked   = false;
        _chkNextMonthInvoice.ForeColor = Color.Crimson;
        _chkNextMonthInvoice.BackColor = Color.FromArgb(255, 235, 238);
        _chkNextMonthInvoice.FlatStyle = FlatStyle.Flat;
        _chkNextMonthInvoice.Padding   = new Padding(4, 0, 0, 0);
        _chkNextMonthInvoice.Font      = new Font(_chkNextMonthInvoice.Font.FontFamily, 10.5F, FontStyle.Bold);
        y += 48;

        _progress.SetBounds(P, y, 0, 6);
        _progress.Visible = false;
        _progress.Anchor  = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 14;

        _lblStatus.SetBounds(P, y, 0, 40);
        _lblStatus.Text      = "請先選擇 Excel 檔案，或將 .xlsx 拖放至視窗。";
        _lblStatus.ForeColor = TextMuted;
        _lblStatus.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 44;

        _lblValidation.SetBounds(P, y, 0, 20);
        _lblValidation.Visible   = false;
        _lblValidation.ForeColor = Color.Crimson;
        _lblValidation.Anchor    = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        y += 22;

        _dgvErrors.SetBounds(P, y, 0, 160);
        _dgvErrors.Visible              = false;
        _dgvErrors.ReadOnly             = true;
        _dgvErrors.AllowUserToAddRows   = false;
        _dgvErrors.RowHeadersVisible    = false;
        _dgvErrors.SelectionMode        = DataGridViewSelectionMode.FullRowSelect;
        _dgvErrors.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvErrors.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        StyleGrid(_dgvErrors);
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "OrderNo", HeaderText = "訂購單號", Width = 120 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn { Name = "Field",   HeaderText = "欄位",     Width = 100 });
        _dgvErrors.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Message", HeaderText = "錯誤訊息",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
        });

        page.Controls.AddRange(new Control[]
        {
            lblTitle, _btnSelectExcel, _lblExcelPath,
            lblDate, _dtpOrderDate, _chkAutoDate,
            sep, _btnGenerate, _chkSplitByVendor, _chkNextMonthInvoice, _progress, _lblStatus,
            _lblValidation, _dgvErrors,
        });

        page.Resize += (_, _) =>
        {
            int w = page.ClientSize.Width - P * 2;
            _lblExcelPath.Width  = page.ClientSize.Width - P - 208 - P;
            sep.Width            = w;
            _progress.Width      = w;
            _lblStatus.Width     = w;
            _lblValidation.Width = w;
            _dgvErrors.Width     = w;
        };

        return page;
    }

    // ============================================================
    // Tab 2 — 檢核設定
    // ============================================================
    TabPage BuildTab2()
    {
        var page = new TabPage("檢核設定") { Padding = new Padding(10) };
        StylePage(page);
        const int P = 12;

        var lbl = new Label
        {
            Text  = "設定訂單資料的檢核規則（選欄位與檢核方式即可）：",
            Left  = P, Top = P, Width = 650, Height = 22,
            ForeColor = TextMuted,
        };

        _dgvRules.SetBounds(P, P + 26, 0, 0);
        _dgvRules.AllowUserToAddRows    = true;
        _dgvRules.AllowUserToDeleteRows = true;
        _dgvRules.EditMode              = DataGridViewEditMode.EditOnEnter;
        _dgvRules.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _dgvRules.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        StyleGrid(_dgvRules);

        _dgvRules.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "啟用", Width = 50 });

        var colField = new DataGridViewComboBoxColumn { Name = "Field", HeaderText = "欄位", Width = 110, FlatStyle = FlatStyle.Flat };
        foreach (var f in ValidationService.KnownFields) colField.Items.Add(f);
        _dgvRules.Columns.Add(colField);

        var colType = new DataGridViewComboBoxColumn
        {
            Name = "RuleType",
            HeaderText = "檢核方式",
            Width = 170,
            FlatStyle = FlatStyle.Flat,
        };
        colType.Items.AddRange(RuleChoices.Cast<object>().ToArray());
        _dgvRules.Columns.Add(colType);

        _dgvRules.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Message",
            HeaderText = "提示訊息",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
        });

        _btnSaveRules.Text   = "儲存設定";
        _btnResetRules.Text  = "還原預設值";
        _btnSaveRules.Size   = _btnResetRules.Size = new Size(110, 30);
        StyleButton(_btnSaveRules, primary: true);
        StyleButton(_btnResetRules);
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
        StylePage(page);
        const int P    = 12;
        const int LblW = Tab3LblW;
        const int Pad  = Tab3Pad;

        var scroll = new Panel { Left = P, Top = P, AutoScroll = true, BackColor = AppBack };
        scroll.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
        _tab3ScrollPanel = scroll;

        var groupBoxes = new List<GroupBox>();
        _tab3GroupBoxes = groupBoxes;

        int y = 0;

        // ---- "一般設定" GroupBox ----
        var gbGeneral = new GroupBox { Text = "一般設定", Left = 0, Top = y, Width = 680, Height = 146 };
        StyleGroupBox(gbGeneral);
        var lblOutputMode = new Label
        {
            Text = "輸出位置：", Left = Pad, Top = 24, Width = LblW, Height = 24,
            TextAlign = ContentAlignment.MiddleRight,
        };
        _rdoOutputExcelDir.Text = "Excel 同目錄";
        _rdoOutputExcelDir.SetBounds(Pad + LblW + 4, 24, 110, 24);
        _rdoOutputDefaultDir.Text = "固定預設目錄";
        _rdoOutputDefaultDir.SetBounds(Pad + LblW + 120, 24, 125, 24);
        _rdoOutputAsk.Text = "每次詢問";
        _rdoOutputAsk.SetBounds(Pad + LblW + 254, 24, 100, 24);
        var lblPdfDir = new Label
        {
            Text = "預設儲存目錄：", Left = Pad, Top = 58, Width = LblW, Height = 26,
            TextAlign = ContentAlignment.MiddleRight,
        };
        _txtDefaultPdfDir.SetBounds(Pad + LblW + 4, 58, 350, 26);
        _btnBrowsePdfDir.Text = "瀏覽…";
        _btnBrowsePdfDir.SetBounds(Pad + LblW + 4 + 330, 58, 80, 26);
        StyleButton(_btnBrowsePdfDir);
        _btnBrowsePdfDir.Click += BtnBrowsePdfDir_Click;
        var lblMaxRows = new Label
        {
            Text = "每頁最多筆數：", Left = Pad, Top = 96, Width = LblW, Height = 22,
            TextAlign = ContentAlignment.MiddleRight,
        };
        _nudMaxRows.SetBounds(Pad + LblW + 4, 96, 70, 22);
        _nudMaxRows.Minimum  = 1;
        _nudMaxRows.Maximum  = 99;
        _nudMaxRows.Value    = 12;
        var lblMaxRowsHint = new Label
        {
            Text = "筆（預設 10，超長品名時實際筆數可能略少）",
            Left = Pad + LblW + 4 + 74, Top = 96, Width = 320, Height = 22,
            TextAlign = ContentAlignment.MiddleLeft, ForeColor = TextMuted,
        };
        gbGeneral.Controls.AddRange(new Control[]
        {
            lblOutputMode, _rdoOutputExcelDir, _rdoOutputDefaultDir, _rdoOutputAsk,
            lblPdfDir, _txtDefaultPdfDir, _btnBrowsePdfDir,
            lblMaxRows, _nudMaxRows, lblMaxRowsHint,
        });
        scroll.Controls.Add(gbGeneral);
        groupBoxes.Add(gbGeneral);
        y += gbGeneral.Height + Pad;

        // Helper to build a labeled TextBox row inside a GroupBox
        void AddField(GroupBox gb, string key, string label, ref int gy)
        {
            var lbl = new Label
            {
                Text = label + "：", Left = Pad, Top = gy, Width = LblW, Height = 26,
                TextAlign = ContentAlignment.MiddleRight,
            };
            var tb = new TextBox { Left = Pad + LblW + 4, Top = gy, Width = 400, Height = 26 };
            tb.BorderStyle = BorderStyle.FixedSingle;
            tb.BackColor = Color.White;
            tb.ForeColor = TextStrong;
            _hsFields[key] = tb;
            gb.Controls.Add(lbl);
            gb.Controls.Add(tb);
            gy += 32;
        }

        // ---- "標題" GroupBox ----
        var gbTitle = new GroupBox { Text = "標題", Left = 0, Top = y, Width = 680 };
        StyleGroupBox(gbTitle);
        int gy = 22;
        AddField(gbTitle, "HospitalName", "醫院名稱（標題大字）", ref gy);
        AddField(gbTitle, "FormTitle",    "表單標題",             ref gy);
        gbTitle.Height = gy + Pad;
        scroll.Controls.Add(gbTitle);
        groupBoxes.Add(gbTitle);
        y += gbTitle.Height + Pad;

        // ---- "發票/法規資訊" GroupBox ----
        var gbInvoice = new GroupBox { Text = "發票/法規資訊", Left = 0, Top = y, Width = 680 };
        StyleGroupBox(gbInvoice);
        gy = 22;
        AddField(gbInvoice, "InvoiceHeader",  "發票抬頭",     ref gy);
        AddField(gbInvoice, "InvoiceAddress", "發票地址",     ref gy);
        AddField(gbInvoice, "TaxId",          "統一編號",     ref gy);
        AddField(gbInvoice, "MedicalCode",    "醫療機構代碼", ref gy);
        AddField(gbInvoice, "DrugLicenseNo",  "管證字號",     ref gy);
        gbInvoice.Height = gy + Pad;
        scroll.Controls.Add(gbInvoice);
        groupBoxes.Add(gbInvoice);
        y += gbInvoice.Height + Pad;

        // ---- "交貨與備註" GroupBox ----
        var gbDelivery = new GroupBox { Text = "交貨與備註", Left = 0, Top = y, Width = 680 };
        StyleGroupBox(gbDelivery);
        gy = 22;
        AddField(gbDelivery, "DeliveryAddress", "交貨地址", ref gy);
        AddField(gbDelivery, "DeliveryNote",    "交貨備注", ref gy);
        AddField(gbDelivery, "ContactPhone",    "聯絡電話", ref gy);
        AddField(gbDelivery, "ContactFax",      "傳真",     ref gy);
        AddField(gbDelivery, "Note1", "備註1", ref gy);
        AddField(gbDelivery, "Note2", "備註2", ref gy);
        AddField(gbDelivery, "Note3", "備註3", ref gy);
        AddField(gbDelivery, "Note4", "備註4", ref gy);
        gbDelivery.Height = gy + Pad;
        scroll.Controls.Add(gbDelivery);
        groupBoxes.Add(gbDelivery);
        y += gbDelivery.Height + Pad;

        _btnSaveHospital.Text   = "儲存所有設定";
        _btnResetHospital.Text  = "還原預設值";
        _btnSaveHospital.SetBounds(Pad, y, 120, 30);
        _btnResetHospital.SetBounds(Pad + 128, y, 120, 30);
        StyleButton(_btnSaveHospital, primary: true);
        StyleButton(_btnResetHospital);
        _btnSaveHospital.Click  += BtnSaveHospital_Click;
        _btnResetHospital.Click += BtnResetHospital_Click;
        scroll.Controls.Add(_btnSaveHospital);
        scroll.Controls.Add(_btnResetHospital);

        page.Controls.Add(scroll);

        page.Resize += (_, _) =>
        {
            int w   = page.ClientSize.Width - P * 2;
            int h   = page.ClientSize.Height - P * 2;
            scroll.Width  = w;
            scroll.Height = h;

            int gbW = Math.Max(400, w - 4);
            foreach (var gb in groupBoxes) gb.Width = gbW;

            int tbW = Math.Max(60, gbW - Pad - LblW - 4 - Pad - 2);
            foreach (var tb in _hsFields.Values) tb.Width = tbW;

            // "一般設定" browse-button row
            int btnLeft = Math.Max(Pad + LblW + 4 + 64, gbW - Pad - 82);
            _btnBrowsePdfDir.Left   = btnLeft;
            _txtDefaultPdfDir.Width = Math.Max(60, btnLeft - (Pad + LblW + 4) - 4);
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
            _dgvRules.Rows.Add(r.Enabled, r.Field, RuleTypeToDisplay(r), r.Message);
    }

    void BindHospitalFields()
    {
        var hs = _hospitalSettings;
        Set("HospitalName",    hs.HospitalName);   Set("FormTitle",       hs.FormTitle);
        Set("InvoiceHeader",   hs.InvoiceHeader);  Set("InvoiceAddress",  hs.InvoiceAddress);
        Set("TaxId",           hs.TaxId);          Set("MedicalCode",     hs.MedicalCode);
        Set("DrugLicenseNo",   hs.DrugLicenseNo);  Set("DeliveryAddress", hs.DeliveryAddress);
        Set("DeliveryNote",    hs.DeliveryNote);   Set("ContactPhone",    hs.ContactPhone);
        Set("ContactFax",      hs.ContactFax);
        Set("Note1", hs.Note1); Set("Note2", hs.Note2);
        Set("Note3", hs.Note3); Set("Note4", hs.Note4);
    }

    void BindGeneralSettings()
    {
        var mode = ResolveOutputMode(_generalSettings);
        _rdoOutputExcelDir.Checked = mode == PdfOutputMode.ExcelDirectory;
        _rdoOutputDefaultDir.Checked = mode == PdfOutputMode.DefaultDirectory;
        _rdoOutputAsk.Checked = mode == PdfOutputMode.AskEveryTime;
        _txtDefaultPdfDir.Text  = _generalSettings.DefaultPdfDirectory ?? "";
        _nudMaxRows.Value       = Math.Clamp(_generalSettings.MaxRowsPerPage, 1, 99);
        UpdateOutputDirectoryControls();
    }

    static PdfOutputMode ResolveOutputMode(GeneralSettings settings)
    {
        if (settings.PdfOutputMode != default || settings.AutoSaveSameDir)
            return settings.PdfOutputMode;
        return PdfOutputMode.AskEveryTime;
    }

    PdfOutputMode SelectedOutputMode()
    {
        if (_rdoOutputDefaultDir.Checked) return PdfOutputMode.DefaultDirectory;
        if (_rdoOutputAsk.Checked) return PdfOutputMode.AskEveryTime;
        return PdfOutputMode.ExcelDirectory;
    }

    void UpdateOutputDirectoryControls()
    {
        bool enabled = _rdoOutputDefaultDir.Checked;
        _txtDefaultPdfDir.Enabled = enabled;
        _btnBrowsePdfDir.Enabled = enabled;
    }

    void Set(string key, string val)    { if (_hsFields.TryGetValue(key, out var tb)) tb.Text = val; }
    string Get(string key)              => _hsFields.TryGetValue(key, out var tb) ? tb.Text : "";

    HospitalSettings ReadHospitalFromUI() => new()
    {
        HospitalName    = Get("HospitalName"),   FormTitle       = Get("FormTitle"),
        InvoiceHeader   = Get("InvoiceHeader"),  InvoiceAddress  = Get("InvoiceAddress"),
        TaxId           = Get("TaxId"),          MedicalCode     = Get("MedicalCode"),
        DrugLicenseNo   = Get("DrugLicenseNo"),  DeliveryAddress = Get("DeliveryAddress"),
        DeliveryNote    = Get("DeliveryNote"),   ContactPhone    = Get("ContactPhone"),
        ContactFax      = Get("ContactFax"),
        Note1 = Get("Note1"), Note2 = Get("Note2"),
        Note3 = Get("Note3"), Note4 = Get("Note4"),
    };

    GeneralSettings ReadGeneralFromUI() => new()
    {
        LastExcelDirectory = _generalSettings.LastExcelDirectory,
        PdfOutputMode = SelectedOutputMode(),
        AutoSaveSameDir = SelectedOutputMode() != PdfOutputMode.AskEveryTime,
        DefaultPdfDirectory = string.IsNullOrWhiteSpace(_txtDefaultPdfDir.Text)
            ? null
            : _txtDefaultPdfDir.Text.Trim(),
        MaxRowsPerPage = (int)_nudMaxRows.Value,
    };

    static string RuleTypeToDisplay(ValidationRule rule)
    {
        if (rule.RuleType == RuleType.Regex)
        {
            string p = rule.Parameter.Trim();
            if (p is @"^\d+(\.\d+)?$" or @"^\d+(\.\d+)?") return "必須是數字";
            if (p is @"^\d+$") return "必須是整數";
            if (p is @"^\d{8}$") return "統一編號 8 碼";
            if (p is @"^[0-9#()\-\.\s+]+$") return "電話/傳真格式";
        }
        return RuleTypeToDisplay(rule.RuleType);
    }

    static string RuleTypeToDisplay(RuleType rt) => rt switch
    {
        RuleType.Required        => "必填",
        RuleType.Number          => "必須是數字",
        RuleType.PositiveNumber  => "必須大於 0",
        RuleType.Integer         => "必須是整數",
        RuleType.PositiveInteger => "必須是正整數",
        RuleType.NotZero         => "不可為 0",
        RuleType.PhoneFax        => "電話/傳真格式",
        RuleType.TaxId8          => "統一編號 8 碼",
        RuleType.MaxLength20     => "長度不超過 20 字",
        RuleType.MaxLength50     => "長度不超過 50 字",
        RuleType.MaxLength100    => "長度不超過 100 字",
        RuleType.MaxLength       => "長度不超過 100 字",
        RuleType.Regex           => "必須是數字",
        _                        => "必填",
    };

    static RuleType DisplayToRuleType(string s) => s switch
    {
        "必填"              => RuleType.Required,
        "必須是數字"        => RuleType.Number,
        "必須大於 0"        => RuleType.PositiveNumber,
        "必須是整數"        => RuleType.Integer,
        "必須是正整數"      => RuleType.PositiveInteger,
        "不可為 0"          => RuleType.NotZero,
        "電話/傳真格式"     => RuleType.PhoneFax,
        "統一編號 8 碼"     => RuleType.TaxId8,
        "長度不超過 20 字"  => RuleType.MaxLength20,
        "長度不超過 50 字"  => RuleType.MaxLength50,
        "長度不超過 100 字" => RuleType.MaxLength100,
        _                   => Enum.TryParse<RuleType>(s, out var rt) ? rt : RuleType.Required,
    };

    void ApplyDefaultRuleMessage(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _dgvRules.Rows.Count) return;
        var row = _dgvRules.Rows[rowIndex];
        if (row.IsNewRow) return;

        string field = row.Cells["Field"].Value?.ToString() ?? "";
        string typeStr = row.Cells["RuleType"].Value?.ToString() ?? "";
        string current = row.Cells["Message"].Value?.ToString() ?? "";
        if (string.IsNullOrWhiteSpace(field) || string.IsNullOrWhiteSpace(typeStr)) return;

        var newMessage = ValidationService.DefaultMessage(field, DisplayToRuleType(typeStr));
        if (string.IsNullOrWhiteSpace(current) || IsKnownDefaultMessage(current))
        {
            _suspendDirtyTracking = true;
            row.Cells["Message"].Value = newMessage;
            _suspendDirtyTracking = false;
            _rulesDirty = true;
        }
    }

    static bool IsKnownDefaultMessage(string message)
        => ValidationService.KnownFields.Any(field =>
            RuleChoices.Any(choice =>
                string.Equals(message, ValidationService.DefaultMessage(field, DisplayToRuleType(choice)),
                    StringComparison.Ordinal)));

    void SetStatus(string msg, bool success = false, bool error = false)
    {
        _lblStatus.ForeColor = error ? Color.Crimson : success ? Color.DarkGreen : Color.DimGray;
        _lblStatus.Text      = msg;
    }

    void SetBusy(bool busy, string? msg = null)
    {
        _btnGenerate.Enabled    = !busy;
        _btnSelectExcel.Enabled = !busy;
        _chkSplitByVendor.Enabled = !busy;
        _chkNextMonthInvoice.Enabled = !busy;
        _chkAutoDate.Enabled = !busy;
        _dtpOrderDate.Enabled = !busy && !_chkAutoDate.Checked;
        _dgvRules.Enabled = !busy;
        _btnSaveRules.Enabled = !busy;
        _btnResetRules.Enabled = !busy;
        SetSettingsControlsEnabled(!busy);
        _progress.Visible       = busy;
        _progress.Style         = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
        if (msg != null) SetStatus(msg);
    }

    void SetSettingsControlsEnabled(bool enabled)
    {
        _rdoOutputExcelDir.Enabled = enabled;
        _rdoOutputDefaultDir.Enabled = enabled;
        _rdoOutputAsk.Enabled = enabled;
        _txtDefaultPdfDir.Enabled = enabled && _rdoOutputDefaultDir.Checked;
        _btnBrowsePdfDir.Enabled = enabled && _rdoOutputDefaultDir.Checked;
        _nudMaxRows.Enabled = enabled;
        _btnSaveHospital.Enabled = enabled;
        _btnResetHospital.Enabled = enabled;
        foreach (var tb in _hsFields.Values) tb.Enabled = enabled;
    }

    // ============================================================
    // Drag & Drop  (H5)
    // ============================================================
    void MainForm_DragEnter(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        e.Effect = files?.Any(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) == true
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    void MainForm_DragDrop(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        var xlsx  = files?.FirstOrDefault(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
        if (xlsx != null) LoadExcelFile(xlsx);
    }

    // ============================================================
    // Tab 1 event handlers
    // ============================================================
    void BtnSelectExcel_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title            = "選擇訂購 Excel 檔",
            Filter           = "Excel 活頁簿 (*.xlsx)|*.xlsx",
            InitialDirectory = _generalSettings.LastExcelDirectory
                               ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            LoadExcelFile(dlg.FileName);
    }

    void LoadExcelFile(string path)
    {
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

        string? selectedSheet;
        if (sheets.Count == 1)
        {
            selectedSheet = sheets[0];
        }
        else
        {
            using var sel = new SheetSelectForm(sheets);
            if (sel.ShowDialog(this) != DialogResult.OK) return;
            selectedSheet = sel.SelectedSheet;
        }

        _excelPath     = path;
        _selectedSheet = selectedSheet;

        string displayName = Path.GetFileName(path)
                           + (selectedSheet != null ? $"  [{selectedSheet}]" : "");
        _lblExcelPath.Text      = displayName;
        _lblExcelPath.ForeColor = Color.Black;
        _toolTip.SetToolTip(_lblExcelPath, path);

        _btnGenerate.Enabled   = true;
        _dgvErrors.Visible     = false;
        _lblValidation.Visible = false;

        // Remember last directory (H6)
        _generalSettings.LastExcelDirectory = Path.GetDirectoryName(path);
        AppSettings.SaveGeneral(_generalSettings);

        SetStatus("已選擇：" + Path.GetFileName(path));
        ActivityLogger.Log("載入Excel", source: Path.GetFileName(path));
    }

    string? GetOutputDirectory(string excelPath, GeneralSettings settings, string title)
    {
        string excelDir = Path.GetDirectoryName(excelPath)
                          ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        if (settings.PdfOutputMode == PdfOutputMode.ExcelDirectory)
            return excelDir;

        if (settings.PdfOutputMode == PdfOutputMode.DefaultDirectory
            && !string.IsNullOrWhiteSpace(settings.DefaultPdfDirectory))
            return settings.DefaultPdfDirectory;

        using var dirDlg = new FolderBrowserDialog
        {
            Description            = title,
            UseDescriptionForTitle = true,
            SelectedPath           = !string.IsNullOrWhiteSpace(settings.DefaultPdfDirectory)
                                     ? settings.DefaultPdfDirectory
                                     : excelDir,
        };
        return dirDlg.ShowDialog(this) == DialogResult.OK ? dirDlg.SelectedPath : null;
    }

    string? GetSinglePdfSavePath(string excelPath, GeneralSettings settings)
    {
        string fileName = Path.GetFileNameWithoutExtension(excelPath) + "_訂購單.pdf";
        if (settings.PdfOutputMode == PdfOutputMode.AskEveryTime)
        {
            string initDir = !string.IsNullOrWhiteSpace(settings.DefaultPdfDirectory)
                             ? settings.DefaultPdfDirectory
                             : Path.GetDirectoryName(excelPath)
                               ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            using var saveDlg = new SaveFileDialog
            {
                Title            = "儲存 PDF",
                Filter           = "PDF 文件 (*.pdf)|*.pdf",
                FileName         = fileName,
                DefaultExt       = "pdf",
                OverwritePrompt  = true,
                InitialDirectory = initDir,
            };
            return saveDlg.ShowDialog(this) == DialogResult.OK ? saveDlg.FileName : null;
        }

        var dir = GetOutputDirectory(excelPath, settings, "選擇 PDF 輸出目錄");
        return dir == null ? null : Path.Combine(dir, fileName);
    }

    void EnsureOutputDirectory(string filePath)
    {
        string? dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(dir))
            Directory.CreateDirectory(dir);
    }

    async void BtnGenerate_Click(object? sender, EventArgs e)
    {
        if (_excelPath is null) return;
        string  excelPath  = _excelPath;
        string? sheetName  = _selectedSheet;

        // Step 1: Read Excel once (H2)
        List<OrderRow>? orders    = null;
        string?         readError = null;

        SetBusy(true, "正在讀取 Excel…");
        try
        {
            orders = await Task.Run(() =>
            {
                using var s = File.OpenRead(excelPath);
                return ExcelReader.ReadOrders(s, sheetName);
            });
        }
        catch (Exception ex) { readError = ex.Message; }
        finally                { SetBusy(false); }

        if (orders == null)
        {
            ActivityLogger.Log("讀取Excel", source: Path.GetFileName(excelPath),
                success: false, error: readError);
            SetStatus("讀取失敗：" + readError, error: true);
            MessageBox.Show(readError, "讀取失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (orders.Count == 0)
        {
            SetStatus("Excel 沒有可讀取的訂單資料。", error: true);
            MessageBox.Show("Excel 沒有可輸出的訂單資料。", "無資料",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Step 2: Validate
        var validationErrors = ValidationService.Validate(orders, _validationConfig);
        ActivityLogger.Log("驗證Excel", source: Path.GetFileName(excelPath),
            success: validationErrors.Count == 0,
            error:   validationErrors.Count > 0 ? $"{validationErrors.Count} 筆問題" : null);

        // Step 3: Show validation results and confirm (N3)
        if (validationErrors.Count > 0)
        {
            _lblValidation.Text    = $"發現 {validationErrors.Count} 筆檢核問題：";
            _lblValidation.Visible = true;
            _dgvErrors.Rows.Clear();
            foreach (var err in validationErrors)
                _dgvErrors.Rows.Add(err.OrderNo, err.Field, err.Message);
            _dgvErrors.Visible = true;

            using var confirmForm = new ValidationConfirmForm(validationErrors.Count);
            var res = confirmForm.ShowDialog(this);
            if (res == DialogResult.Cancel) return;
            if (res == DialogResult.Retry)
            {
                try { Process.Start(new ProcessStartInfo(excelPath) { UseShellExecute = true }); }
                catch { }
                return;
            }
        }
        else
        {
            _dgvErrors.Visible     = false;
            _lblValidation.Visible = false;
        }

        // Step 4: Determine order date (H1)
        string orderDate;
        if (_chkAutoDate.Checked)
        {
            var inferred = TextHelper.InferOrderDate(Path.GetFileName(excelPath), orders);
            if (inferred == null)
            {
                orderDate = DateTime.Today.ToString("yyyy-MM-dd");
                MessageBox.Show("無法從檔名推算日期，已設為今天，請確認。",
                    "日期推算", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                orderDate = inferred;
            }
        }
        else
        {
            orderDate = _dtpOrderDate.Value.ToString("yyyy-MM-dd");
        }

        var hospitalSnap = ReadHospitalFromUI();
        var generalSnap  = ReadGeneralFromUI();
        int maxRowsSnap  = generalSnap.MaxRowsPerPage;
        bool nextMonthInvoiceSnap = _chkNextMonthInvoice.Checked;

        // ---- 分廠商輸出 mode ----
        if (_chkSplitByVendor.Checked)
        {
            string? splitDir = GetOutputDirectory(excelPath, generalSnap, "選擇分廠商 PDF 的輸出目錄");
            if (splitDir == null) return;

            var    ordSnap2  = orders!;
            string dateSnap2 = orderDate;
            string stem2     = Path.GetFileNameWithoutExtension(excelPath);

            SetBusy(true, "正在產生 PDF（分廠商）…");
            List<(string Vendor, string FilePath)>? splitFiles = null;
            string? batchPrintPath = null;
            string? splitErr = null;
            try
            {
                (splitFiles, batchPrintPath) = await Task.Run(() =>
                {
                    var files = PdfGenerator.BuildPdfPerVendor(ordSnap2, splitDir, dateSnap2, stem2, hospitalSnap,
                        maxRowsSnap, nextMonthInvoiceSnap);

                    string batchPath = Path.Combine(splitDir, $"000_{stem2}_整批列印用.pdf");
                    string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
                    try
                    {
                        using (var pdfStream = File.Create(tempPath))
                            PdfGenerator.BuildPdf(ordSnap2, pdfStream, dateSnap2, hospitalSnap,
                                maxRowsSnap, nextMonthInvoiceSnap);
                        File.Move(tempPath, batchPath, overwrite: true);
                    }
                    catch
                    {
                        try { File.Delete(tempPath); } catch { }
                        throw;
                    }

                    return (files, batchPath);
                });
            }
            catch (Exception ex) { splitErr = ex.Message; }
            finally { SetBusy(false); }

            if (splitFiles != null)
            {
                foreach (var (_, fp) in splitFiles)
                    ActivityLogger.Log("產生PDF", source: Path.GetFileName(excelPath),
                        output: Path.GetFileName(fp), success: true);
                if (batchPrintPath != null)
                    ActivityLogger.Log("產生PDF", source: Path.GetFileName(excelPath),
                        output: Path.GetFileName(batchPrintPath), success: true);
            }

            if (splitErr != null)
            {
                ActivityLogger.Log("產生PDF", source: Path.GetFileName(excelPath),
                    success: false, error: splitErr);
                SetStatus("錯誤：" + splitErr, error: true);
                MessageBox.Show(splitErr, "產生失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int sc = splitFiles!.Count;
            string batchName = batchPrintPath != null ? Path.GetFileName(batchPrintPath) : "000_整批列印用.pdf";
            SetStatus($"完成！共 {sc} 家廠商，已產生整批列印檔 {batchName}", success: true);
            if (MessageBox.Show(
                    $"已產生 {sc} 個分廠商 PDF，並另外產生整批列印用 PDF：\n{batchName}\n\n請列印這一個 000_ 開頭檔案，避免 Windows 多選列印亂序。\n\n是否開啟輸出目錄？",
                    "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(splitDir) { UseShellExecute = true });
            return;
        }

        // Step 5: Determine save path (H7)
        string? savePathMaybe = GetSinglePdfSavePath(excelPath, generalSnap);
        if (savePathMaybe == null) return;
        string savePath = savePathMaybe;
        var ordersSnap   = orders;
        var dateSnap     = orderDate;

        // Step 6: Generate PDF via temp file (H3)
        SetBusy(true, "正在產生 PDF…");

        int     rowCount    = 0;
        string  finalDate   = "";
        int     vendorCount = 0;
        bool    pdfOk       = false;
        string? pdfError    = null;

        try
        {
            (rowCount, finalDate, vendorCount) = await Task.Run(() =>
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
                try
                {
                    using (var pdfStream = File.Create(tempPath))
                        PdfGenerator.BuildPdf(ordersSnap, pdfStream, dateSnap, hospitalSnap,
                            maxRowsSnap, nextMonthInvoiceSnap);
                    EnsureOutputDirectory(savePath);
                    File.Move(tempPath, savePath, overwrite: true);
                    int vc = ordersSnap.Select(o => o.Vendor).Distinct().Count();
                    return (ordersSnap.Count, dateSnap, vc);
                }
                catch
                {
                    try { File.Delete(tempPath); } catch { }
                    throw;
                }
            });
            pdfOk = true;
        }
        catch (Exception ex) { pdfError = ex.Message; }
        finally                { SetBusy(false); }

        ActivityLogger.Log("產生PDF",
            source:  Path.GetFileName(excelPath),
            output:  Path.GetFileName(savePath),
            success: pdfOk,
            error:   pdfError);

        if (!pdfOk)
        {
            SetStatus("錯誤：" + pdfError, error: true);
            MessageBox.Show(pdfError, "產生失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        SetStatus($"完成！共 {rowCount} 筆、{vendorCount} 家廠商、訂貨日期 {finalDate}。", success: true);

        // Step 7: Show/reuse PreviewForm (N4)
        try
        {
            if (_previewForm == null || _previewForm.IsDisposed)
            {
                _previewForm = new PreviewForm(savePath);
                _previewForm.Show(this);
            }
            else
            {
                _previewForm.NavigateTo(savePath);
                if (!_previewForm.Visible) _previewForm.Show(this);
                _previewForm.Activate();
            }
        }
        catch (Exception ex)
        {
            if (MessageBox.Show(
                    "PDF 已產生，是否立即開啟？\n(WebView2 預覽不可用：" + ex.Message + ")",
                    "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }

    // ============================================================
    // Tab 2 event handlers
    // ============================================================
    void BtnSaveRules_Click(object? sender, EventArgs e) => SaveCurrentRules();

    void SaveCurrentRules(bool silent = false)
    {
        var rules = new List<ValidationRule>();
        foreach (DataGridViewRow row in _dgvRules.Rows)
        {
            if (row.IsNewRow) continue;
            var enabled = row.Cells["Enabled"].Value is true;
            var field   = row.Cells["Field"].Value?.ToString()    ?? "";
            var typeStr = row.Cells["RuleType"].Value?.ToString()  ?? "必填";
            var msg     = row.Cells["Message"].Value?.ToString()   ?? "";
            if (string.IsNullOrWhiteSpace(field)) continue;
            var ruleType = DisplayToRuleType(typeStr);
            rules.Add(new ValidationRule
            {
                Enabled   = enabled,
                Field     = field,
                RuleType  = ruleType,
                Parameter = "",
                Message   = string.IsNullOrWhiteSpace(msg)
                    ? ValidationService.DefaultMessage(field, ruleType)
                    : msg,
            });
        }
        _validationConfig = new ValidationConfig { Rules = rules };
        AppSettings.SaveValidation(_validationConfig);
        _rulesDirty = false;
        if (!silent)
            MessageBox.Show("檢核設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetRules_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設檢核規則嗎？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _validationConfig = ValidationConfig.Default();
        AppSettings.SaveValidation(_validationConfig);
        _suspendDirtyTracking = true;
        BindRulesGrid();
        _suspendDirtyTracking = false;
        _rulesDirty = false;
    }

    // ============================================================
    // Tab 3 event handlers
    // ============================================================
    void BtnSaveHospital_Click(object? sender, EventArgs e) => SaveCurrentSettings();

    void SaveCurrentSettings(bool silent = false)
    {
        _hospitalSettings                    = ReadHospitalFromUI();
        _generalSettings.PdfOutputMode       = SelectedOutputMode();
        _generalSettings.AutoSaveSameDir     = _generalSettings.PdfOutputMode != PdfOutputMode.AskEveryTime;
        _generalSettings.DefaultPdfDirectory = string.IsNullOrWhiteSpace(_txtDefaultPdfDir.Text)
                                               ? null
                                               : _txtDefaultPdfDir.Text.Trim();
        _generalSettings.MaxRowsPerPage      = (int)_nudMaxRows.Value;
        AppSettings.SaveHospital(_hospitalSettings);
        AppSettings.SaveGeneral(_generalSettings);
        _settingsDirty = false;
        if (!silent)
            MessageBox.Show("設定已儲存。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnResetHospital_Click(object? sender, EventArgs e)
    {
        if (MessageBox.Show("確定要還原為預設值嗎？（一般設定不受影響）", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        _hospitalSettings = HospitalSettings.Default();
        AppSettings.SaveHospital(_hospitalSettings);
        _suspendDirtyTracking = true;
        BindHospitalFields();
        _suspendDirtyTracking = false;
        _settingsDirty = false;
    }

    void BtnBrowsePdfDir_Click(object? sender, EventArgs e)
    {
        using var dlg = new FolderBrowserDialog
        {
            Description            = "選擇 PDF 預設儲存目錄",
            UseDescriptionForTitle = true,
            SelectedPath           = _txtDefaultPdfDir.Text.Trim(),
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            _txtDefaultPdfDir.Text = dlg.SelectedPath;
    }

    // ============================================================
    // FormClosing — dirty flag prompt (N9)
    // ============================================================
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_rulesDirty || _settingsDirty)
        {
            var sb = new StringBuilder("有未儲存的設定變更：\n");
            if (_rulesDirty)    sb.AppendLine("• 檢核設定");
            if (_settingsDirty) sb.AppendLine("• PDF 樣式設定");
            sb.AppendLine("\n是否在關閉前儲存？");

            var res = MessageBox.Show(sb.ToString(), "未儲存的變更",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (res == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
            if (res == DialogResult.Yes)
            {
                if (_rulesDirty)    SaveCurrentRules(silent: true);
                if (_settingsDirty) SaveCurrentSettings(silent: true);
            }
        }

        if (_previewForm != null && !_previewForm.IsDisposed)
            _previewForm.Dispose();

        base.OnFormClosing(e);
    }
}
