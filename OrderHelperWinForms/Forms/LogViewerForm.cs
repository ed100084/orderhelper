using System.Text;
using OrderHelperWinForms.Services;

namespace OrderHelperWinForms.Forms;

public class LogViewerForm : Form
{
    readonly DataGridView _grid    = new();
    readonly Label        _lblInfo = new();
    readonly ComboBox     _cboRange = new();
    List<LogEntry>        _entries  = new();

    public LogViewerForm()
    {
        Text          = "操作記錄";
        Size          = new Size(960, 520);
        StartPosition = FormStartPosition.CenterParent;

        _grid.Dock                           = DockStyle.Fill;
        _grid.ReadOnly                       = true;
        _grid.AllowUserToAddRows             = false;
        _grid.SelectionMode                  = DataGridViewSelectionMode.FullRowSelect;
        _grid.ColumnHeadersHeightSizeMode    = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _grid.RowHeadersVisible              = false;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Timestamp", HeaderText = "時間",     Width = 145 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "User",      HeaderText = "使用者",   Width = 80  });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Action",    HeaderText = "操作",     Width = 130 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Source",    HeaderText = "來源檔案", Width = 200 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Output",    HeaderText = "輸出檔案", Width = 200 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Success",   HeaderText = "結果",     Width = 50  });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Error",     HeaderText = "錯誤訊息",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

        _cboRange.Items.AddRange(new object[] { "最近 7 天", "最近 30 天", "最近 90 天", "全部" });
        _cboRange.SelectedIndex  = 1;  // default: last 30 days
        _cboRange.DropDownStyle  = ComboBoxStyle.DropDownList;
        _cboRange.Width          = 100;
        _cboRange.Left           = 8;
        _cboRange.Top            = 6;
        _cboRange.Height         = 26;
        _cboRange.SelectedIndexChanged += (_, _) => LoadData();

        var toolbar    = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.WhiteSmoke };
        var btnRefresh = new Button { Text = "重新整理", Left = 116, Top = 5, Width = 90, Height = 26 };
        var btnExport  = new Button { Text = "匯出 CSV", Left = 214, Top = 5, Width = 90, Height = 26 };
        _lblInfo.Left  = 314; _lblInfo.Top = 9; _lblInfo.Width = 400; _lblInfo.ForeColor = Color.DimGray;

        btnRefresh.Click += (_, _) => LoadData();
        btnExport.Click  += BtnExport_Click;

        toolbar.Controls.AddRange(new Control[] { _cboRange, btnRefresh, btnExport, _lblInfo });

        Controls.Add(_grid);
        Controls.Add(toolbar);

        LoadData();
    }

    void LoadData()
    {
        DateTime? since = _cboRange.SelectedIndex switch
        {
            0 => DateTime.Today.AddDays(-7),
            1 => DateTime.Today.AddDays(-30),
            2 => DateTime.Today.AddDays(-90),
            _ => (DateTime?)null,
        };

        _entries = ActivityLogger.LoadAll(since);
        _grid.Rows.Clear();
        foreach (var e in _entries)
        {
            int i = _grid.Rows.Add(
                e.Timestamp, e.User, e.Action,
                e.Source ?? "", e.Output ?? "",
                e.Success ? "✓" : "✗",
                e.Error ?? "");
            if (!e.Success)
                _grid.Rows[i].DefaultCellStyle.ForeColor = Color.Crimson;
        }
        _lblInfo.Text = $"共 {_entries.Count} 筆記錄";
    }

    void BtnExport_Click(object? sender, EventArgs e)
    {
        using var dlg = new SaveFileDialog
        {
            Title      = "匯出操作記錄",
            Filter     = "CSV 檔案 (*.csv)|*.csv",
            FileName   = $"ActivityLog_{DateTime.Today:yyyyMMdd}.csv",
            DefaultExt = "csv",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("時間,使用者,操作,來源檔案,輸出檔案,結果,錯誤訊息");
            foreach (var entry in _entries)
            {
                sb.AppendLine(string.Join(",",
                    CsvEscape(entry.Timestamp),
                    CsvEscape(entry.User),
                    CsvEscape(entry.Action),
                    CsvEscape(entry.Source ?? ""),
                    CsvEscape(entry.Output ?? ""),
                    entry.Success ? "成功" : "失敗",
                    CsvEscape(entry.Error ?? "")));
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), System.Text.Encoding.UTF8);
            MessageBox.Show("已匯出：" + dlg.FileName, "完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出失敗：" + ex.Message, "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    static string CsvEscape(string s)
    {
        if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }
}
