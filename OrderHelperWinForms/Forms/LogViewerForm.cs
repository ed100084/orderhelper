using OrderHelperWinForms.Services;

namespace OrderHelperWinForms.Forms;

public class LogViewerForm : Form
{
    readonly DataGridView _grid    = new();
    readonly Label        _lblInfo = new();

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

        var toolbar    = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.WhiteSmoke };
        var btnRefresh = new Button { Text = "重新整理", Left = 8, Top = 5, Width = 90, Height = 26 };
        _lblInfo.Left  = 108; _lblInfo.Top = 9; _lblInfo.Width = 400; _lblInfo.ForeColor = Color.DimGray;
        btnRefresh.Click += (_, _) => LoadData();
        toolbar.Controls.AddRange(new Control[] { btnRefresh, _lblInfo });

        Controls.Add(_grid);
        Controls.Add(toolbar);

        LoadData();
    }

    void LoadData()
    {
        _grid.Rows.Clear();
        var entries = ActivityLogger.LoadAll();
        foreach (var e in entries)
        {
            int i = _grid.Rows.Add(
                e.Timestamp, e.User, e.Action,
                e.Source ?? "", e.Output ?? "",
                e.Success ? "✓" : "✗",
                e.Error ?? "");
            if (!e.Success)
                _grid.Rows[i].DefaultCellStyle.ForeColor = Color.Crimson;
        }
        _lblInfo.Text = $"共 {entries.Count} 筆記錄";
    }
}
