namespace OrderHelperWinForms.Forms;

public class SheetSelectForm : Form
{
    readonly ListBox _list = new()
    {
        Dock           = DockStyle.Fill,
        SelectionMode  = SelectionMode.One,
        IntegralHeight = false,
    };

    public string? SelectedSheet { get; private set; }

    public SheetSelectForm(IEnumerable<string> sheets)
    {
        Text            = "選擇工作表";
        Size            = new Size(300, 260);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;

        foreach (var s in sheets) _list.Items.Add(s);
        if (_list.Items.Count > 0) _list.SelectedIndex = 0;
        _list.DoubleClick += (_, _) => Accept();

        var ok     = new Button { Text = "確定", Width = 80, Height = 28, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "取消", Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
        ok.Click += (_, _) => Accept();

        var btnPanel = new FlowLayoutPanel
        {
            Dock          = DockStyle.Bottom,
            Height        = 40,
            FlowDirection = FlowDirection.RightToLeft,
            Padding       = new Padding(4),
        };
        btnPanel.Controls.Add(cancel);
        btnPanel.Controls.Add(ok);

        var lblHint = new Label
        {
            Text      = "Excel 含有多個工作表，請選擇要處理的工作表：",
            Dock      = DockStyle.Top,
            Height    = 36,
            Padding   = new Padding(6, 6, 0, 0),
        };

        Controls.Add(_list);
        Controls.Add(btnPanel);
        Controls.Add(lblHint);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    void Accept()
    {
        if (_list.SelectedItem is string s)
        {
            SelectedSheet = s;
            DialogResult  = DialogResult.OK;
            Close();
        }
    }
}
