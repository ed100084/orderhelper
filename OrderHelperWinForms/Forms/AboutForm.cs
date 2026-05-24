using System.Reflection;

namespace OrderHelperWinForms.Forms;

public class AboutForm : Form
{
    public AboutForm()
    {
        var asm     = Assembly.GetExecutingAssembly();
        var ver     = asm.GetName().Version;
        string verStr = ver != null ? $"{ver.Major}.{ver.Minor}.{ver.Build}" : "?";

        // Environment.ProcessPath is the actual exe path even in single-file publish
        string exePath  = Environment.ProcessPath
            ?? Path.Combine(AppContext.BaseDirectory, "OrderHelper.exe");
        string buildStr = File.Exists(exePath)
            ? File.GetLastWriteTime(exePath).ToString("yyyy-MM-dd")
            : DateTime.Today.ToString("yyyy-MM-dd");

        Text            = "關於 OrderHelper";
        Size            = new Size(400, 260);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;

        var lblApp = new Label
        {
            Text      = "OrderHelper",
            Font      = new Font("微軟正黑體", 16f, FontStyle.Bold),
            Left      = 20, Top = 20, Width = 360, Height = 36,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        var lblSub = new Label
        {
            Text      = "義大醫院 藥品訂購單 PDF 產生器",
            Font      = new Font("微軟正黑體", 10f),
            Left      = 20, Top = 56, Width = 360, Height = 22,
        };
        var lblVer = new Label
        {
            Text      = $"版本：v{verStr}",
            Left      = 20, Top = 84, Width = 360, Height = 20,
            ForeColor = Color.DimGray,
        };
        var lblBuild = new Label
        {
            Text      = $"建置日期：{buildStr}",
            Left      = 20, Top = 106, Width = 360, Height = 20,
            ForeColor = Color.DimGray,
        };
        var lblDesc = new Label
        {
            Text      = "從 Excel 訂購檔批次產生藥品訂購單 PDF，\n" +
                        "支援多廠商分頁、資料驗證規則與 CLI 靜默模式。",
            Left      = 20, Top = 134, Width = 360, Height = 46,
            ForeColor = Color.DimGray,
        };

        var sep = new Panel { Left = 0, Top = 188, Width = 400, Height = 1, BackColor = Color.LightGray };

        var btnOk = new Button
        {
            Text         = "確定",
            DialogResult = DialogResult.OK,
            Left         = 300, Top = 198, Width = 80, Height = 28,
        };
        AcceptButton = btnOk;

        Controls.AddRange(new Control[] { lblApp, lblSub, lblVer, lblBuild, lblDesc, sep, btnOk });
    }
}
