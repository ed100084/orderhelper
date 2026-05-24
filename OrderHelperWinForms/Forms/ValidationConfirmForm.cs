namespace OrderHelperWinForms.Forms;

/// <summary>
/// Shown when validation warnings are found.
/// DialogResult.OK     → ignore warnings, continue generating
/// DialogResult.Cancel → stop, let user fix manually
/// DialogResult.Retry  → open Excel file for user to fix
/// </summary>
public class ValidationConfirmForm : Form
{
    public ValidationConfirmForm(int errorCount)
    {
        Text          = "驗證警告";
        Size          = new Size(440, 200);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox   = false;
        MinimizeBox   = false;

        var icon = new Label
        {
            Text      = "⚠",
            Font      = new Font("Segoe UI", 24f, FontStyle.Regular),
            ForeColor = Color.DarkOrange,
            Left      = 20, Top = 30,
            Width     = 50, Height = 55,
            TextAlign = ContentAlignment.MiddleCenter,
        };

        var msg = new Label
        {
            Text      = $"發現 {errorCount} 筆驗證問題。\n\n請選擇處理方式：",
            Left      = 78, Top = 30,
            Width     = 340, Height = 55,
            TextAlign = ContentAlignment.MiddleLeft,
            Font      = new Font("微軟正黑體", 10f),
        };

        var btnContinue = new Button
        {
            Text         = "忽略警告，繼續產生",
            DialogResult = DialogResult.OK,
            Left         = 20, Top = 110, Width = 130, Height = 32,
        };
        var btnOpenExcel = new Button
        {
            Text         = "開啟 Excel 並修正",
            DialogResult = DialogResult.Retry,
            Left         = 160, Top = 110, Width = 130, Height = 32,
        };
        var btnStop = new Button
        {
            Text         = "停止",
            DialogResult = DialogResult.Cancel,
            Left         = 300, Top = 110, Width = 110, Height = 32,
        };

        AcceptButton = btnContinue;
        CancelButton = btnStop;

        Controls.AddRange(new Control[] { icon, msg, btnContinue, btnOpenExcel, btnStop });
    }
}
