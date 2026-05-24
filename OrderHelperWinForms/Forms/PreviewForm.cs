using Microsoft.Web.WebView2.WinForms;

namespace OrderHelperWinForms.Forms;

public class PreviewForm : Form
{
    readonly WebView2 _webView = new();
    readonly string   _pdfPath;

    public PreviewForm(string pdfPath)
    {
        _pdfPath = pdfPath;
        Text          = "PDF 預覽 — " + Path.GetFileName(pdfPath);
        Size          = new Size(1060, 780);
        StartPosition = FormStartPosition.CenterParent;

        var toolbar   = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Color.WhiteSmoke };
        var btnSaveAs = new Button { Text = "另存新檔…", Left = 8,   Top = 7, Width = 100, Height = 28 };
        var btnClose  = new Button { Text = "關閉",     Left = 116, Top = 7, Width = 80,  Height = 28 };

        btnSaveAs.Click += BtnSaveAs_Click;
        btnClose.Click  += (_, _) => Close();
        toolbar.Controls.AddRange(new Control[] { btnSaveAs, btnClose });

        _webView.Dock = DockStyle.Fill;

        // Add Fill first so it occupies the background, then Top toolbar lays over it
        Controls.Add(_webView);
        Controls.Add(toolbar);

        Load += async (_, _) => await InitAsync();
    }

    async Task InitAsync()
    {
        try
        {
            await _webView.EnsureCoreWebView2Async();
            _webView.CoreWebView2.Navigate(new Uri(_pdfPath).AbsoluteUri);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "WebView2 初始化失敗，無法顯示 PDF 預覽。\n\n" +
                "原因：" + ex.Message + "\n\n" +
                "請確認已安裝 Microsoft Edge WebView2 Runtime。\n" +
                "下載：https://go.microsoft.com/fwlink/p/?LinkId=2124703",
                "預覽失敗",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            Close();
        }
    }

    void BtnSaveAs_Click(object? sender, EventArgs e)
    {
        using var dlg = new SaveFileDialog
        {
            Title      = "另存 PDF",
            Filter     = "PDF 文件 (*.pdf)|*.pdf",
            FileName   = Path.GetFileName(_pdfPath),
            DefaultExt = "pdf",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        if (string.Equals(dlg.FileName, _pdfPath, StringComparison.OrdinalIgnoreCase)) return;
        try
        {
            File.Copy(_pdfPath, dlg.FileName, overwrite: true);
            MessageBox.Show("已儲存：" + dlg.FileName, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("另存失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
