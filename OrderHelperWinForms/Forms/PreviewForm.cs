using Microsoft.Web.WebView2.WinForms;

namespace OrderHelperWinForms.Forms;

public class PreviewForm : Form
{
    readonly WebView2 _webView  = new();
    readonly Label    _lblLoading = new();
    string            _pdfPath;

    public PreviewForm(string pdfPath)
    {
        _pdfPath      = pdfPath;
        Text          = "PDF 預覽 — " + Path.GetFileName(pdfPath);
        Size          = new Size(1060, 780);
        StartPosition = FormStartPosition.CenterParent;

        var toolbar   = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Color.WhiteSmoke };
        var btnSaveAs = new Button { Text = "另存新檔…", Left = 8,   Top = 7, Width = 100, Height = 28 };
        var btnPrint  = new Button { Text = "列印",     Left = 116, Top = 7, Width = 80,  Height = 28 };
        var btnClose  = new Button { Text = "關閉",     Left = 204, Top = 7, Width = 80,  Height = 28 };

        btnSaveAs.Click += BtnSaveAs_Click;
        btnPrint.Click  += BtnPrint_Click;
        btnClose.Click  += (_, _) => Hide();
        toolbar.Controls.AddRange(new Control[] { btnSaveAs, btnPrint, btnClose });

        _lblLoading.Text      = "PDF 載入中…";
        _lblLoading.Dock      = DockStyle.Fill;
        _lblLoading.TextAlign = ContentAlignment.MiddleCenter;
        _lblLoading.Font      = new Font("微軟正黑體", 14f);
        _lblLoading.ForeColor = Color.Gray;
        _lblLoading.Visible   = true;

        _webView.Dock    = DockStyle.Fill;
        _webView.Visible = false;

        Controls.Add(_webView);
        Controls.Add(_lblLoading);
        Controls.Add(toolbar);

        Load += async (_, _) => await InitAsync();
    }

    /// <summary>Navigate the already-initialised WebView2 to a new PDF path.</summary>
    public void NavigateTo(string pdfPath)
    {
        _pdfPath = pdfPath;
        Text     = "PDF 預覽 — " + Path.GetFileName(pdfPath);
        if (_webView.CoreWebView2 != null)
            _webView.CoreWebView2.Navigate(new Uri(pdfPath).AbsoluteUri);
    }

    async Task InitAsync()
    {
        try
        {
            await _webView.EnsureCoreWebView2Async();
            _webView.CoreWebView2.Navigate(new Uri(_pdfPath).AbsoluteUri);
            _lblLoading.Visible = false;
            _webView.Visible    = true;
        }
        catch (Exception ex)
        {
            _lblLoading.Visible = false;
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

    async void BtnPrint_Click(object? sender, EventArgs e)
    {
        if (_webView.CoreWebView2 == null) return;
        try
        {
            await _webView.CoreWebView2.ExecuteScriptAsync("window.print()");
        }
        catch (Exception ex)
        {
            MessageBox.Show("列印失敗：" + ex.Message, "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            MessageBox.Show("已儲存：" + dlg.FileName, "完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("另存失敗：" + ex.Message, "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // Hide instead of close so the form can be reused
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Hide();
        }
        else
        {
            base.OnFormClosing(e);
        }
    }
}
