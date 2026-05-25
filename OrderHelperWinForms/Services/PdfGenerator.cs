using System.Text;
using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

/// <summary>
/// Generates the pharmacy order PDF.
/// Coordinates are in PDF points matching the original reportlab layout exactly
/// (iText7 shares the same bottom-left origin, Y-up coordinate system).
/// </summary>
public static class PdfGenerator
{
    // --- Layout constants (copied directly from app.py) ---
    const float PAGE_W           = 841.89f;
    const float PAGE_H           = 595.28f;
    const float DETAIL_TOP       = 410.56f;
    const float DETAIL_HEADER_H  = 19.52f;
    const float DETAIL_LINE_H    = 10.2f;
    const float DETAIL_ROW_PAD   = 4.0f;
    const float DETAIL_MIN_ROW_H = 24.0f;

    // --- Windows system font: 微軟正黑體 Bold ---
    // Try Bold variant first; fall back to Regular if Bold isn't installed.
    static readonly string[] FontCandidates =
    {
        @"C:\Windows\Fonts\msjhbd.ttc,0",   // Microsoft JhengHei Bold
        @"C:\Windows\Fonts\msjh.ttc,0",     // Microsoft JhengHei Regular
        @"C:\Windows\Fonts\msjhl.ttc,0",    // Microsoft JhengHei Light (last resort)
    };

    static string ResolveFontPath()
    {
        foreach (var entry in FontCandidates)
        {
            var path = entry.Split(',')[0];
            if (File.Exists(path)) return entry;
        }
        throw new FileNotFoundException(
            "找不到微軟正黑體字型（msjhbd.ttc）。\n" +
            "請確認 C:\\Windows\\Fonts\\msjhbd.ttc 存在。");
    }

    static readonly Lazy<string> FontEntry = new(ResolveFontPath);

    // Shared font for text-width measurement (no PdfDocument needed, metrics only).
    static readonly Lazy<PdfFont> MeasureFont = new(() =>
        PdfFontFactory.CreateFont(FontEntry.Value, PdfEncodings.IDENTITY_H));

    // -------------------------------------------------------
    // Public entry points
    // -------------------------------------------------------

    /// <summary>Reads Excel from stream, then generates the PDF.</summary>
    public static (int rowCount, string orderDate, int vendorCount) BuildPdf(
        Stream excelSource,
        Stream pdfOutput,
        string filename = "orders.xlsx",
        string? orderDate = null,
        HospitalSettings? hospital = null,
        string? sheetName = null)
    {
        var orders = ExcelReader.ReadOrders(excelSource, sheetName);
        if (orders.Count == 0)
            throw new InvalidOperationException("Excel 沒有可輸出的訂單資料。");

        string finalDate = orderDate
            ?? TextHelper.InferOrderDate(filename, orders)
            ?? DateTime.Today.ToString("yyyy-MM-dd");
        var hs = hospital ?? HospitalSettings.Default();
        BuildPdfFromOrders(orders, pdfOutput, finalDate, hs);

        int vendorCount = orders.Select(o => o.Vendor).Distinct().Count();
        return (orders.Count, finalDate, vendorCount);
    }

    /// <summary>Generates PDF from a pre-read list of orders.</summary>
    public static (int rowCount, string orderDate, int vendorCount) BuildPdf(
        List<OrderRow> orders,
        Stream pdfOutput,
        string orderDate,
        HospitalSettings? hospital = null)
    {
        if (orders.Count == 0)
            throw new InvalidOperationException("Excel 沒有可輸出的訂單資料。");

        var hs = hospital ?? HospitalSettings.Default();
        BuildPdfFromOrders(orders, pdfOutput, orderDate, hs);

        int vendorCount = orders.Select(o => o.Vendor).Distinct().Count();
        return (orders.Count, orderDate, vendorCount);
    }

    // -------------------------------------------------------
    // Core PDF builder
    // -------------------------------------------------------
    static void BuildPdfFromOrders(List<OrderRow> orders, Stream output, string orderDate,
                                   HospitalSettings hs)
    {
        using var writer  = new PdfWriter(output);
        using var pdfDoc  = new PdfDocument(writer);
        var font = PdfFontFactory.CreateFont(FontEntry.Value, PdfEncodings.IDENTITY_H, pdfDoc);

        var vendorPages = BuildVendorPages(orders);
        // Track cumulative row count per vendor so sequence numbers continue across pages
        var seqByVendor = new Dictionary<string, int>();
        foreach (var vp in vendorPages)
        {
            int seqStart = seqByVendor.GetValueOrDefault(vp.Vendor, 0) + 1;
            seqByVendor[vp.Vendor] = seqStart - 1 + vp.Orders.Count;

            var page   = pdfDoc.AddNewPage(new PageSize(PAGE_W, PAGE_H));
            var canvas = new PdfCanvas(page);
            canvas.SetStrokeColor(ColorConstants.BLACK)
                  .SetFillColor(ColorConstants.BLACK);

            var rowHeights = DrawRects(canvas, vp.Orders, font);
            DrawStatic(canvas, font, vp.PageNo, vp.TotalPages, hs);
            DrawVendorPage(canvas, font, vp, orderDate, rowHeights, seqStart);
        }
    }

    // -------------------------------------------------------
    // Vendor grouping & pagination  (mirrors app.py logic)
    // -------------------------------------------------------
    static List<VendorPage> BuildVendorPages(List<OrderRow> orders)
    {
        // Group by vendor, preserving insertion order
        var grouped = new Dictionary<string, List<OrderRow>>();
        var keys    = new List<string>();
        foreach (var o in orders)
        {
            if (!grouped.ContainsKey(o.Vendor))
            {
                grouped[o.Vendor] = new List<OrderRow>();
                keys.Add(o.Vendor);
            }
            grouped[o.Vendor].Add(o);
        }

        var pages = new List<VendorPage>();
        foreach (var key in keys)
        {
            var vendorOrders = grouped[key];
            var first  = vendorOrders[0];
            var chunks = PaginateVendorOrders(vendorOrders);
            for (int i = 0; i < chunks.Count; i++)
            {
                pages.Add(new VendorPage(
                    Vendor:     key,
                    Tel:        first.Tel,
                    Fax:        first.Fax,
                    PageNo:     i + 1,
                    TotalPages: chunks.Count,
                    Orders:     chunks[i]
                ));
            }
        }
        return pages;
    }

    const int MaxRowsPerPage = 10;

    static List<List<OrderRow>> PaginateVendorOrders(List<OrderRow> orders)
    {
        float capacity = DETAIL_TOP - 58.0f; // usable body height (DETAIL_BOTTOM = 58 pt)
        var pages    = new List<List<OrderRow>>();
        var current  = new List<OrderRow>();
        float used   = 0f;

        foreach (var order in orders)
        {
            float rh = DetailRowHeight(order);
            bool countFull  = current.Count >= MaxRowsPerPage;
            bool heightFull = used + rh > capacity;
            if (current.Count > 0 && (countFull || heightFull))
            {
                pages.Add(current);
                current = new List<OrderRow>();
                used    = 0f;
            }
            current.Add(order);
            used += rh;
        }
        if (current.Count > 0) pages.Add(current);
        return pages;
    }

    // -------------------------------------------------------
    // Row height calculation
    // -------------------------------------------------------
    static float DetailRowHeight(OrderRow order)
    {
        int lineCount = Math.Max(1, FitDetailName(order.Name).Count);
        return Math.Max(DETAIL_MIN_ROW_H, DETAIL_ROW_PAD * 2 + lineCount * DETAIL_LINE_H);
    }

    static List<string> FitDetailName(string text)
        => FitText(text, 345f, MeasureFont.Value, 9f);

    static List<string> FitText(string text, float maxWidth, PdfFont font, float fontSize, int maxLines = 5)
    {
        var lines   = new List<string>();
        var current = new StringBuilder();

        foreach (char ch in text)
        {
            current.Append(ch);
            string trial = current.ToString();
            if (current.Length > 1 && font.GetWidth(trial, fontSize) > maxWidth)
            {
                // Back off: remove the character that pushed it over the limit
                current.Length--;
                lines.Add(current.ToString());
                if (lines.Count >= maxLines) return lines;
                current.Clear();
                current.Append(ch);
            }
        }
        if (current.Length > 0 && lines.Count < maxLines)
            lines.Add(current.ToString());

        return lines.Count > 0 ? lines : new List<string> { "" };
    }

    // -------------------------------------------------------
    // Rectangle / line drawing  (_draw_rects in app.py)
    // -------------------------------------------------------
    static List<float> DrawRects(PdfCanvas canvas, List<OrderRow> orders, PdfFont font)
    {
        canvas.SetLineWidth(0.4f);

        var rowHeights   = orders.Select(o => DetailRowHeightWithFont(o, font)).ToList();
        float bodyH      = Math.Max(45.44f, rowHeights.Sum());
        float detailBot  = DETAIL_TOP - bodyH;

        // Outer rectangles (x, y_bottom, w, h) — identical to reportlab values
        (float x, float y, float w, float h)[] rects =
        {
            (21.76f, 504.24f, 801.0f, 24.0f),
            (21.76f, 430.04f, 801.0f, 74.23f),
            (21.76f, 410.56f, 801.0f, 19.52f),
            (63.0f,  410.52f, 105.0f, 19.52f),
            (240.76f,410.52f,  75.76f,19.52f),
            (678.76f,410.52f,  75.76f,19.52f),
            (270.0f, 430.0f,  293.24f,74.23f),
            (21.76f, detailBot, 801.0f, bodyH),
        };
        foreach (var (x, y, w, h) in rects)
            canvas.Rectangle(x, y, w, h);

        // Horizontal row separators
        float yPos = DETAIL_TOP;
        foreach (float rh in rowHeights.SkipLast(1))
        {
            yPos -= rh;
            canvas.MoveTo(21.76f, yPos).LineTo(822.76f, yPos);
        }

        // Vertical column separators
        float[] colX = { 63.04f, 168.04f, 240.68f, 316.48f, 678.76f, 754.48f };
        foreach (float cx in colX)
            canvas.MoveTo(cx, DETAIL_TOP).LineTo(cx, detailBot);

        canvas.Stroke();
        return rowHeights;
    }

    // Row height using the document's font (for the final drawing pass, same result)
    static float DetailRowHeightWithFont(OrderRow order, PdfFont font)
    {
        int lineCount = Math.Max(1, FitText(order.Name, 345f, font, 9f).Count);
        return Math.Max(DETAIL_MIN_ROW_H, DETAIL_ROW_PAD * 2 + lineCount * DETAIL_LINE_H);
    }

    // -------------------------------------------------------
    // Static (fixed) text/graphics per page  (_draw_static)
    // -------------------------------------------------------
    static void DrawStatic(PdfCanvas canvas, PdfFont font, int pageNo, int totalPages,
                           HospitalSettings hs)
    {
        // Title
        DrawCenterBold(canvas, font, 20f, 421f, 566.4f, hs.HospitalName);
        DrawCenterBold(canvas, font, 18f, 421f, 540.1f, hs.FormTitle);

        // Page number
        DrawStr(canvas, font, 14f, 757.6f, 538.2f, pageNo.ToString());
        DrawStr(canvas, font, 14f, 778.0f, 536.7f, "/");
        DrawStr(canvas, font, 14f, 800.8f, 538.2f, totalPages.ToString());

        DrawStr(canvas, font, 10f, 28.8f, 535.8f, "報表代碼：INV_APP_07");

        // Vendor header row labels
        DrawStr(canvas, font, 12f,  30.8f, 512.7f, "廠商名稱");
        DrawStr(canvas, font, 12f, 389.3f, 512.7f, "FAX：");
        // □ mail   訂貨日期  □  [date]  — two separate checkboxes, matching original layout
        canvas.SetLineWidth(0.4f)
              .Rectangle(605.8f, 511.2f, 7.0f, 7.0f)   // □ before "mail"
              .Stroke();
        DrawStr(canvas, font, 12f, 616.0f, 512.7f, "mail");
        DrawStr(canvas, font, 12f, 683.0f, 512.7f, "訂貨日期");

        // Invoice block
        DrawStr(canvas, font, 10f, 30.8f, 488.6f, "發票抬頭：" + hs.InvoiceHeader);
        DrawStr(canvas, font, 10f, 30.8f, 476.6f, "發票地址：" + hs.InvoiceAddress);
        DrawStr(canvas, font, 10f, 30.8f, 464.6f, "統一編號：" + hs.TaxId);
        DrawStr(canvas, font, 10f, 30.8f, 452.6f, "醫療機構代碼：" + hs.MedicalCode);
        DrawStr(canvas, font, 10f, 30.8f, 440.6f, "**管證字號：" + hs.DrugLicenseNo);

        // Delivery block
        DrawStr(canvas, font, 10f, 279.7f, 485.6f, "交貨地址：" + hs.DeliveryAddress);
        DrawStr(canvas, font, 10f, 279.7f, 473.6f, hs.DeliveryNote);
        DrawStr(canvas, font, 10f, 279.7f, 461.6f, "聯絡電話：" + hs.ContactPhone);
        DrawStr(canvas, font, 10f, 279.7f, 449.6f, "傳真：" + hs.ContactFax);

        // Notes — use DrawFitStr so long text shrinks instead of overflowing off-page
        const float NoteX    = 578.9f;
        const float NoteMaxW = 822.76f - NoteX - 4f;   // ≈ 240 pt
        DrawStr    (canvas, font, 9f, NoteX, 490.3f, "※備註：");
        DrawFitStr (canvas, font, NoteX, 479.3f, hs.Note1, NoteMaxW, 9f, 6f);
        DrawFitStr (canvas, font, NoteX, 468.2f, hs.Note2, NoteMaxW, 9f, 6f);
        DrawFitStr (canvas, font, NoteX, 457.2f, hs.Note3, NoteMaxW, 9f, 6f);
        DrawFitStr (canvas, font, NoteX, 446.2f, hs.Note4, NoteMaxW, 9f, 6f);

        // Column headers
        DrawStr(canvas, font, 12f,  30.4f, 417.1f, "序號");
        DrawStr(canvas, font, 12f,  72.0f, 417.1f, "訂購編號");
        DrawStr(canvas, font, 12f, 181.5f, 417.1f, "料號項次");
        DrawStr(canvas, font, 12f, 256.5f, 417.1f, "料號");
        DrawStr(canvas, font, 12f, 432.0f, 417.1f, "品名規格");
        DrawStr(canvas, font, 12f, 691.0f, 417.1f, "計價單位");
        DrawStr(canvas, font, 12f, 766.0f, 417.1f, "訂購量");
    }

    // -------------------------------------------------------
    // Per-vendor dynamic content  (_draw_vendor_page)
    // -------------------------------------------------------
    /// <param name="seqStart">
    /// 1-based sequence number of the first row on this page.
    /// Pass 1 for the first page of a vendor; subsequent pages receive the
    /// cumulative count so numbers continue across page breaks.
    /// </param>
    static void DrawVendorPage(PdfCanvas canvas, PdfFont font, VendorPage page,
                               string orderDate, List<float> rowHeights, int seqStart = 1)
    {
        DrawFitStr(canvas, font,  88.6f, 513.4f, page.Vendor,           98.0f, 12f);
        DrawFitStr(canvas, font, 191.0f, 512.7f, $"TEL：{page.Tel}",   190.0f, 12f);
        DrawFitStr(canvas, font, 423.0f, 512.7f, page.Fax,             136.0f, 12f);
        // □ before the date (second checkbox)
        canvas.SetLineWidth(0.4f)
              .Rectangle(737.0f, 511.2f, 7.0f, 7.0f)
              .Stroke();
        DrawStr(canvas, font, 12f, 747.6f, 511.9f, orderDate);

        float yTop = DETAIL_TOP;
        for (int idx = 0; idx < page.Orders.Count; idx++)
        {
            var order  = page.Orders[idx];
            float rh   = rowHeights[idx];
            float y    = yTop - 14.0f;   // baseline inside the row

            DrawCenter(canvas, font, 9f,  42.4f, y, (seqStart + idx).ToString());
            DrawStr   (canvas, font, 9f,  68.0f, y, order.OrderNo);
            DrawStr   (canvas, font, 9f, 181.5f, y, order.ItemNo);
            DrawStr   (canvas, font, 9f, 259.5f, y, order.Code);
            DrawStr   (canvas, font, 9f, 691.0f, y, order.Unit);
            DrawRight (canvas, font, 9f, 812.0f, y, order.Quantity);

            var lines = FitText(order.Name, 345f, font, 9f);
            for (int li = 0; li < lines.Count; li++)
                DrawStr(canvas, font, 9f, 321.4f, y - li * DETAIL_LINE_H, lines[li]);

            yTop -= rh;
        }
    }

    // -------------------------------------------------------
    // Low-level draw helpers
    // -------------------------------------------------------

    // DrawCenterBold: simulate the double-draw bold trick from reportlab.
    // Since we're using the Bold font, one draw is sufficient visually,
    // but we replicate the tiny offset for pixel-accurate matching.
    static void DrawCenterBold(PdfCanvas canvas, PdfFont font, float size,
                               float x, float y, string text)
    {
        DrawCenter(canvas, font, size, x,        y, text);
        DrawCenter(canvas, font, size, x + 0.35f, y, text);
    }

    static void DrawStr(PdfCanvas canvas, PdfFont font, float size,
                        float x, float y, string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        canvas.BeginText()
              .SetFontAndSize(font, size)
              .SetTextMatrix(1f, 0f, 0f, 1f, x, y)
              .ShowText(text)
              .EndText();
    }

    static void DrawRight(PdfCanvas canvas, PdfFont font, float size,
                          float x, float y, string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        float w = font.GetWidth(text, size);
        DrawStr(canvas, font, size, x - w, y, text);
    }

    static void DrawCenter(PdfCanvas canvas, PdfFont font, float size,
                           float x, float y, string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        float w = font.GetWidth(text, size);
        DrawStr(canvas, font, size, x - w / 2f, y, text);
    }

    static void DrawFitStr(PdfCanvas canvas, PdfFont font,
                           float x, float y, string text,
                           float maxWidth, float fontSize, float minFontSize = 7f)
    {
        if (string.IsNullOrEmpty(text)) return;
        float size = fontSize;
        while (size > minFontSize && font.GetWidth(text, size) > maxWidth)
            size -= 0.5f;
        DrawStr(canvas, font, size, x, y, text);
    }

    // -------------------------------------------------------
    // Per-vendor split output
    // -------------------------------------------------------

    /// <summary>
    /// Generates one PDF file per vendor and writes them to <paramref name="outputDir"/>.
    /// The file name pattern is  {inputStem}_{vendorName}_訂購單.pdf.
    /// Returns a list of (vendorName, filePath) in vendor order.
    /// </summary>
    public static List<(string Vendor, string FilePath)> BuildPdfPerVendor(
        List<OrderRow> orders,
        string outputDir,
        string orderDate,
        string inputStem,
        HospitalSettings? hospital = null)
    {
        if (orders.Count == 0)
            throw new InvalidOperationException("Excel 沒有可輸出的訂單資料。");

        var hs       = hospital ?? HospitalSettings.Default();
        var allPages = BuildVendorPages(orders);
        var result   = new List<(string, string)>();

        // Distinct vendor names, preserving order
        var vendorNames = allPages.Select(vp => vp.Vendor).Distinct().ToList();

        Directory.CreateDirectory(outputDir);

        foreach (var vendor in vendorNames)
        {
            var vendorPages = allPages.Where(vp => vp.Vendor == vendor).ToList();
            string safeName = SanitizeFileName(vendor);
            string fileName = $"{inputStem}_{safeName}_訂購單.pdf";
            string filePath = System.IO.Path.Combine(outputDir, fileName);

            using var writer = new PdfWriter(filePath);
            using var pdfDoc = new PdfDocument(writer);
            var font = PdfFontFactory.CreateFont(FontEntry.Value, PdfEncodings.IDENTITY_H, pdfDoc);

            int seqStart = 1;
            foreach (var vp in vendorPages)
            {
                var page   = pdfDoc.AddNewPage(new PageSize(PAGE_W, PAGE_H));
                var canvas = new PdfCanvas(page);
                canvas.SetStrokeColor(ColorConstants.BLACK)
                      .SetFillColor(ColorConstants.BLACK);
                var rowHeights = DrawRects(canvas, vp.Orders, font);
                DrawStatic(canvas, font, vp.PageNo, vp.TotalPages, hs);
                DrawVendorPage(canvas, font, vp, orderDate, rowHeights, seqStart);
                seqStart += vp.Orders.Count;
            }

            result.Add((vendor, filePath));
        }

        return result;
    }

    static string SanitizeFileName(string name)
    {
        var invalid = new HashSet<char>(System.IO.Path.GetInvalidFileNameChars());
        return string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c));
    }
}
