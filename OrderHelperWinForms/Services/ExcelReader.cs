using ClosedXML.Excel;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static class ExcelReader
{
    // Each canonical field key maps to its ordered list of recognised aliases.
    // Matching is done case-insensitively and ignores spaces / full-width spaces.
    static readonly Dictionary<string, string[]> FieldAliases = new()
    {
        ["order_no"] = new[] { "訂購單號", "訂單號", "採購單號", "po number", "order no", "order number", "po no" },
        ["item_no"]  = new[] { "項次", "序號", "行次", "item no", "item number", "line no", "項目" },
        ["code"]     = new[] { "料號", "藥品代碼", "藥品編號", "drug code", "item code", "品號", "code", "藥碼" },
        ["name"]     = new[] { "品名規格", "品名", "藥品名稱", "drug name", "品名/規格", "name", "規格", "名稱", "product name" },
        ["unit"]     = new[] { "計價單位", "單位", "計算單位", "unit", "uom" },
        ["quantity"] = new[] { "訂購量", "訂購數量", "採購量", "數量", "訂量", "quantity", "qty" },
        ["vendor"]   = new[] { "廠商名稱", "廠商", "供應商名稱", "供應商", "vendor", "supplier" },
        ["tel"]      = new[] { "電話", "廠商電話", "聯絡電話", "tel", "phone", "telephone" },
        ["fax"]      = new[] { "傳真", "廠商傳真", "fax", "facsimile" },
    };

    static readonly string[] RequiredFieldKeys =
    {
        "order_no", "name", "quantity", "vendor",
    };

    // -------------------------------------------------------
    // Public API
    // -------------------------------------------------------

    /// <summary>Returns all worksheet names in the workbook.</summary>
    public static List<string> ListSheets(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return ListSheets(stream);
    }

    public static List<string> ListSheets(Stream stream)
    {
        using var wb = new XLWorkbook(stream);
        return wb.Worksheets.Select(ws => ws.Name).ToList();
    }

    /// <summary>
    /// Reads order rows from <paramref name="stream"/>.
    /// If <paramref name="sheetName"/> is provided, only that sheet is read;
    /// otherwise every sheet is searched for a recognisable header row.
    /// </summary>
    public static List<OrderRow> ReadOrders(Stream stream, string? sheetName = null)
    {
        XLWorkbook wb;
        try
        {
            wb = new XLWorkbook(stream);
        }
        catch (Exception ex)
        {
            throw new InvalidDataException(
                "無法開啟 Excel 檔案，請確認檔案格式正確（.xlsx）。\n原因：" + ex.Message, ex);
        }

        using (wb)
        {
            IXLWorksheet? target = null;
            if (sheetName != null && !wb.TryGetWorksheet(sheetName, out target))
                throw new InvalidDataException($"找不到工作表「{sheetName}」。");

            var sheets = target != null
                ? new List<IXLWorksheet> { target }
                : wb.Worksheets.ToList();

            foreach (var ws in sheets)
            {
                var result = TryReadSheet(ws);
                if (result != null) return result;
            }

            throw new InvalidDataException(
                "找不到包含「訂購單號」與「廠商名稱」欄位的工作表。\n" +
                "請確認 Excel 第一列（或前十列之一）為欄位標題；必要欄位為訂購單號、品名規格、訂購量、廠商名稱。");
        }
    }

    // -------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------

    static List<OrderRow>? TryReadSheet(IXLWorksheet ws)
    {
        // Scan the first 10 rows for a header
        for (int rowNo = 1; rowNo <= 10; rowNo++)
        {
            // Build a cell-text → column-number index for this row.
            // Use all cells, not just CellsUsed(), to catch any that have only whitespace.
            var headerRow = ws.Row(rowNo);
            var index     = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            int lastHeaderCol = ws.Row(rowNo).LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (int col = 1; col <= lastHeaderCol; col++)
            {
                var cell = ws.Cell(rowNo, col);
                var text = GetCellText(ResolveCell(cell));
                if (!string.IsNullOrWhiteSpace(text))
                    index[text] = col;
            }

            if (index.Count == 0) continue;

            // Must recognise both order_no and vendor to treat this as the header
            if (!HasAlias(index, "order_no") || !HasAlias(index, "vendor")) continue;

            // Map recognised fields. Only business-critical columns are hard-required;
            // optional contact/code fields can be validated later by user rules.
            var cols    = new Dictionary<string, int>();
            var missing = new List<string>();
            foreach (var (key, _) in FieldAliases)
            {
                int col = FindColumn(index, key);
                if (col < 0 && RequiredFieldKeys.Contains(key))
                    missing.Add(FieldAliases[key][0]);   // show the primary alias in the error
                else if (col >= 0)
                    cols[key] = col;
            }

            if (missing.Count > 0)
                throw new InvalidDataException(
                    "Excel 缺少以下欄位（或欄位名稱無法辨識）：\n" +
                    string.Join("、", missing));

            // Read data rows
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? rowNo;
            var orders  = new List<OrderRow>();

            for (int r = rowNo + 1; r <= lastRow; r++)
            {
                // Skip fully empty rows
                bool anyValue = false;
                for (int c = 1; c <= lastHeaderCol; c++)
                {
                    if (!ws.Cell(r, c).IsEmpty()) { anyValue = true; break; }
                }
                if (!anyValue) continue;

                string orderNo = GetCellText(ResolveCell(ws.Cell(r, cols["order_no"])));
                if (string.IsNullOrWhiteSpace(orderNo)) continue;

                orders.Add(new OrderRow(
                    OrderNo:  orderNo,
                    ItemNo:   CellTextOrEmpty(ws, r, cols, "item_no"),
                    Code:     CellTextOrEmpty(ws, r, cols, "code"),
                    Name:     CellTextOrEmpty(ws, r, cols, "name"),
                    Unit:     CellTextOrEmpty(ws, r, cols, "unit"),
                    Quantity: CellTextOrEmpty(ws, r, cols, "quantity"),
                    Vendor:   CellTextOrEmpty(ws, r, cols, "vendor"),
                    Tel:      CellTextOrEmpty(ws, r, cols, "tel"),
                    Fax:      CellTextOrEmpty(ws, r, cols, "fax")
                ));
            }

            return orders;
        }

        return null; // no valid header found on this sheet
    }

    static string CellTextOrEmpty(IXLWorksheet ws, int row, Dictionary<string, int> cols, string key)
        => cols.TryGetValue(key, out int col)
            ? GetCellText(ResolveCell(ws.Cell(row, col)))
            : "";

    // If the cell belongs to a merged range, return the top-left cell's value.
    static XLCellValue ResolveCell(IXLCell cell)
    {
        if (cell.IsMerged())
            return cell.MergedRange().FirstCell().Value;
        return cell.Value;
    }

    static bool HasAlias(Dictionary<string, int> index, string fieldKey)
        => FieldAliases[fieldKey].Any(alias =>
               index.Keys.Any(k => Normalize(k).Contains(Normalize(alias))));

    static int FindColumn(Dictionary<string, int> index, string fieldKey)
    {
        foreach (var alias in FieldAliases[fieldKey])
        {
            var match = index.Keys.FirstOrDefault(
                k => Normalize(k).Contains(Normalize(alias)));
            if (match != null) return index[match];
        }
        return -1;
    }

    // Normalise for fuzzy matching: lowercase, strip spaces (half-width and full-width)
    static string Normalize(string s)
        => s.ToLowerInvariant()
            .Replace(" ",  "")
            .Replace("　", ""); // full-width space

    static string GetCellText(XLCellValue val)
    {
        if (val.IsBlank) return "";
        if (val.IsNumber)
        {
            double d = val.GetNumber();
            string raw = (d == Math.Floor(d)) ? ((long)d).ToString() : d.ToString();
            return ApplyReplacements(raw.Trim());
        }
        if (val.IsText) return ApplyReplacements(val.GetText().Trim());
        return ApplyReplacements(val.ToString()?.Trim() ?? "");
    }

    static string ApplyReplacements(string text) => text.Replace("逹", "達");
}
