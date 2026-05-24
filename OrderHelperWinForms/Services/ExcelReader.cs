using ClosedXML.Excel;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static class ExcelReader
{
    // Column terms to search for in the header row
    static readonly Dictionary<string, string> Terms = new()
    {
        ["order_no"] = "訂購單號",
        ["item_no"]  = "項次",
        ["code"]     = "料號",
        ["name"]     = "品名規格",
        ["unit"]     = "單位",
        ["quantity"] = "訂購量",
        ["vendor"]   = "廠商",
        ["tel"]      = "電話",
        ["fax"]      = "傳真",
    };

    public static List<OrderRow> ReadOrders(Stream stream)
    {
        using var wb = new XLWorkbook(stream);
        foreach (var ws in wb.Worksheets)
        {
            // Search the first 10 rows for the header
            for (int rowNo = 1; rowNo <= 10; rowNo++)
            {
                var headerRow = ws.Row(rowNo);
                // Build map: cell text → column number (1-based)
                var index = new Dictionary<string, int>();
                foreach (var cell in headerRow.CellsUsed())
                {
                    var text = GetCellText(cell.Value);
                    if (!string.IsNullOrEmpty(text))
                        index[text] = cell.Address.ColumnNumber;
                }

                bool hasOrderNo = index.Keys.Any(k => k.Contains(Terms["order_no"]));
                bool hasVendor  = index.Keys.Any(k => k.Contains(Terms["vendor"]));
                if (!hasOrderNo || !hasVendor) continue;

                // Map field names → column numbers
                var cols = new Dictionary<string, int>();
                foreach (var (key, term) in Terms)
                {
                    var match = index.Keys.FirstOrDefault(k => k.Contains(term));
                    if (match is null)
                        throw new InvalidDataException($"Excel 缺少欄位：{term}");
                    cols[key] = index[match];
                }

                // Read data rows
                var orders = new List<OrderRow>();
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? rowNo;
                for (int r = rowNo + 1; r <= lastRow; r++)
                {
                    string orderNo = GetCellText(ws.Cell(r, cols["order_no"]).Value);
                    if (string.IsNullOrEmpty(orderNo)) continue;

                    orders.Add(new OrderRow(
                        OrderNo:  orderNo,
                        ItemNo:   GetCellText(ws.Cell(r, cols["item_no"]).Value),
                        Code:     GetCellText(ws.Cell(r, cols["code"]).Value),
                        Name:     GetCellText(ws.Cell(r, cols["name"]).Value),
                        Unit:     GetCellText(ws.Cell(r, cols["unit"]).Value),
                        Quantity: GetCellText(ws.Cell(r, cols["quantity"]).Value),
                        Vendor:   GetCellText(ws.Cell(r, cols["vendor"]).Value),
                        Tel:      GetCellText(ws.Cell(r, cols["tel"]).Value),
                        Fax:      GetCellText(ws.Cell(r, cols["fax"]).Value)
                    ));
                }
                return orders;
            }
        }
        throw new InvalidDataException("找不到包含「訂購單號」與「廠商」欄位的訂單工作表。");
    }

    static string GetCellText(XLCellValue val)
    {
        if (val.IsBlank) return "";
        if (val.IsNumber)
        {
            double d = val.GetNumber();
            string raw = (d == Math.Floor(d)) ? ((long)d).ToString() : d.ToString();
            return ApplyReplacements(raw.Trim());
        }
        if (val.IsText)  return ApplyReplacements(val.GetText().Trim());
        return ApplyReplacements(val.ToString()?.Trim() ?? "");
    }

    static string ApplyReplacements(string text) => text.Replace("逹", "達");
}
