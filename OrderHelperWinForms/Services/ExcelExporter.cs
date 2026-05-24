using ClosedXML.Excel;

namespace OrderHelperWinForms.Services;

public static class ExcelExporter
{
    static readonly string[] Headers =
    {
        "訂購單號", "項次", "料號", "品名規格", "計價單位", "訂購量", "廠商名稱", "電話", "傳真",
    };

    static readonly object[][] SampleRows =
    {
        new object[] { "PO11401010001", "1", "A001", "阿斯匹靈錠 100mg/tab",   "盒", "10",  "台灣武田藥品股份有限公司", "02-27182000", "02-27182001" },
        new object[] { "PO11401010001", "2", "A002", "普拿疼錠 500mg/tab",     "盒", "20",  "台灣武田藥品股份有限公司", "02-27182000", "02-27182001" },
        new object[] { "PO11401010002", "1", "B001", "福賜多錠 75mg/tab",      "瓶", "5",   "中化製藥股份有限公司",     "04-22388000", "04-22388001" },
        new object[] { "PO11401010003", "1", "C001", "安百嘉注射液 250mg/5mL", "支", "100", "永信藥品工業股份有限公司", "04-26222222", "04-26222220" },
    };

    public static void ExportSample(string outputPath)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("訂購單");

        // Header row
        for (int i = 0; i < Headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = Headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(197, 217, 241);
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Sample data rows
        for (int r = 0; r < SampleRows.Length; r++)
        {
            for (int c = 0; c < SampleRows[r].Length; c++)
            {
                var cell = ws.Cell(r + 2, c + 1);
                cell.Value = SampleRows[r][c].ToString();
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(outputPath);
    }
}
