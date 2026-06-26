using System.Runtime.InteropServices;
using System.Text.Json;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

/// <summary>
/// Handles silent (headless) CLI execution when the exe is launched with arguments.
/// Attaches to the parent console so output reaches the calling terminal.
/// </summary>
public static class CliRunner
{
    [DllImport("kernel32.dll")] static extern bool AttachConsole(int pid);
    [DllImport("kernel32.dll")] static extern bool FreeConsole();

    static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Converters    = { new System.Text.Json.Serialization.JsonStringEnumConverter() },
    };

    public static int Run(string[] args)
    {
        // Attach to the parent process's console (cmd / PowerShell).
        // Output may appear after the shell prompt on some terminals — this is cosmetic only.
        AttachConsole(-1);

        try
        {
            return Execute(args);
        }
        finally
        {
            Console.Out.Flush();
            Console.Error.Flush();
            FreeConsole();
        }
    }

    static int Execute(string[] args)
    {
        string? inputPath  = null;
        string? outputPath = null;
        string? outputDir  = null;
        string? configPath = null;
        bool    force      = false;
        bool    help       = false;
        bool    nextMonthInvoice = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--help": case "-h": case "/?":
                    help = true; break;

                case "--input": case "-i":
                    if (i + 1 < args.Length) inputPath = args[++i]; break;

                case "--output": case "-o":
                    if (i + 1 < args.Length) outputPath = args[++i]; break;

                case "--output-dir":
                    if (i + 1 < args.Length) outputDir = args[++i]; break;

                case "--config":
                    if (i + 1 < args.Length) configPath = args[++i]; break;

                case "--force": case "-f":
                    force = true; break;

                case "--next-month-invoice":
                    nextMonthInvoice = true; break;
            }
        }

        if (help)
        {
            PrintHelp();
            return 0;
        }

        if (inputPath == null)
        {
            Console.Error.WriteLine("錯誤：請指定 --input <Excel路徑>。");
            Console.Error.WriteLine("執行 OrderHelper.exe --help 查看說明。");
            return 1;
        }

        if (outputPath == null && outputDir == null)
        {
            Console.Error.WriteLine("錯誤：請指定 --output <PDF路徑> 或 --output-dir <目錄>。");
            return 1;
        }

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"錯誤：找不到輸入檔案：{inputPath}");
            return 1;
        }

        // Load settings (--config overrides HospitalSettings only)
        var validation = AppSettings.LoadValidation();
        var hospital   = LoadHospital(configPath);

        // Read Excel
        List<OrderRow> orders;
        try
        {
            using var stream = File.OpenRead(inputPath);
            orders = ExcelReader.ReadOrders(stream);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"錯誤：無法讀取 Excel 檔案。\n{ex.Message}");
            ActivityLogger.Log("CLI讀取Excel", source: Path.GetFileName(inputPath),
                success: false, error: ex.Message);
            return 1;
        }

        if (orders.Count == 0)
        {
            Console.Error.WriteLine("錯誤：Excel 沒有可輸出的訂單資料。");
            ActivityLogger.Log("CLI讀取Excel", source: Path.GetFileName(inputPath),
                success: false, error: "無資料");
            return 1;
        }

        // Validate
        var errors = ValidationService.Validate(orders, validation);
        if (errors.Count > 0)
        {
            Console.Error.WriteLine($"驗證問題（{errors.Count} 筆）：");
            foreach (var err in errors)
                Console.Error.WriteLine($"  [{err.OrderNo}] {err.Field} — {err.Message}");

            if (!force)
            {
                Console.Error.WriteLine("\n加 --force 可忽略警告繼續產出。");
                ActivityLogger.Log("CLI驗證Excel", source: Path.GetFileName(inputPath),
                    success: false, error: $"{errors.Count} 筆驗證問題");
                return 1;
            }

            Console.Error.WriteLine("（--force 忽略，繼續產出）");
        }

        // Determine save path
        string savePath = outputPath
            ?? Path.Combine(outputDir!, Path.GetFileNameWithoutExtension(inputPath) + "_訂購單.pdf");

        // Infer order date
        string orderDate = TextHelper.InferOrderDate(Path.GetFileName(inputPath), orders)
                           ?? DateTime.Today.ToString("yyyy-MM-dd");

        // Generate PDF via temp file
        try
        {
            string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
            string? saveDir = Path.GetDirectoryName(savePath);
            if (!string.IsNullOrEmpty(saveDir)) Directory.CreateDirectory(saveDir);

            try
            {
                using (var pdfStream = File.Create(tempPath))
                    PdfGenerator.BuildPdf(orders, pdfStream, orderDate, hospital,
                        nextMonthInvoice: nextMonthInvoice);
                File.Move(tempPath, savePath, overwrite: true);
            }
            catch
            {
                try { File.Delete(tempPath); } catch { }
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"錯誤：PDF 產生失敗。\n{ex.Message}");
            ActivityLogger.Log("CLI產生PDF", source: Path.GetFileName(inputPath),
                output: Path.GetFileName(savePath), success: false, error: ex.Message);
            return 1;
        }

        int vendorCount = orders.Select(o => o.Vendor).Distinct().Count();
        Console.WriteLine($"完成：{orders.Count} 筆訂單、{vendorCount} 家廠商、訂貨日期 {orderDate}");
        Console.WriteLine($"輸出：{Path.GetFullPath(savePath)}");

        ActivityLogger.Log("CLI產生PDF", source: Path.GetFileName(inputPath),
            output: Path.GetFileName(savePath), success: true);
        return 0;
    }

    static HospitalSettings LoadHospital(string? configPath)
    {
        if (configPath == null) return AppSettings.LoadHospital();

        try
        {
            var json = File.ReadAllText(configPath);
            return JsonSerializer.Deserialize<HospitalSettings>(json, JsonOpts)
                   ?? AppSettings.LoadHospital();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"警告：無法載入設定檔 {configPath}，使用預設值。\n{ex.Message}");
            return AppSettings.LoadHospital();
        }
    }

    static void PrintHelp()
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        string verStr = ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}" : "";

        Console.WriteLine($"OrderHelper {verStr} — 義大醫院藥品訂購單 PDF 產生器");
        Console.WriteLine();
        Console.WriteLine("用法：");
        Console.WriteLine("  OrderHelper.exe --input <Excel路徑> --output <PDF路徑>");
        Console.WriteLine("  OrderHelper.exe --input <Excel路徑> --output-dir <目錄>");
        Console.WriteLine("  OrderHelper.exe --help");
        Console.WriteLine();
        Console.WriteLine("選項：");
        Console.WriteLine("  --input,      -i   Excel 訂購檔路徑（必填）");
        Console.WriteLine("  --output,     -o   輸出 PDF 路徑（與 --output-dir 擇一）");
        Console.WriteLine("  --output-dir       輸出目錄，自動以 <原檔名>_訂購單.pdf 命名");
        Console.WriteLine("  --config           自訂 hospital_settings.json 路徑");
        Console.WriteLine("  --next-month-invoice  在 PDF 備註區加註紅字：請開立下個月發票");
        Console.WriteLine("  --force,      -f   忽略驗證警告，繼續產出 PDF");
        Console.WriteLine("  --help,       -h   顯示此說明");
        Console.WriteLine();
        Console.WriteLine("範例：");
        Console.WriteLine("  OrderHelper.exe --input orders.xlsx --output orders.pdf");
        Console.WriteLine("  OrderHelper.exe --input orders.xlsx --output-dir D:\\output");
        Console.WriteLine("  OrderHelper.exe --input orders.xlsx --output orders.pdf --force");
        Console.WriteLine("  OrderHelper.exe --input orders.xlsx --output orders.pdf --next-month-invoice");
        Console.WriteLine("  OrderHelper.exe --input orders.xlsx --output orders.pdf --config custom_hs.json");
        Console.WriteLine();
        Console.WriteLine("退出碼：  0 = 成功   1 = 失敗");
        Console.WriteLine();
        Console.WriteLine("設定檔預設路徑（GUI 模式儲存）：");
        Console.WriteLine($"  %APPDATA%\\OrderHelper\\");
    }
}
