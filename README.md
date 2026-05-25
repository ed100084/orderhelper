# OrderHelper

醫院藥品訂購單 PDF 產生器 — 單一 `.exe`，免安裝，Windows 10/11 x64。

從 Excel 訂購檔批次讀取訂單，依廠商分頁產生 A4 橫向 PDF，可直接列印或傳真。

---

## 功能清單

| 功能 | 說明 |
|------|------|
| 讀取 Excel | 自動識別欄位標題（中英文別名皆可）、合併儲存格解析、支援多工作表 |
| 資料驗證 | 可設定必填、正則格式、最大長度規則；驗證失敗時列出訂購單號與錯誤原因 |
| PDF 產生 | A4 橫向，依廠商分頁，自動分頁分割過長訂單 |
| 日期推算 | 從檔名或訂購單號自動辨識民國日期（YYYMMDD），失敗時提示並設為今天 |
| 拖放操作 | 直接拖放 `.xlsx` 至視窗，無需按選擇按鈕 |
| PDF 預覽 | 內建 WebView2 預覽，可直接列印或另存 |
| 多工作表 | 讀取時彈出工作表選擇清單 |
| 範例 Excel | 一鍵匯出填寫範本 |
| 驗證設定 | 可新增/刪除/停用規則，JSON 格式儲存於 `%APPDATA%\OrderHelper\` |
| 醫院資料 | 標題、發票、交貨地址、備註等均可在 UI 編輯 |
| 一般設定 | 自動儲存至 Excel 同目錄、預設儲存目錄 |
| 操作記錄 | 所有動作寫入每日 `.jsonl`，支援日期過濾與 CSV 匯出 |
| CLI 靜默模式 | 無 GUI，讀 Excel → 驗證 → 產 PDF，exit code 0/1 |
| 版本資訊 | 工具列「關於」對話框顯示版本號與建置日期 |

---

## 快速開始

1. 從 `bin/Publish/` 取得 `OrderHelper.exe`（或自行 publish，見下方）。
2. 雙擊執行，將 `.xlsx` 訂購檔拖放至視窗，或按「選擇檔案…」。
3. 點「產生 PDF」，完成後自動預覽。

---

## CLI 靜默模式

有命令列參數時自動進入 CLI 模式，不開 GUI 視窗。

### 基本用法

```powershell
# 指定輸出路徑
OrderHelper.exe --input orders.xlsx --output orders.pdf

# 指定輸出目錄（自動命名為 orders_訂購單.pdf）
OrderHelper.exe --input orders.xlsx --output-dir D:\output

# 忽略驗證警告強制產出
OrderHelper.exe --input orders.xlsx --output orders.pdf --force

# 使用自訂醫院設定
OrderHelper.exe --input orders.xlsx --output orders.pdf --config custom_hospital.json

# 顯示說明
OrderHelper.exe --help
```

### 選項說明

| 選項 | 縮寫 | 說明 |
|------|------|------|
| `--input <路徑>` | `-i` | Excel 訂購檔路徑（必填） |
| `--output <路徑>` | `-o` | 輸出 PDF 路徑（與 `--output-dir` 擇一） |
| `--output-dir <目錄>` | | 輸出目錄，自動命名 |
| `--config <路徑>` | | 自訂 `hospital_settings.json` 路徑 |
| `--force` | `-f` | 忽略驗證警告繼續產出 |
| `--help` | `-h` | 顯示說明 |

### 退出碼

| 退出碼 | 意義 |
|--------|------|
| `0` | 成功 |
| `1` | 失敗（讀檔失敗、驗證不通過、PDF 產生失敗等） |

### 在 PowerShell 中呼叫

```powershell
$result = & ".\OrderHelper.exe" --input orders.xlsx --output orders.pdf
if ($LASTEXITCODE -ne 0) { Write-Error "PDF 產生失敗" }
```

> **注意：** CLI 模式輸出可能在 PowerShell 提示符之後才顯示，這是 Windows GUI 子系統的限制，不影響功能。

---

## 設定檔格式

設定檔存放於 `%APPDATA%\OrderHelper\`，可直接編輯或透過 GUI 修改。

### `hospital_settings.json`（CLI `--config` 接受相同格式）

```json
{
  "HospitalName": "xx醫院",
  "FormTitle": "藥品訂購單",
  "InvoiceHeader": "xx醫院",
  "InvoiceAddress": "高雄市xxx",
  "TaxId": "12345678",
  "MedicalCode": "1234567890",
  "DrugLicenseNo": "管藥字第XXXX號",
  "DeliveryAddress": "高雄市xxx",
  "DeliveryNote": "收貨時間：週一至週五 08:00-17:00",
  "ContactPhone": "07-615-0011",
  "ContactFax": "07-615-0022",
  "Note1": "請於訂單日期後 3 個工作天內交貨",
  "Note2": "",
  "Note3": "",
  "Note4": ""
}
```

---

## Build / Publish

### 前置需求

- .NET 8 SDK（若 PATH 未包含，可用 `dotnet-install.ps1` 安裝至 `%LOCALAPPDATA%\dotnet-sdk8`）

### 建置（Debug）

```powershell
$env:DOTNET_ROOT="$env:LOCALAPPDATA\dotnet-sdk8"
& "$env:DOTNET_ROOT\dotnet.exe" build OrderHelperWinForms/OrderHelperWinForms.csproj
```

### 發行（單一 .exe，win-x64，self-contained）

```powershell
$env:DOTNET_ROOT="$env:LOCALAPPDATA\dotnet-sdk8"
& "$env:DOTNET_ROOT\dotnet.exe" publish OrderHelperWinForms/OrderHelperWinForms.csproj /p:PublishProfile=SingleFile
```

輸出位置：`OrderHelperWinForms\bin\Publish\OrderHelper.exe`

發行設定檔位於 `OrderHelperWinForms\Properties\PublishProfiles\SingleFile.pubxml`。

---

## 系統需求

| 項目 | 需求 |
|------|------|
| 作業系統 | Windows 10 / 11 x64 |
| 字型 | 微軟正黑體（`C:\Windows\Fonts\msjhbd.ttc`），Windows 內建 |
| PDF 預覽 | Microsoft Edge WebView2 Runtime（Windows 11 內建；Windows 10 請至 [Microsoft](https://go.microsoft.com/fwlink/p/?LinkId=2124703) 下載） |
| 執行環境 | 單一 `.exe`，self-contained，不需安裝 .NET Runtime |

---

## 專案架構

```
OrderHelperWinForms/
├── Forms/
│   ├── MainForm.cs           主視窗（TabControl × 3）
│   ├── AboutForm.cs          關於對話框
│   ├── PreviewForm.cs        WebView2 PDF 預覽
│   ├── LogViewerForm.cs      操作記錄檢視
│   ├── SheetSelectForm.cs    多工作表選擇
│   └── ValidationConfirmForm.cs  驗證警告確認
├── Models/
│   ├── OrderRow.cs           Excel 資料列
│   ├── VendorPage.cs         PDF 廠商分頁
│   ├── ValidationRule.cs     驗證規則 + ValidationError
│   ├── HospitalSettings.cs   醫院設定
│   └── GeneralSettings.cs    一般設定（目錄、儲存選項）
├── Services/
│   ├── ExcelReader.cs        Excel 讀取（ClosedXML）
│   ├── ExcelExporter.cs      範例 Excel 匯出
│   ├── PdfGenerator.cs       PDF 產生（iText7）
│   ├── ValidationService.cs  資料驗證
│   ├── TextHelper.cs         文字正規化、日期推算
│   ├── AppSettings.cs        JSON 設定檔讀寫
│   ├── ActivityLogger.cs     操作記錄（JSONL）
│   └── CliRunner.cs          CLI 靜默模式
└── Program.cs                進入點（CLI / GUI 分流）
```

---

## Changelog

### v1.0.0（2026-05-24）

**初始發行版本**

- 從 Python FastAPI 原型移植為 .NET 8 WinForms 單一 `.exe`
- Excel 讀取：自動識別欄位標題（中英文別名）、合併儲存格解析
- PDF 產生：A4 橫向，iText7，座標完全對應原 reportlab 輸出
- 多廠商分頁、自動換頁、長品名多行排版
- 資料驗證：必填、正則格式、最大長度規則，GUI 可編輯
- 醫院資料全欄可設定（發票抬頭、交貨地址、備註等）
- 一般設定：AutoSaveSameDir、DefaultPdfDirectory
- WebView2 PDF 預覽（含列印、另存）
- 操作記錄（每日 JSONL），日期過濾，CSV 匯出
- 多工作表選擇
- 拖放 `.xlsx` 至視窗
- CLI 靜默模式（`--input / --output / --output-dir / --config / --force`）
- 關於對話框顯示版本號與建置日期
- 驗證確認對話框三按鈕（繼續 / 停止 / 開啟 Excel）
- 髒旗標追蹤（關閉前提示儲存）
