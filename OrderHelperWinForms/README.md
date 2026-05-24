# OrderHelper WinForms

義大醫院藥品訂購單 PDF 產生器 — Windows 桌面版 (C# / .NET 8 WinForms)

## 功能

- 選擇 Excel 訂購表（.xlsx）
- 自動依廠商分組分頁
- 產出橫式 A4 PDF（格式與原 FastAPI 版完全相同）
- 訂貨日期可自動從檔名/訂購單號推算，或手動選擇
- 輸出單一 .exe（無需安裝 .NET runtime）

## 事前需求

- Windows 10/11 x64
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)（只有 build 時需要）

## Build

```
cd OrderHelperWinForms
dotnet build -c Release
```

## 發布單一 exe

```
cd OrderHelperWinForms
dotnet publish /p:PublishProfile=Properties/PublishProfiles/SingleFile.pubxml
```

或直接指定參數：

```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

輸出位置：`bin\Publish\OrderHelper.exe`

該 exe 完全自包含（含 .NET 8 runtime + 字型），可直接複製給使用者，不需安裝任何東西。

## 字型

使用 Windows 系統字型 **微軟正黑體 Bold**（`C:\Windows\Fonts\msjhbd.ttc`），
無需嵌入任何字型檔，也不需要處理 EUDC 造字。
啟動時按優先順序嘗試 `msjhbd.ttc → msjh.ttc → msjhl.ttc`；若全部找不到才拋出例外。
產出的 PDF 會將用到的字型子集嵌入，可在沒有微軟正黑體的機器上正確顯示。

## 專案結構

```
OrderHelperWinForms/
├── OrderHelperWinForms.csproj
├── Program.cs
├── Models/
│   ├── OrderRow.cs          # 訂購明細資料模型
│   └── VendorPage.cs        # 廠商分頁資料模型
├── Services/
│   ├── ExcelReader.cs       # 讀 Excel，自動偵測欄位
│   ├── PdfGenerator.cs      # 產生 PDF（iText7，座標與原 reportlab 完全相同）
│   └── TextHelper.cs        # 文字正規化、ROC 日期推算
├── Forms/
│   └── MainForm.cs          # WinForms UI
└── Properties/PublishProfiles/
    └── SingleFile.pubxml    # 單檔發布設定
```

## 授權

- iText7 Community: AGPL v3（院內使用無需商業授權）
- ClosedXML: MIT
- NotoSansTC-Bold.ttf: SIL Open Font License
