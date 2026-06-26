# OrderHelper - WinForms 發佈

本專案目前以 .NET 8 WinForms 桌面程式為主，不再部署到 Azure App Service。

## 目前發佈方式

- 目標平台：Windows 10 / 11 x64
- 輸出型態：單一 self-contained `.exe`
- 專案檔：`OrderHelperWinForms/OrderHelperWinForms.csproj`
- 發佈設定：`OrderHelperWinForms/Properties/PublishProfiles/SingleFile.pubxml`
- 輸出位置：`OrderHelperWinForms/bin/Publish/OrderHelper.exe`

## 本機發佈

需要 .NET 8 SDK。

```powershell
dotnet publish OrderHelperWinForms/OrderHelperWinForms.csproj /p:PublishProfile=SingleFile
```

發佈完成後，把整個 `OrderHelperWinForms/bin/Publish/` 目錄交付給使用者。目錄內包含主程式與 WebView2 相關檔案。

## GitHub Actions

`.github/workflows/main_app-orderhelper.yml` 已改為只建置並發佈 WinForms 桌面版 artifact：

- 使用 `windows-latest`
- 還原 NuGet 套件
- Release build
- 依 `SingleFile.pubxml` 產生 win-x64 self-contained exe
- 上傳 `OrderHelper-win-x64` artifact

此 workflow 不再包含 Azure login、Azure Web App deploy，push 到 `main` 不會部署到 Azure。

## Azure 狀態

舊版 Python FastAPI / Azure App Service 部署已停用於 CI。若 Azure Portal 仍保留 Deployment Center 或 GitHub 連線，建議在 Portal 端也移除或停用，避免平台端設定自行觸發同步。

舊版 Web 入口仍保留於 `app.py`，目前只作為歷史原型或備援參考，不是主要發佈目標。
