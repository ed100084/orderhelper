# Order Helper - Azure 部署

Azure App Service Linux，Python 3.12。以下指令以 Windows PowerShell 為主。

## 目前 Azure 目標

```text
Resource group: orderhelper
App Service plan: asp-orderhelper-prod
Web App: app-orderhelper
Runtime: PYTHON|3.12
Region: East Asia
URL: https://app-orderhelper.azurewebsites.net
```

目前 App Service 已連到 GitHub Actions：

```text
Repository: https://github.com/ed100084/orderhelper
Branch: main
Workflow: .github/workflows/main_app-orderhelper.yml
```

## 字型檔

App Service Linux 不要依賴 `apt install` 安裝 PDF 字型。專案把 Noto Sans TC 字型放在 `fonts/`，部署時會跟著應用程式一起上傳：

```text
fonts/NotoSansTC-Bold.ttf
fonts/NotoSansTC-Regular.otf
fonts/NotoSansTC-Regular.ttf
fonts/NotoSansTC-VariableFont_wght.ttf
```

程式會先檢查 `ORDERHELPER_FONT_PATH`，再依序使用 repo 內的字型檔，最後才檢查本機 OS fallback 路徑。

## 必要 App Settings

```powershell
az webapp config appsettings set `
  --resource-group orderhelper `
  --name app-orderhelper `
  --settings `
    SCM_DO_BUILD_DURING_DEPLOYMENT=True `
    MAX_UPLOAD_BYTES=15728640
```

`SCM_DO_BUILD_DURING_DEPLOYMENT=True` 很重要。GitHub Actions artifact 會排除本機 `antenv/`，所以 Azure 必須在部署時跑 Oryx，依照 `requirements.txt` 建立 runtime 套件環境。

如果要明確指定字型，可以加：

```powershell
az webapp config appsettings set `
  --resource-group orderhelper `
  --name app-orderhelper `
  --settings ORDERHELPER_FONT_PATH=/home/site/wwwroot/fonts/NotoSansTC-Bold.ttf
```

## 啟動指令

```powershell
az webapp config set `
  --resource-group orderhelper `
  --name app-orderhelper `
  --startup-file "gunicorn -w 2 -k uvicorn.workers.UvicornWorker --timeout 120 -b 0.0.0.0:8000 app:app"
```

App 監聽 `8000`，符合目前 App Service Python image 預設值。

## 部署

日常部署走 GitHub Actions：

```powershell
git push origin main
```

GitHub Actions 還在部署時，不要同時跑 `az webapp up`、Portal Deployment Center 操作、改 App Settings、restart、或手動 zip deploy。這些管理操作可能讓 SCM container 重啟，造成部署只完成一半，最後 runtime 找不到 `uvicorn` 或其他 Python 套件。

如果要手動觸發 Azure 目前 GitHub source connection 重新同步：

```powershell
az webapp deployment source sync `
  --resource-group orderhelper `
  --name app-orderhelper
```

## 驗證

```powershell
Invoke-RestMethod -Uri https://app-orderhelper.azurewebsites.net/health
```

預期回應：

```json
{"status":"ok"}
```

如果網站啟動失敗，log 出現 `ModuleNotFoundError: No module named 'uvicorn'`，代表該次部署沒有完成 Oryx dependency build。先等目前部署完全結束，再重新部署一次。

## 日常維護

| 動作 | 指令 |
|---|---|
| 查看 app config | `az webapp config show -g orderhelper -n app-orderhelper` |
| 查看 app settings | `az webapp config appsettings list -g orderhelper -n app-orderhelper` |
| 看 log | `az webapp log tail -g orderhelper -n app-orderhelper` |
| 重啟 | `az webapp restart -g orderhelper -n app-orderhelper` |
| SSH | `az webapp ssh -g orderhelper -n app-orderhelper` |
| 健康檢查 | `Invoke-RestMethod -Uri https://app-orderhelper.azurewebsites.net/health` |

## 認證

應用程式本身尚未實作登入。對外使用前，建議用 Azure App Service Authentication：

Portal -> Web App -> Authentication -> Add identity provider -> Microsoft -> Single tenant -> Require authentication -> Save.

之後可限制特定使用者或 security group 存取。

## 後續可加

- 加 staging slot，再做 production swap。
- 加 Application Insights 或 Log Analytics。
- 在 GitHub Actions 增加部署後 `/health` 檢查。
