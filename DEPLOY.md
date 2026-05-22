# Order Helper - Azure 部署

Azure App Service (Linux, Python 3.12)。指令在 Windows PowerShell 跑。

## 1. 準備字型（一次性）

App Service Linux 不能 `apt install`，字型必須隨 repo 一起上傳。

從 https://github.com/notofonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansTC-Regular.otf 下載，放到 `D:\workspace\orderhelper\fonts\NotoSansTC-Regular.otf`。

或 PowerShell：

```powershell
cd D:\workspace\orderhelper
New-Item -ItemType Directory -Force fonts | Out-Null
Invoke-WebRequest `
  -Uri "https://github.com/notofonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansTC-Regular.otf" `
  -OutFile "fonts\NotoSansTC-Regular.otf"
```

## 2. 部署

```powershell
az login
cd D:\workspace\orderhelper

# 第一次：建立資源 + 上傳。之後改 code 重跑同一行就會重新部署。
az webapp up `
  --name app-orderhelper-prod `
  --resource-group rg-orderhelper-prod `
  --plan asp-orderhelper-prod `
  --runtime "PYTHON:3.12" `
  --location eastasia `
  --sku B1
```

> Web App 名稱要全 Azure 唯一，撞名就改 `app-orderhelper-1234`。

設定啟動指令、字型路徑、HTTPS only：

```powershell
az webapp config set `
  --resource-group rg-orderhelper-prod `
  --name app-orderhelper-prod `
  --startup-file "gunicorn -w 2 -k uvicorn.workers.UvicornWorker --timeout 120 -b 0.0.0.0:8000 app:app"

az webapp config appsettings set `
  --resource-group rg-orderhelper-prod `
  --name app-orderhelper-prod `
  --settings `
    ORDERHELPER_FONT_PATH=/home/site/wwwroot/fonts/NotoSansTC-Regular.otf `
    MAX_UPLOAD_BYTES=15728640

az webapp update `
  --resource-group rg-orderhelper-prod `
  --name app-orderhelper-prod `
  --https-only true
```

確認：

```powershell
curl https://app-orderhelper-prod.azurewebsites.net/health
```

## 3. 啟用認證（上線前必做）

Portal → 你的 Web App → **Authentication** → Add identity provider → Microsoft → Single tenant → Require authentication → Save。

之後可在 Portal 限制只允許特定 security group 存取。

## 日常維護

| 動作 | 指令 |
|---|---|
| 改 code 重新部署 | `az webapp up`（同上那行） |
| 看 log | `az webapp log tail -g rg-orderhelper-prod -n app-orderhelper-prod` |
| 重啟 | `az webapp restart -g rg-orderhelper-prod -n app-orderhelper-prod` |
| 上 SSH 進去看 | `az webapp ssh -g rg-orderhelper-prod -n app-orderhelper-prod` |
| 砍掉重練 | `az group delete -n rg-orderhelper-prod --yes` |

## 之後要時再加

- **GitHub Actions 自動部署**：Portal → Deployment Center → GitHub → 一鍵生成 workflow
- **staging slot 藍綠部署**：SKU 升 P1v3 後 `az webapp deployment slot create`
- **Private Endpoint / VNet 整合**：醫院內網才能存取
- **Log Analytics / Application Insights**：合規與監控
