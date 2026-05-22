# Order Helper

Web tool for uploading an Excel order sheet and generating a single PDF purchase order file.

## Current Status

This project is a working FastAPI prototype deployed to Azure App Service.

Completed:

- Upload form available at `/`.
- PDF generation endpoint available at `/generate`.
- Health check endpoint available at `/health`.
- CLI generation path preserved for local testing.
- Runtime font strategy changed from Windows `MingLiU` to bundled Noto Sans TC fonts.
- Azure deployment configured for App Service on Linux with Python 3.12 and GitHub Actions OIDC.
- Excel/PDF/DOCX samples and generated files are ignored by git.

Current Azure target:

```text
Resource group: orderhelper
Web App:        app-orderhelper
Runtime:        PYTHON|3.12
URL:            https://app-orderhelper.azurewebsites.net
```

Latest known pushed commit before this documentation update:

```text
1e9a915 auto-instance variable font to bold at startup
```

The current workflow:

1. Upload an `.xlsx` file.
2. Read the order sheet in Excel row order.
3. Generate one PDF page per order row.
4. Return one combined PDF.

## Local Run

```powershell
python -m pip install -r requirements.txt
python app.py --port 8000
```

Open:

```text
http://127.0.0.1:8000/
```

## CLI Test

```powershell
python app.py --input .\orders.xlsx --output .\orders.pdf --date 2026-04-30
```

## Azure Deployment

Target: Azure App Service on Linux with native Python 3.12. No Docker image is used.

CI/CD: GitHub Actions zip deploy via OIDC. Azure remote build is enabled with `SCM_DO_BUILD_DURING_DEPLOYMENT=True`, so App Service runs Oryx and installs `requirements.txt` during deployment.

See [`DEPLOY.md`](./DEPLOY.md) for provisioning, deployment, and operations.

Endpoints:

- `/` upload form
- `/generate` PDF generation endpoint
- `/health` health check endpoint

## Font Files

The repository includes Noto Sans TC fonts under `fonts/` so Azure does not need OS-level font installation:

- `fonts/NotoSansTC-Bold.ttf`
- `fonts/NotoSansTC-Regular.otf`
- `fonts/NotoSansTC-Regular.ttf`
- `fonts/NotoSansTC-VariableFont_wght.ttf`

The app tries bundled bold and variable fonts first, then regular fonts. A custom font path can still be supplied with `ORDERHELPER_FONT_PATH`.

## Validation Performed

Validated during development:

- `python -m compileall app.py`
- CLI generation:

```powershell
python app.py --input .\mail訂購--1150430.xlsx --output .\sample_output.pdf --date 2026-04-30
```

- FastAPI server:

```powershell
python app.py --host 127.0.0.1 --port 8000
```

- `/health` returned HTTP 200 with:

```json
{"status":"ok"}
```

- `/generate` accepted the sample Excel upload and returned HTTP 200 with a 96-page PDF.

## Known Follow-Ups

- PDF layout still needs visual fine-tuning against `1031訂單PDF.PDF`.
- Current output intentionally uses fixed PDF coordinates, one Excel row per PDF page.
- Noto Sans TC changes the visual metrics compared with the original PDF font, so x/y positions and column widths may need adjustment.
- Authentication is not implemented in application code. Prefer Azure App Service Easy Auth for external deployment.
- Upload limit defaults to 15 MB and can be changed with `MAX_UPLOAD_BYTES`.

## Notes

- Uploaded Excel files and generated PDFs are intentionally ignored by git.
- The PDF layout is generated with fixed coordinates and is expected to be refined against the hospital's reference PDF.
- Local `prompt.txt` is intentionally ignored.
