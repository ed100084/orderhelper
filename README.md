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
3. Group rows by vendor name while preserving the first vendor appearance order.
4. Generate one or more PDF pages per vendor, depending on how many detail rows fit.
5. Return one combined PDF.

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

- `/generate` accepted the sample Excel upload and returned HTTP 200.
- CLI generation from `mail訂購--1150430.xlsx` produced 9 vendors, 96 rows, and a 13-page grouped PDF.

## Known Follow-Ups

- PDF layout still needs visual fine-tuning against `1031訂單PDF.PDF`.
- Current output intentionally uses fixed PDF coordinates, grouping multiple Excel rows into vendor order pages.
- Noto Sans TC changes the visual metrics compared with the original PDF font, so x/y positions and column widths may need adjustment.
- Authentication is not implemented in application code. Prefer Azure App Service Easy Auth for external deployment.
- Upload limit defaults to 15 MB and can be changed with `MAX_UPLOAD_BYTES`.

## Roadmap

Recommended next improvements, in priority order:

1. Deployment stability

   - Add GitHub Actions `concurrency` so overlapping deployments do not interrupt each other.
   - Add a post-deployment `/health` check to the workflow.
   - Pin dependency versions in `requirements.txt` so Azure builds are repeatable.
   - Update the workflow if GitHub's Node.js 20 actions deprecation warning starts requiring changes.

2. Error handling

   - Return clear 400 responses for user input problems such as missing Excel columns, missing order sheets, empty order data, or invalid files.
   - Add a deeper readiness endpoint that checks bundled fonts and basic PDF generation.
   - Improve the upload page so users see friendly error messages instead of raw JSON errors.

3. PDF quality

   - Compare output against `1031訂單PDF.PDF` and tune fixed coordinates, spacing, and column widths.
   - Add a sample Excel regression test that verifies generated PDF page count and non-empty output.
   - Handle long `品名規格` values more deliberately instead of silently limiting to five lines.

4. Operations and access control

   - Enable Azure App Service Easy Auth before external use.
   - Add limits for Excel row count, generated page count, and processing time to protect the App Service instance.
   - Include app version or commit SHA in `/health` so the deployed version is easy to confirm.

## Notes

- Uploaded Excel files and generated PDFs are intentionally ignored by git.
- The PDF layout is generated with fixed coordinates and is expected to be refined against the hospital's reference PDF.
- Local `prompt.txt` is intentionally ignored.
