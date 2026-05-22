# Order Helper

Web tool for uploading an Excel order sheet and generating a single PDF purchase order file.

## Current Status

This project is a working prototype that has been converted to a deployable FastAPI app.

Completed:

- GitHub repository initialized and pushed.
- App converted from `http.server` to FastAPI.
- Upload form available at `/`.
- PDF generation endpoint available at `/generate`.
- Health check endpoint available at `/health`.
- CLI generation path preserved for local testing.
- Runtime font strategy changed from Windows `MingLiU` to deployable `Noto Sans TC/CJK`.
- Azure deployment configured for **App Service (Linux) with native Python 3.12 runtime** (no Docker). See `DEPLOY.md`.
- Excel/PDF/DOCX samples and generated files are ignored by git.

Latest known pushed commit:

```text
bdf15d0 convert app to fastapi
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

Target: **Azure App Service (Linux) with native Python 3.12 runtime**. No container.

CI/CD: GitHub Actions zip deploy via OIDC (no long-lived secrets).

See [`DEPLOY.md`](./DEPLOY.md) for full provisioning, deployment, and operations.

Endpoints:

- `/` upload form
- `/generate` PDF generation endpoint
- `/health` health check endpoint

## Validation Performed

Validated locally on Windows:

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
- Noto Sans TC/CJK changes the visual metrics compared with the original PDF font, so x/y positions and column widths may need adjustment.
- Authentication is not implemented in application code. Prefer Azure App Service Easy Auth for external deployment.
- Upload limit defaults to 15 MB and can be changed with `MAX_UPLOAD_BYTES`.
- A custom font path can be supplied with `ORDERHELPER_FONT_PATH`.

## Notes

- Uploaded Excel files and generated PDFs are intentionally ignored by git.
- The PDF layout is generated with fixed coordinates and is expected to be refined against the hospital's reference PDF.
- The app uses Noto Sans TC/CJK for Traditional Chinese PDF output. Set `ORDERHELPER_FONT_PATH` if the runtime needs an explicit font file path.
- Local untracked `prompt.txt` was intentionally not committed.
