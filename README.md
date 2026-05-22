# Order Helper

Web tool for uploading an Excel order sheet and generating a single PDF purchase order file.

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

## Notes

- Uploaded Excel files and generated PDFs are intentionally ignored by git.
- The PDF layout is generated with fixed coordinates and is expected to be refined against the hospital's reference PDF.
- The app currently uses Windows MingLiU fonts for Traditional Chinese PDF output.
