# PO Tag Labeler

A small local web app that:

1. Accepts an Excel order file and a multi-page PDF of shipping labels.
2. Reads `Cust Po No` and `Tag NO` from Excel.
3. Reads each label's `REF:` number (with `PO:` as a fallback).
4. Adds `TAG NO: ...` to the matching PDF page.
5. Provides the completed PDF for download plus a page-by-page match report.

Duplicate customer PO rows are supported. Their tag numbers are combined on the same label, such as `TAG NO: 75, 76`.

## Run in VS Code on Windows

Open this folder in VS Code, then run these commands in the terminal:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
pip install -r requirements.txt
py -m streamlit run app.py
```

Streamlit will open the app in your browser, normally at `http://localhost:8501`.

For later use, activate the environment and run `start.bat`, or change `start.bat` to point directly to `.venv\Scripts\python.exe`.

## Expected Excel headers

The app recognizes common variations of:

- `Cust Po No`
- `Tag NO`

It searches all worksheets and checks the first 20 rows for the headers.

## Important limitation

The first version works with PDFs that contain selectable text, like the supplied FedEx PDF. A scanned image-only PDF would require OCR support.
