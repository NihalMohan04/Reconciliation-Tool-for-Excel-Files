# Reconciliation Tool with OpenPyXL

Python tool to compare two Excel files and generate a reconciled output workbook using `openpyxl`.

## Quick start

1. Create a virtual environment:
   - Windows (PowerShell):
     ```powershell
     py -m venv .venv
     ```
2. Activate it:
   - PowerShell:
     ```powershell
     .\.venv\Scripts\Activate.ps1
     ```
3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```

## Run (GUI)

```powershell
python main.py
```

GUI flow:

- Window title and heading: `Recon Tool by nimo`
- Choose `Source Folder`
- Choose `Target Folder`
- App validates same-named `.xlsx` / `.xlsm` files
- Click `Run Reconciliation`
- Completion dialog shows processed totals and output location
- After completion, the app opens the `Recon` output folder in File Explorer

Output folder behavior:

- If Source and Target share the same parent folder, output is written to `<parent>\Recon`
- Otherwise, output is written to `<project_root>\Recon`

## Run (CLI fallback)

```powershell
python main.py --cli
```

Optional custom project root for CLI mode:

```powershell
python main.py --cli --root C:\path\to\project
```

## Reconciliation behavior

- Automatically reads same-named Excel files from `Source` and `Target` folders.
- Generates separate reconciliation output files in `Recon` folder as `Recon_<filename>.xlsx`.
- Compares first worksheet from each file pair.
- Sorts both datasets by first 3 columns in ascending order before comparing.
- Output first column header is `Match Status` with values `Matched` or `Not Matched`.
- `Match Status` is formula-driven and updates dynamically when ECC/S4 cell values are edited.
- Output columns are interleaved field-by-field: `<Column>_ECC`, `<Column>_S4`.
- Mismatched ECC/S4 value pairs are highlighted dynamically (bold red text) via conditional formatting.
- All `_ECC` columns are filled cyan.
- All `_S4` columns are filled yellow.
- Header row is frozen for easier scrolling.
- Columns are auto-sized based on content.
- Header row is bold.
- `Match Status` column is center-aligned.
