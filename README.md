# Figure Case Study — All Code

## Files

### JavaScript (Node.js — saved as files)
- **create_deck.js** — Builds the 9-slide PowerPoint using pptxgenjs. Embeds chart PNGs, creates all slide layouts, text, tables, and speaker notes.
- **speaker_notes.js** — Generates the Speaker_Notes.docx companion document using docx library.
- **workbook_doc.js** — Creates the Excel analysis spreadsheet (6 sheets) using openpyxl via Python subprocess + formatting.

### Python (run inline — reconstructed below)
The Python scripts were executed as inline bash heredocs rather than saved files. The two main scripts are included below as .py files:

- **analysis_and_charts.py** — Pandas analysis + matplotlib chart generation (5 charts)
- **excel_workbook.py** — openpyxl workbook creation with formulas and formatting (2 passes: base + extended sheets)

## Dependencies
```bash
npm install -g pptxgenjs
pip install pandas openpyxl matplotlib --break-system-packages
```

## Run Order
1. `python analysis_and_charts.py` — generates chart1-5 PNGs
2. `node create_deck.js` — builds .pptx (needs charts from step 1)
3. `node speaker_notes.js` — builds .docx
4. `python excel_workbook.py` — builds .xlsx
5. `python /mnt/skills/public/xlsx/scripts/recalc.py analysis.xlsx` — recalculates Excel formulas
