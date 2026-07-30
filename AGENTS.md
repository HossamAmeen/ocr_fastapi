# AGENTS.md — Living Guide for AI Coding Agents

This file documents the OCR/Excel automation project so future agent sessions
have accurate context. Keep it up to date after every meaningful change.

## Project overview

FastAPI app that extracts procedure/text data from PDFs (via OCR/text
extraction) and writes it into pre-built Excel templates (`.xlsm`/`.xlsx`).
There are three independent flows, each with its own extractor + writer +
router/service/schema/UI page:

- **Proforma** — `extract_performa.py`, `excel_proforma/`
- **SOE (time log)** — `extract_soe.py`, `excel_soe/`
- **Job Order** — `extract_job_order.py`, `excel_job_offer/`

Shared PDF text extraction lives in `pdf_extractor/`.

Web layer (`app/`):
- `app/routers/*.py` — FastAPI endpoints per flow (e.g. `job_order.py`)
- `app/services/*.py` — glue between extractor and Excel writer
- `app/schemas/*.py` — Pydantic request/response models
- `app/templates/*.html` + `app/static/js/*.js` — simple vanilla-JS upload UI
  per flow, all read-only result tables (no in-browser editing today)
- `app/config.py` — `UPLOAD_DIR` / `OUTPUT_DIR` paths

Generated workbooks are written to `outputs/`; uploads are staged in
`uploads/` and deleted after processing.

## Job Order flow specifics (`extract_job_order.py`, `excel_job_offer/writer.py`)

- `extract_job_order_data()` auto-detects (or is told) which of several PDF
  templates (`1c`, `running`, `completion_procedure`, `custom`) to parse, and
  returns `{"lines": [{"line_no": N, "text": "..."}, ...]}` **already in
  document order** — extraction does not need re-ordering logic.
- `write_job_offer_table()` appends those lines to the end of the
  `"JOB ORDER "` sheet (note trailing space in the sheet name), starting at
  the first fully-blank row (`_find_next_row`).
- Row "kind" (step / bullet / note / section / intro / text) is classified
  once in Python from the line's leading characters (`_prepare_row_values`)
  and determines styling (`_STYLE_TEMPLATE_ROWS`) and whether column A (step
  number) gets a formula.
- **Step numbering is a live Excel formula, not a static value**:
  `_step_order_formula(row)` returns
  `=LOOKUP(9.99E+307,$A$1:A{row-1})+1`. This means:
  - Only "step" rows get a number in column A; bullets/notes/sections stay
    blank so they don't count.
  - `LOOKUP(9.99E+307, range)` is the standard "last numeric value in a
    range" trick — it continues from whatever row is the last numbered row
    directly above, not the highest number anywhere above. This matters if a
    stray/out-of-order number ever ends up earlier in the column; `MAX()`
    would have jumped to that stray value, `LOOKUP` correctly ignores it and
    follows the sequential last-numbered row instead. **Do not switch this
    back to `MAX()` or a hardcoded Python-computed baseline** — both were
    tried and rejected (MAX picks the wrong row on out-of-order edits; a
    hardcoded baseline doesn't recalculate at all).
  - Because it's a live formula (not a Python-computed baseline), editing a
    row's text/number above, inserting/deleting rows, etc. causes every step
    formula below to recalculate automatically — no workbook regeneration
    needed.
  - Verified via LibreOffice headless recalculation (`soffice --headless
    --convert-to xlsx`) that editing an earlier step number correctly
    cascades to later rows.

## Testing notes

- No automated test suite exists yet (`find . -iname "test_*.py"` is empty).
- To sanity-check Excel formula behavior without opening real Excel, use
  LibreOffice headless to force recalculation, since `openpyxl` never
  evaluates formulas itself:
  ```bash
  soffice --headless --convert-to xlsx <file>.xlsm
  ```
  then read the resulting `.xlsx` with `openpyxl.load_workbook(path,
  data_only=True)` to see computed values. (First `soffice` invocation in a
  fresh sandbox/profile sometimes fails with a `DeploymentException` on the
  very first call — just retry once.)
- `venv/` in the repo already has `openpyxl` etc. installed; `source
  venv/bin/activate` before running ad-hoc scripts.

## Conventions

- Business/product docs (`business-analysis.txt`, `documentation.txt`) do
  not exist in this repo as of now — if they are added later, read them
  before planning changes and do not contradict them without an explicit
  request to do so.
