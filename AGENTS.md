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
Shared Excel workbook utilities live in `excel_utils/` (`workbook_utils.py` contains `find_sheet` and `validate_workbook_sheets` for whitespace-tolerant and case-insensitive sheet matching across all writers).

Web & Desktop layer:
- `desktop.py` — PyQt6 desktop GUI application (can be compiled to standalone executable)
- `app/routers/*.py` — FastAPI endpoints per flow (e.g. `job_order.py`)
- `app/services/*.py` — glue between extractor and Excel writer
- `app/schemas/*.py` — Pydantic request/response models
- `app/templates/*.html` + `app/static/js/*.js` — simple vanilla-JS upload UI
  per flow, all read-only result tables (no in-browser editing today)
- `app/config.py` — `UPLOAD_DIR` / `OUTPUT_DIR` paths, `REQUIRED_SHEETS`, `validate_template`

Generated workbooks are written to `outputs/`; uploads are staged in
`uploads/` and deleted after processing.

## SOE flow specifics (`extract_soe.py`, `excel_soe/writer.py`)

- SOE rows are written to the `"SOE"` sheet table starting at row 11
  (`DATA_START_ROW`), with Date in column A (merged A:C), Time in D, and
  Event in column E.
- **Event is always merged E:R** (`EVENT_COLUMN` → `EVENT_MERGE_END_COLUMN`)
  for every appended data row. The template's sample layout row often only
  merges Event to column O; `_copy_table_row_layout` / `_ensure_event_merge`
  expand that to R to match the header row. Table right-edge borders for
  Event are applied via `_apply_event_right_border`.

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
- **Table rows are tagged during extraction, not the writer.** PDFs
 sometimes contain data tables (parts/torque specs, etc.) introduced by a
 header line such as `QTY THREAD DESCRIPTION`, `THREAD DESCRIPTION`,
 `DESCRIPTION FUNCTION`, uppercase `SUB-ASSEMBLY # N`, `SEAL TEST PORT`, or
 `HOLD POINT` (`_is_table_header_line` in `extract_job_order.py`). Uppercase
 `SUB-ASSEMBLY` is required — lowercase wrap continuations like
 `sub-assembly # 2. Fill the tubing.` must stay as step text, not headers.
 Lines following one of these headers are tagged `is_table=True` until the
 table's paragraph ends — a blank line in the PDF text, or (as a fallback)
 a real procedure step line (`_is_procedure_step_line`: dotted steps like
 `7. M/U` / `10.P/U`, or `x.y` steps). Bare qty digits (`1 4-1/2" ...`) and
 measurements (`5.91" OD ...`) do **not** end the table. `_NUMBERED_STEP_DOT`
 uses `(?!\d)` so `5.91"` is never treated as step `5.`. `_merge_wrapped_lines`
 also refuses to glue table rows onto non-table steps (and vice versa), so
 broken PDF fragments like `QT`/`Y` from `QTY` cannot flip a step to
 `is_table=True`. `_extract_section_lines` / `_merge_wrapped_lines` /
 `_build_result` all carry `(text, is_table)` tuples through, and
 `_build_result` puts `is_table` on each line dict. `write_job_offer_table`
 (`excel_job_offer/writer.py`) forces `is_table` lines to kind `"text"` even
 if they'd otherwise look like a numbered step (e.g. start with a bare
 digit), so they're appended as plain rows under the previous step number
 instead of creating a new numbered row.
- **Job Order text font is fixed to Abadi 12** for every appended row
 (steps, bullets, notes, sections, text), set explicitly in
 `_write_formatted_row` (`_FONT_NAME` / `_FONT_SIZE` constants) rather than
 inherited from the template row's font — only bold/italic/color are kept
 from the existing cell font.

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
