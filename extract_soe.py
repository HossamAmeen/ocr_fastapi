"""Extract Time Log data from SOE daily operations report PDFs."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pdfplumber

from pdf_extractor.extractor import _extract_raw_text

_TIME_LOG_ROW = re.compile(
    r"^(\d{1,2}:\d{2})\s+([\d.]+)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\w)\s+(.+)$"
)
_NOTE_ROW = re.compile(r"^\*+\s*(.+)$")
_OAMN_RANGE = re.compile(
    r"^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*Hr'?s?\s*$",
    re.IGNORECASE,
)
_REPORT_PERIOD = re.compile(
    r"Rpt\. Period:\s*(.+?)\s+to\s+(.+?)(?:\n|$)",
    re.IGNORECASE,
)
_WELL_NAME = re.compile(r"Well Name:\s*(\S+)")
_REPORT_DATE = re.compile(r"DATE:\s*(.+?)(?:\n|$)", re.IGNORECASE)
_FROM_TIME = re.compile(r"^\d{1,2}:\d{2}$")


def _to_float(value: str) -> float:
    return float(value.replace(",", ""))


def _finalize_entry(entry: dict[str, Any] | None) -> dict[str, Any] | None:
    if not entry:
        return None
    entry["operation"] = entry["operation"].strip()
    entry["notes"] = [note.strip() for note in entry["notes"] if note.strip()]
    return entry


def _parse_time_log_section(section: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    lines = [line.strip() for line in section.splitlines() if line.strip()]
    if len(lines) >= 2 and lines[0].lower() == "time log":
        lines = lines[2:]
    elif lines and "to duration phase code sub type operation" in lines[0].lower():
        lines = lines[1:]

    entries: list[dict[str, Any]] = []
    oamn_entries: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    pending_notes: list[str] = []
    in_oamn = False
    pending_oamn: dict[str, Any] | None = None

    for line in lines:
        if line.upper() == "O.A.M.N":
            finalized = _finalize_entry(current)
            if finalized:
                entries.append(finalized)
            current = None
            pending_notes = []
            in_oamn = True
            continue
        if in_oamn and line.startswith("="):
            continue

        row_match = _TIME_LOG_ROW.match(line)
        if row_match:
            finalized = _finalize_entry(current)
            if finalized:
                entries.append(finalized)
            current = {
                "to": row_match.group(1),
                "duration": _to_float(row_match.group(2)),
                "phase": row_match.group(3),
                "code": row_match.group(4),
                "sub": row_match.group(5),
                "type": row_match.group(6),
                "operation": row_match.group(7).strip(),
                "notes": pending_notes.copy(),
            }
            pending_notes = []
            continue

        note_match = _NOTE_ROW.match(line)
        if note_match:
            note = note_match.group(1).strip()
            if in_oamn:
                if pending_oamn is not None:
                    pending_oamn["notes"].append(note)
            elif current is not None:
                current["notes"].append(note)
            else:
                pending_notes.append(note)
            continue

        oamn_match = _OAMN_RANGE.match(line)
        if oamn_match and in_oamn:
            if pending_oamn:
                oamn_entries.append(pending_oamn)
            pending_oamn = {
                "from": oamn_match.group(1),
                "to": oamn_match.group(2),
                "operation": "",
                "notes": [],
            }
            continue

        if in_oamn and pending_oamn is not None:
            if pending_oamn["operation"]:
                pending_oamn["operation"] = f"{pending_oamn['operation']} {line}".strip()
            else:
                pending_oamn["operation"] = line
            continue

        if current is not None:
            current["operation"] = f"{current['operation']} {line}".strip()

    finalized = _finalize_entry(current)
    if finalized:
        entries.append(finalized)
    if pending_oamn:
        oamn_entries.append(pending_oamn)

    return entries, oamn_entries


def extract_time_log(pdf_path: str | Path) -> dict[str, Any]:
    """Extract the Time Log table from an SOE daily operations report PDF."""
    text = _extract_raw_text(Path(pdf_path))

    start = text.find("Time Log")
    if start < 0:
        return {
            "well_name": _WELL_NAME.search(text).group(1) if _WELL_NAME.search(text) else "",
            "report_period_from": "",
            "report_period_to": "",
            "entries": [],
            "oamn_entries": [],
            "total_duration": 0.0,
        }

    end = text.find("Phase Time and Cost Summary", start)
    section = text[start:end].strip() if end > 0 else text[start:].strip()
    entries, oamn_entries = _parse_time_log_section(section)

    period_match = _REPORT_PERIOD.search(text)
    well_match = _WELL_NAME.search(text)

    return {
        "well_name": well_match.group(1) if well_match else "",
        "report_period_from": period_match.group(1).strip() if period_match else "",
        "report_period_to": period_match.group(2).strip() if period_match else "",
        "entries": entries,
        "oamn_entries": oamn_entries,
        "total_duration": round(sum(entry["duration"] for entry in entries), 2),
    }


def _cell_text(value: object) -> str:
    return str(value).strip() if value else ""


def _find_table_columns(cells: list[str]) -> tuple[int | None, int | None]:
    from_col: int | None = None
    details_col: int | None = None
    for index, cell in enumerate(cells):
        label = cell.lower()
        if label == "from":
            from_col = index
        if "operation details" in label:
            details_col = index
    return from_col, details_col


def extract_operational_time_summary(pdf_path: str | Path) -> dict[str, Any]:
    """Extract DATE and Operational Time Summary rows from a drilling report PDF."""
    pdf_path = Path(pdf_path)
    report_date = ""
    entries: list[dict[str, str]] = []

    with pdfplumber.open(pdf_path) as document:
        for page in document.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 3:
                    continue

                header_row_index: int | None = None
                from_col: int | None = None
                details_col: int | None = None

                for row_index, row in enumerate(table):
                    cells = [_cell_text(cell) for cell in row]
                    joined = " ".join(cells)

                    if not report_date and "DATE:" in joined.upper():
                        match = _REPORT_DATE.search(joined)
                        if match:
                            report_date = match.group(1).strip()

                    found_from, found_details = _find_table_columns(cells)
                    if found_from is not None and found_details is not None:
                        header_row_index = row_index
                        from_col = found_from
                        details_col = found_details
                        break

                if header_row_index is None or from_col is None or details_col is None:
                    continue

                for row in table[header_row_index + 1 :]:
                    cells = [_cell_text(cell) for cell in row]
                    if not any(cells):
                        continue

                    time_from = cells[from_col] if from_col < len(cells) else ""
                    operation_details = cells[details_col] if details_col < len(cells) else ""
                    if not _FROM_TIME.match(time_from):
                        break

                    entries.append(
                        {
                            "from": time_from,
                            "operation_details": operation_details,
                        }
                    )

    if not report_date:
        text = _extract_raw_text(pdf_path)
        match = _REPORT_DATE.search(text)
        if match:
            report_date = match.group(1).strip()

    return {
        "source": "operational_time_summary",
        "date": report_date,
        "entries": entries,
    }


def extract_soe_data(pdf_path: str | Path) -> dict[str, Any]:
    """Extract SOE rows from either Time Log or Operational Time Summary PDFs."""
    text = _extract_raw_text(Path(pdf_path))
    if "Operational Time Summary" in text:
        return extract_operational_time_summary(pdf_path)

    data = extract_time_log(pdf_path)
    data["source"] = "time_log"
    return data


if __name__ == "__main__":
    import json
    import sys

    pdf = sys.argv[1] if len(sys.argv) > 1 else "SOE/template/1.pdf"
    print(json.dumps(extract_soe_data(pdf), indent=2, ensure_ascii=False))
