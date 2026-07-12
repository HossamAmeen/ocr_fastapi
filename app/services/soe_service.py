from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path

from excel_soe.writer import read_template_rig, soe_data_to_rows, write_soe_table
from extract_soe import extract_soe_data

from app.config import OUTPUT_DIR


def _format_row_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%y")
    return str(value or "")


def _pdf_summary(data: dict, filename: str, row_count: int, skipped: bool) -> dict:
    if data.get("source") == "operational_time_summary":
        return {
            "filename": filename,
            "source": data.get("source", ""),
            "well_name": "",
            "rig": data.get("rig", ""),
            "report_date": data.get("date", ""),
            "report_period_from": "",
            "report_period_to": "",
            "row_count": row_count,
            "skipped": skipped,
            "skip_reason": "rig_mismatch" if data.get("skipped_rig_mismatch") else "",
        }

    skip_reason = ""
    if skipped and data.get("skipped_rig_mismatch"):
        skip_reason = "rig_mismatch"

    return {
        "filename": filename,
        "source": data.get("source", "time_log"),
        "well_name": data.get("well_name", ""),
        "rig": data.get("rig", ""),
        "report_date": "",
        "report_period_from": data.get("report_period_from", ""),
        "report_period_to": data.get("report_period_to", ""),
        "row_count": row_count,
        "skipped": skipped,
        "skip_reason": skip_reason,
    }


def process_soe(
    pdf_entries: list[tuple[Path, str]],
    template_path: Path,
) -> tuple[list[dict], list[dict], Path, int]:
    """Extract time logs from multiple PDFs and append them into one Excel workbook.

    When the Excel template has a Rig value, every PDF page whose ``Rig:``
    field matches that value is extracted (not only the first page).
    """
    if not pdf_entries:
        raise ValueError("At least one PDF is required.")

    pdf_entries = sorted(pdf_entries, key=lambda entry: entry[1].lower())
    rig_filter = read_template_rig(template_path) or None

    output_name = f"soe_{uuid.uuid4().hex[:8]}.xlsm"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    pdf_summaries: list[dict] = []
    all_rows: list[dict] = []
    total_appended = 0

    for pdf_path, display_name in pdf_entries:
        data = extract_soe_data(pdf_path, rig_filter=rig_filter)
        rows = soe_data_to_rows(data)
        if not rows:
            pdf_summaries.append(_pdf_summary(data, display_name, 0, True))
            continue

        _, appended = write_soe_table(
            output_path,
            output_path,
            data,
            template_path=template_path,
        )
        total_appended += appended
        pdf_summaries.append(_pdf_summary(data, display_name, appended, False))
        for row in rows:
            all_rows.append(
                {
                    "date": _format_row_date(row["date"]),
                    "time": str(row["time"]),
                    "event": str(row["event"]),
                }
            )

    if total_appended == 0:
        output_path.unlink(missing_ok=True)
        if rig_filter:
            found = sorted(
                {
                    str(summary.get("rig") or "").strip()
                    for summary in pdf_summaries
                    if str(summary.get("rig") or "").strip()
                }
            )
            found_text = ", ".join(found) if found else "none"
            raise ValueError(
                f"No time-log rows found for Rig '{rig_filter}'. "
                f"PDF Rig values found: {found_text}. "
                "Set the SOE sheet Rig cell to match a PDF Rig: value."
            )
        raise ValueError("No time-log rows found in any uploaded PDF.")

    return pdf_summaries, all_rows, output_path, total_appended
