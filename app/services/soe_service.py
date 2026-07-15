from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path

from excel_soe.writer import (
    read_template_rig,
    soe_data_to_rows,
    sort_soe_rows,
    write_soe_rows,
)
from extract_soe import extract_soe_data, format_rig_for_display, rigs_match

from app.config import OUTPUT_DIR


def parse_table_names(values: list[str]) -> list[str] | None:
    names: list[str] = []
    for value in values:
        for part in value.replace("\n", ",").split(","):
            cleaned = part.strip()
            if cleaned:
                names.append(cleaned)
    return names or None


def _format_row_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%y")
    return str(value or "")


def _pdf_summary(data: dict, filename: str, row_count: int, skipped: bool) -> dict:
    skip_reason = str(data.get("skip_reason") or "")
    if skipped and not skip_reason and data.get("skipped_rig_mismatch"):
        skip_reason = "rig_mismatch"
    if skipped and not skip_reason:
        skip_reason = "no_matching_table"

    rig_display = format_rig_for_display(str(data.get("rig") or ""))

    if data.get("source") == "operational_time_summary":
        return {
            "filename": filename,
            "source": data.get("source", ""),
            "well_name": "",
            "rig": rig_display,
            "report_date": data.get("date", ""),
            "report_period_from": "",
            "report_period_to": "",
            "row_count": row_count,
            "skipped": skipped,
            "skip_reason": skip_reason,
        }

    return {
        "filename": filename,
        "source": data.get("source", "time_log"),
        "well_name": data.get("well_name", ""),
        "rig": rig_display,
        "report_date": "",
        "report_period_from": data.get("report_period_from", ""),
        "report_period_to": data.get("report_period_to", ""),
        "row_count": row_count,
        "skipped": skipped,
        "skip_reason": skip_reason,
    }


def _format_table_names(table_names: list[str] | None) -> list[str]:
    if table_names:
        return table_names
    return ["Time Log", "Job Time Log"]


def _build_no_rows_error(
    pdf_summaries: list[dict],
    rig_filter: str | None,
    table_names: list[str] | None,
) -> str:
    table_list = _format_table_names(table_names)
    pdf_count = len(pdf_summaries)

    rig_mismatch = [s for s in pdf_summaries if s.get("skip_reason") == "rig_mismatch"]
    no_table = [s for s in pdf_summaries if s.get("skip_reason") == "no_matching_table"]
    empty_table = [s for s in pdf_summaries if s.get("skip_reason") == "empty_table"]

    pdf_rigs = sorted(
        {
            str(summary.get("rig") or "").strip()
            for summary in pdf_summaries
            if str(summary.get("rig") or "").strip()
        }
    )
    # Expand comma-separated rig lists from multi-page PDF summaries.
    expanded_rigs: list[str] = []
    for rig in pdf_rigs:
        for part in rig.split(","):
            cleaned = part.strip()
            if cleaned and cleaned not in expanded_rigs:
                expanded_rigs.append(cleaned)
    rigs_text = ", ".join(expanded_rigs) if expanded_rigs else "none detected"

    lines = [
        "No time-log rows could be extracted from the uploaded PDFs.",
        "",
        f"PDFs checked: {pdf_count}",
        f"Table names searched: {', '.join(table_list)}",
    ]

    if rig_filter:
        lines.append(f"Excel Rig filter: {rig_filter}")
        lines.append(f"Rigs found in PDFs: {rigs_text}")
        rig_matches = any(
            rigs_match(rig_filter, rig)
            for rig in pdf_rigs
        )
    else:
        lines.append(f"Rigs found in PDFs: {rigs_text}")
        rig_matches = False

    lines.append("")

    if rig_mismatch:
        lines.extend(
            [
                "What went wrong:",
                f"- No pages matched the Excel Rig filter ({rig_filter!r}).",
                f"- Rigs found across PDF pages: {rigs_text}.",
                "",
                "What to do:",
                "- The PDF is scanned page by page; only pages whose Rig matches Excel are used.",
                "- Set the SOE sheet Rig cell to one of the rig codes listed above.",
                "- Or clear the Rig cell to extract from all pages/rigs.",
            ]
        )
        return "\n".join(lines)

    if rig_filter and rig_matches:
        lines.extend(
            [
                "What went wrong:",
                f"- The Excel Rig ({rig_filter!r}) matches the PDF ({rigs_text}), "
                "so this is not a Rig mismatch.",
                f"- {len(no_table)} PDF(s) had no table matching: {', '.join(table_list)}.",
            ]
        )
        if empty_table:
            lines.append(
                f"- {len(empty_table)} PDF(s) had a matching table title but no data rows."
            )
        lines.extend(
            [
                "",
                "What to do:",
                "- Check the exact table title in your PDF and add it to the "
                "'Table names to extract' list.",
                "- Common titles: Time Log, Job Time Log, Operational Time Summary.",
            ]
        )
        if no_table:
            lines.append("")
            lines.append("PDFs with no matching table:")
            for summary in no_table[:8]:
                lines.append(f"  • {summary['filename']}")
            if len(no_table) > 8:
                lines.append(f"  • …and {len(no_table) - 8} more")
        return "\n".join(lines)

    if no_table or empty_table:
        lines.extend(
            [
                "What went wrong:",
                f"- No PDF contained a table matching: {', '.join(table_list)}.",
            ]
        )
        if empty_table:
            lines.append("- Some PDFs had a matching title but the table was empty.")
        lines.extend(
            [
                "",
                "What to do:",
                "- Add the exact table title from your PDF to 'Table names to extract'.",
            ]
        )
        return "\n".join(lines)

    lines.extend(
        [
            "What went wrong:",
            "- The PDFs were read, but no time-log rows were found.",
            "",
            "What to do:",
            "- Confirm the PDF contains a time-log table.",
            "- Add the exact table title to 'Table names to extract'.",
        ]
    )
    return "\n".join(lines)


def process_soe(
    pdf_entries: list[tuple[Path, str]],
    template_path: Path,
    *,
    table_names: list[str] | None = None,
) -> tuple[list[dict], list[dict], Path, int]:
    """Extract time logs from multiple PDFs and write them into one Excel workbook.

    When the Excel template has a Rig value, every PDF page whose ``Rig:``
    field matches that value is extracted (not only the first page).
    Rows from all PDFs are collected and sorted by date and time before writing.
    """
    if not pdf_entries:
        raise ValueError("At least one PDF is required.")

    pdf_entries = sorted(pdf_entries, key=lambda entry: entry[1].lower())
    rig_filter = read_template_rig(template_path) or None
    normalized_table_names = [
        name.strip()
        for name in (table_names or [])
        if name and name.strip()
    ] or None

    output_name = f"soe_{uuid.uuid4().hex[:8]}.xlsm"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    pdf_summaries: list[dict] = []
    all_row_data: list[dict] = []

    for pdf_path, display_name in pdf_entries:
        data = extract_soe_data(
            pdf_path,
            rig_filter=rig_filter,
            table_names=normalized_table_names,
        )
        rows = soe_data_to_rows(data)
        if not rows:
            pdf_summaries.append(_pdf_summary(data, display_name, 0, True))
            continue

        pdf_summaries.append(_pdf_summary(data, display_name, len(rows), False))
        all_row_data.extend(rows)

    if not all_row_data:
        output_path.unlink(missing_ok=True)
        raise ValueError(
            _build_no_rows_error(pdf_summaries, rig_filter, normalized_table_names)
        )

    sorted_rows = sort_soe_rows(all_row_data)
    _, total_appended = write_soe_rows(
        output_path,
        output_path,
        sorted_rows,
        template_path=template_path,
    )

    all_rows = [
        {
            "date": _format_row_date(row["date"]),
            "time": str(row["time"]),
            "event": str(row["event"]),
        }
        for row in sorted_rows
    ]

    return pdf_summaries, all_rows, output_path, total_appended
