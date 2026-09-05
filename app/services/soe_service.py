from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path

from excel_soe.writer import (
    collapse_repeated_row_dates,
    read_template_rig,
    soe_data_to_rows,
    sort_soe_rows,
    write_soe_rows,
)
from extract_soe import extract_soe_data, format_rig_for_display, rigs_match

from app.config import OUTPUT_DIR
from app.services.extraction_errors import (
    format_extraction_error,
    soe_extraction_mode_label,
)


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


def _pdf_summary(
    data: dict,
    filename: str,
    row_count: int,
    skipped: bool,
    *,
    is_summary: bool = False,
) -> dict:
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
            "extraction_mode": "table",
        }

    if data.get("source") == "paragraph_summary":
        summary = {
            "filename": filename,
            "source": data.get("source", ""),
            "well_name": data.get("well_name", ""),
            "rig": rig_display,
            "report_date": "",
            "report_period_from": data.get("report_period_from", ""),
            "report_period_to": data.get("report_period_to", ""),
            "row_count": row_count,
            "skipped": skipped,
            "skip_reason": skip_reason,
            "extraction_mode": "summary",
        }
        hints = data.get("title_hints")
        if hints:
            summary["title_hints"] = hints
        return summary

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
        "extraction_mode": "summary" if is_summary else "table",
    }


def _format_table_names(table_names: list[str] | None) -> list[str]:
    if table_names:
        return table_names
    return ["Time Log", "Job Time Log", "Operational Time Summary"]


def _resolve_soe_summary_mode(
    pdf_summaries: list[dict],
    is_summary: bool,
) -> bool:
    if is_summary:
        return True
    modes = {str(summary.get("extraction_mode") or "") for summary in pdf_summaries}
    if modes == {"summary"}:
        return True
    if "summary" in modes and "table" not in modes:
        return True
    return any(summary.get("source") == "paragraph_summary" for summary in pdf_summaries)


def _build_no_rows_error(
    pdf_summaries: list[dict],
    rig_filter: str | None,
    table_names: list[str] | None,
    *,
    is_summary: bool = False,
) -> str:
    is_summary = _resolve_soe_summary_mode(pdf_summaries, is_summary)
    mode_label = soe_extraction_mode_label(is_summary)
    table_list = _format_table_names(table_names)
    pdf_count = len(pdf_summaries)

    rig_mismatch = [s for s in pdf_summaries if s.get("skip_reason") == "rig_mismatch"]
    no_table = [s for s in pdf_summaries if s.get("skip_reason") == "no_matching_table"]
    empty_table = [s for s in pdf_summaries if s.get("skip_reason") == "empty_table"]

    title_hints = sorted(
        {
            hint.strip()
            for summary in pdf_summaries
            for hint in (summary.get("title_hints") or [])
            if str(hint).strip()
        }
    )

    pdf_rigs = sorted(
        {
            str(summary.get("rig") or "").strip()
            for summary in pdf_summaries
            if str(summary.get("rig") or "").strip()
        }
    )
    expanded_rigs: list[str] = []
    for rig in pdf_rigs:
        for part in rig.split(","):
            cleaned = part.strip()
            if cleaned and cleaned not in expanded_rigs:
                expanded_rigs.append(cleaned)
    rigs_text = ", ".join(expanded_rigs) if expanded_rigs else "none detected"

    if rig_mismatch:
        return format_extraction_error(
            "SOE",
            [
                "Cause: Rig mismatch detected.",
                f"• Excel Rig filter: {rig_filter}",
                f"• Rigs found in PDFs: {rigs_text}",
                "",
                "Solution:",
                "• Change the Rig name in the Excel template to match one of the PDF rigs above.",
                "• Or leave the Rig cell empty in Excel to extract from all rigs.",
            ],
            mode=mode_label,
        )

    if no_table:
        if is_summary:
            body = [
                "Cause: No paragraph with the specified title was found in the PDFs.",
                f"• Paragraph titles searched: {', '.join(table_list)}",
                f"• PDFs checked: {pdf_count}",
            ]
            if title_hints:
                body.append(f"• Similar lines found in PDF: {', '.join(title_hints)}")
            body.extend(
                [
                    "",
                    "Solution:",
                    "• Open the PDF and copy the exact paragraph heading line.",
                    "• Paste that exact title into the paragraph title field.",
                    "• Summary mode expects the heading as its own line, or at the start of a line.",
                ]
            )
        else:
            body = [
                "Cause: No tables matching the specified names were found in the PDFs.",
                f"• Tables searched: {', '.join(table_list)}",
                f"• PDFs checked: {pdf_count}",
                "",
                "Solution:",
                "• Check the table heading title in your PDF file.",
                "• Add that exact table title to 'Table names to extract'.",
            ]
        return format_extraction_error("SOE", body, mode=mode_label)

    if empty_table:
        if is_summary:
            body = [
                "Cause: The paragraph title was found, but no text followed it.",
                "",
                "Solution:",
                "• Ensure the paragraph text appears on the same line or on the lines right after the title.",
            ]
        else:
            body = [
                "Cause: The table title was found, but it contained no valid data rows.",
                "",
                "Solution:",
                "• Ensure the PDF contains readable text and valid data rows.",
            ]
        return format_extraction_error("SOE", body, mode=mode_label)

    body = [
        "Cause: No valid time-log data could be extracted.",
        "",
        "Solution:",
        (
            "• Confirm that the uploaded PDFs contain a paragraph with the specified title."
            if is_summary
            else "• Confirm that the uploaded PDFs contain a time-log table."
        ),
        (
            "• Check the paragraph title field and add the exact heading from the PDF."
            if is_summary
            else "• Check 'Table names to extract' and add the exact title from the PDF."
        ),
    ]
    return format_extraction_error("SOE", body, mode=mode_label)


def process_soe(
    pdf_entries: list[tuple[Path, str]],
    template_path: Path,
    *,
    table_names: list[str] | None = None,
    is_summary: bool = False,
    progress_callback: callable | None = None,
    is_cancelled: callable | None = None,
) -> tuple[list[dict], list[dict], Path, int, str]:
    """Extract time logs from multiple PDFs and write them into one Excel workbook.

    When the Excel template has a Rig value, every PDF page whose ``Rig:``
    field matches that value is extracted (not only the first page).
    Rows from all PDFs are collected and sorted by date and time before writing.

    When ``is_summary`` is true, no table is read — the paragraph following
    the title given in ``table_names`` is captured from each PDF and written
    as a single Excel row (see ``extract_paragraph_summary``).
    """
    if not pdf_entries:
        raise ValueError("At least one PDF is required.")

    if is_cancelled and is_cancelled():
        raise InterruptedError("Operation was cancelled by the user.")

    pdf_entries = sorted(pdf_entries, key=lambda entry: entry[1].lower())
    rig_filter = read_template_rig(template_path) or None
    normalized_table_names = [
        name.strip()
        for name in (table_names or [])
        if name and name.strip()
    ] or None

    suffix = template_path.suffix.lower()
    output_name = f"soe_{uuid.uuid4().hex[:8]}{suffix}"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    total_soe_files = len(pdf_entries)
    extracted_results: list[tuple[int, dict, str, list[dict]]] = []

    def _extract_single(idx: int, pdf_path: Path, display_name: str):
        if is_cancelled and is_cancelled():
            return idx, {}, display_name, []
        data = extract_soe_data(
            pdf_path,
            rig_filter=rig_filter,
            table_names=normalized_table_names,
            is_summary=is_summary,
        )
        rows = soe_data_to_rows(data)
        return idx, data, display_name, rows

    import concurrent.futures
    max_workers = min(8, max(1, total_soe_files))
    completed_count = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(_extract_single, i, p, d): (i, d)
            for i, (p, d) in enumerate(pdf_entries)
        }
        for future in concurrent.futures.as_completed(future_to_idx):
            if is_cancelled and is_cancelled():
                executor.shutdown(wait=False, cancel_futures=True)
                output_path.unlink(missing_ok=True)
                raise InterruptedError("Operation was cancelled by the user.")
            idx, data, display_name, rows = future.result()
            extracted_results.append((idx, data, display_name, rows))
            completed_count += 1
            if progress_callback:
                progress_callback(completed_count, total_soe_files, f"Processed {completed_count}/{total_soe_files}: {display_name}")

    if is_cancelled and is_cancelled():
        output_path.unlink(missing_ok=True)
        raise InterruptedError("Operation was cancelled by the user.")

    # Re-sort results back to original document order
    extracted_results.sort(key=lambda x: x[0])

    pdf_summaries: list[dict] = []
    all_row_data: list[dict] = []

    for _, data, display_name, rows in extracted_results:
        if not rows:
            pdf_summaries.append(_pdf_summary(data, display_name, 0, True, is_summary=is_summary))
            continue

        pdf_summaries.append(_pdf_summary(data, display_name, len(rows), False, is_summary=is_summary))
        all_row_data.extend(rows)

    if not all_row_data:
        output_path.unlink(missing_ok=True)
        raise ValueError(
            _build_no_rows_error(
                pdf_summaries,
                rig_filter,
                normalized_table_names,
                is_summary=is_summary,
            )
        )

    sorted_rows = sort_soe_rows(all_row_data)
    _, total_appended = write_soe_rows(
        output_path,
        output_path,
        sorted_rows,
        template_path=template_path,
    )

    display_rows = collapse_repeated_row_dates(sorted_rows)
    all_rows = [
        {
            "date": _format_row_date(row["date"]) if row.get("date") else "",
            "time": str(row["time"]),
            "event": str(row["event"]),
        }
        for row in display_rows
    ]

    return pdf_summaries, all_rows, output_path, total_appended, rig_filter or ""
