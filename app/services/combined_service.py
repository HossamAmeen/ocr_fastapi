from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path

from excel_job_offer.writer import write_job_offer_table, _prepare_row_values
from excel_proforma.writer import write_proforma_table
from excel_soe.writer import collapse_repeated_row_dates, read_template_rig, soe_data_to_rows, sort_soe_rows, write_soe_rows
from extract_job_order import extract_job_order_data
from extract_performa import extract_proforma_items
from extract_soe import extract_soe_data

from app.config import OUTPUT_DIR
from app.services.extraction_errors import format_extraction_error
from app.services.soe_service import _build_no_rows_error, _pdf_summary


def _format_row_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%y")
    return str(value or "")


def process_combined(
    template_path: Path,
    *,
    proforma_pdf: Path | None = None,
    soe_pdfs: list[tuple[Path, str]] | None = None,
    job_order_pdf: Path | None = None,
    job_order_source: str = "auto",
    job_order_start_marker: str = "",
    job_order_end_marker: str = "",
    soe_table_names: list[str] | None = None,
    soe_is_summary: bool = False,
    progress_callback: callable | None = None,
    is_cancelled: callable | None = None,
) -> tuple[dict, Path]:
    """Extract selected PDFs and write all sections into one Excel workbook."""
    if not proforma_pdf and not soe_pdfs and not job_order_pdf:
        raise ValueError("Provide at least one PDF: Proforma, SOE, or Job Order.")

    suffix = template_path.suffix.lower()
    output_name = f"workbook_{uuid.uuid4().hex[:8]}{suffix}"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    processed_sections: list[str] = []
    result: dict = {"processed_sections": processed_sections}

    if proforma_pdf:
        if is_cancelled and is_cancelled():
            output_path.unlink(missing_ok=True)
            raise InterruptedError("Operation was cancelled by the user.")
        if progress_callback:
            progress_callback(0, 1, f"Extracting Proforma: {proforma_pdf.name}")
        items = extract_proforma_items(proforma_pdf)
        if not items:
            output_path.unlink(missing_ok=True)
            raise ValueError(
                format_extraction_error(
                    "Proforma",
                    [
                        "Cause: No Proforma line items were found in the uploaded PDF.",
                        "",
                        "Solution:",
                        "• Ensure the uploaded file is a valid Proforma Purchase Order PDF containing price items.",
                    ],
                )
            )

        for index, item in enumerate(items, start=1):
            item["sno"] = index
            item["total"] = item["per_day_rate"] * item["days"]

        write_proforma_table(output_path, output_path, items)
        processed_sections.append("proforma")
        result["proforma"] = {
            "items": items,
            "item_count": len(items),
            "gross_total": sum(item["total"] for item in items),
        }

    if soe_pdfs:
        if is_cancelled and is_cancelled():
            output_path.unlink(missing_ok=True)
            raise InterruptedError("Operation was cancelled by the user.")

        pdf_entries = sorted(soe_pdfs, key=lambda entry: entry[1].lower())
        rig_filter = read_template_rig(template_path) or None
        normalized_table_names = [
            name.strip()
            for name in (soe_table_names or [])
            if name and name.strip()
        ] or None

        total_soe_files = len(pdf_entries)
        extracted_results: list[tuple[int, dict, str, list[dict]]] = []

        def _extract_single(idx: int, pdf_path: Path, display_name: str):
            if is_cancelled and is_cancelled():
                return idx, {}, display_name, []
            data = extract_soe_data(
                pdf_path,
                rig_filter=rig_filter,
                table_names=normalized_table_names,
                is_summary=soe_is_summary,
            )
            rows = soe_data_to_rows(data)
            return idx, data, display_name, rows

        # Use ThreadPoolExecutor for multi-file parallel extraction
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
                pdf_summaries.append(
                    _pdf_summary(data, display_name, 0, True, is_summary=soe_is_summary)
                )
                continue

            pdf_summaries.append(
                _pdf_summary(data, display_name, len(rows), False, is_summary=soe_is_summary)
            )
            all_row_data.extend(rows)

        if not all_row_data:
            output_path.unlink(missing_ok=True)
            raise ValueError(
                _build_no_rows_error(
                    pdf_summaries,
                    rig_filter,
                    normalized_table_names,
                    is_summary=soe_is_summary,
                )
            )

        sorted_rows = sort_soe_rows(all_row_data)
        _, total_appended = write_soe_rows(
            output_path,
            output_path,
            sorted_rows,
            template_path=template_path,
        )

        processed_sections.append("soe")
        display_rows = collapse_repeated_row_dates(sorted_rows)
        result["soe"] = {
            "pdf_summaries": pdf_summaries,
            "rows": [
                {
                    "date": _format_row_date(row["date"]) if row.get("date") else "",
                    "time": str(row["time"]),
                    "event": str(row["event"]),
                }
                for row in display_rows
            ],
            "row_count": total_appended,
            "pdf_count": len(pdf_entries),
            "rig_filter": rig_filter or "",
        }

    if job_order_pdf:
        if is_cancelled and is_cancelled():
            output_path.unlink(missing_ok=True)
            raise InterruptedError("Operation was cancelled by the user.")
        if progress_callback:
            progress_callback(0, 1, f"Extracting Job Order: {job_order_pdf.name}")
        data = extract_job_order_data(
            job_order_pdf,
            source=job_order_source,
            start_marker=job_order_start_marker,
            end_marker=job_order_end_marker,
        )
        if not data.get("lines"):
            output_path.unlink(missing_ok=True)
            if job_order_source == "custom":
                raise ValueError(
                    format_extraction_error(
                        "Job Order",
                        [
                            f"Cause: No procedure content found starting at {job_order_start_marker!r}.",
                            "",
                            "Solution:",
                            "• Check that the start text matches the PDF exactly.",
                        ],
                    )
                )
            raise ValueError(
                format_extraction_error(
                    "Job Order",
                    [
                        "Cause: No completion procedure content was found in the uploaded PDF.",
                        "",
                        "Solution:",
                        "• Ensure the PDF contains readable completion procedure text.",
                        "• Try a different extraction template if the PDF layout differs.",
                    ],
                )
            )

        _, appended_rows = write_job_offer_table(output_path, output_path, data)
        processed_sections.append("job_order")
        result["job_order"] = {
            "section_title": str(data.get("section_title") or ""),
            "source": str(data.get("source") or job_order_source),
            "lines": [
                {
                    "line_no": row["line_no"],
                    "text": row["text"],
                    "kind": _prepare_row_values(row["text"], data.get("source", ""))[0]
                }
                for row in data["lines"]
            ],
            "line_count": appended_rows,
        }

    return result, output_path
