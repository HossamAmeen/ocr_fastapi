from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path

from excel_job_offer.writer import write_job_offer_table
from excel_proforma.writer import write_proforma_table
from excel_soe.writer import read_template_rig, soe_data_to_rows, write_soe_table
from extract_job_order import extract_job_order_data
from extract_performa import extract_proforma_items
from extract_soe import extract_soe_data

from app.config import OUTPUT_DIR
from app.services.soe_service import _pdf_summary


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
        items = extract_proforma_items(proforma_pdf)
        if not items:
            output_path.unlink(missing_ok=True)
            raise ValueError("No Proforma line items found in the uploaded PDF.")

        for index, item in enumerate(items, start=1):
            item.setdefault("sno", index)
            item["total"] = item["per_day_rate"] * item["days"]

        write_proforma_table(output_path, output_path, items)
        processed_sections.append("proforma")
        result["proforma"] = {
            "items": items,
            "item_count": len(items),
            "gross_total": sum(item["total"] for item in items),
        }

    if soe_pdfs:
        pdf_entries = sorted(soe_pdfs, key=lambda entry: entry[1].lower())
        pdf_summaries: list[dict] = []
        all_rows: list[dict] = []
        total_appended = 0
        rig_filter = read_template_rig(template_path) or None

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
            raise ValueError("No time-log rows found in any uploaded SOE PDF.")

        processed_sections.append("soe")
        result["soe"] = {
            "pdf_summaries": pdf_summaries,
            "rows": all_rows,
            "row_count": total_appended,
            "pdf_count": len(pdf_entries),
            "rig_filter": rig_filter or "",
        }

    if job_order_pdf:
        data = extract_job_order_data(job_order_pdf, source=job_order_source)
        if not data.get("lines"):
            output_path.unlink(missing_ok=True)
            raise ValueError("No completion procedure content found in the Job Order PDF.")

        _, appended_rows = write_job_offer_table(output_path, output_path, data)
        processed_sections.append("job_order")
        result["job_order"] = {
            "section_title": str(data.get("section_title") or ""),
            "source": str(data.get("source") or job_order_source),
            "lines": [
                {"line_no": row["line_no"], "text": row["text"]}
                for row in data["lines"]
            ],
            "line_count": appended_rows,
        }

    return result, output_path
