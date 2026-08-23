from __future__ import annotations

import shutil
import uuid
from pathlib import Path

from excel_job_offer.writer import write_job_offer_table, _prepare_row_values
from extract_job_order import extract_job_order_data

from app.config import OUTPUT_DIR
from app.services.extraction_errors import format_extraction_error


def process_job_order(
    pdf_path: Path,
    template_path: Path,
    source: str = "auto",
    *,
    start_marker: str = "",
    end_marker: str = "",
) -> tuple[dict, list[dict], Path, int]:
    """Extract completion procedure from PDF and append to the JOB ORDER sheet."""
    data = extract_job_order_data(
        pdf_path,
        source=source,
        start_marker=start_marker,
        end_marker=end_marker,
    )
    if not data.get("lines"):
        if source == "custom":
            raise ValueError(
                format_extraction_error(
                    "Job Order",
                    [
                        f"Cause: No procedure content found starting at {start_marker!r}.",
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

    suffix = template_path.suffix.lower()
    output_name = f"job_order_{uuid.uuid4().hex[:8]}{suffix}"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    _, appended_rows = write_job_offer_table(output_path, output_path, data)
    lines = [
        {
            "line_no": row["line_no"],
            "text": row["text"],
            "kind": _prepare_row_values(row["text"], data.get("source", ""))[0]
        }
        for row in data["lines"]
    ]
    return data, lines, output_path, appended_rows
