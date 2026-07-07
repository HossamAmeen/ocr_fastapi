from __future__ import annotations

import shutil
import uuid
from pathlib import Path

from excel_job_offer.writer import write_job_offer_table
from extract_job_over import extract_job_over_data

from app.config import OUTPUT_DIR


def process_job_over(
    pdf_path: Path,
    template_path: Path,
    source: str = "auto",
) -> tuple[dict, list[dict], Path, int]:
    """Extract completion procedure from PDF and append to the JOB ORDER sheet."""
    data = extract_job_over_data(pdf_path, source=source)
    if not data.get("lines"):
        raise ValueError("No completion procedure content found in the uploaded PDF.")

    suffix = template_path.suffix.lower()
    output_name = f"job_over_{uuid.uuid4().hex[:8]}{suffix}"
    output_path = OUTPUT_DIR / output_name
    shutil.copy2(template_path, output_path)

    _, appended_rows = write_job_offer_table(output_path, output_path, data)
    lines = [
        {"line_no": row["line_no"], "text": row["text"]}
        for row in data["lines"]
    ]
    return data, lines, output_path, appended_rows
