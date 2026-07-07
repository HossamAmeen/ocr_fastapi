from __future__ import annotations

import uuid
from pathlib import Path

from excel_proforma.writer import write_proforma_table
from extract_performa import extract_proforma_items

from app.config import OUTPUT_DIR


def process_proforma(pdf_path: Path, template_path: Path) -> tuple[list[dict], Path]:
    """Extract items from PDF and write them into the Excel template."""
    items = extract_proforma_items(pdf_path)
    if not items:
        raise ValueError("No Proforma line items found in the uploaded PDF.")

    for index, item in enumerate(items, start=1):
        item.setdefault("sno", index)
        item["total"] = item["per_day_rate"] * item["days"]

    output_name = f"proforma_{uuid.uuid4().hex[:8]}.xlsm"
    output_path = OUTPUT_DIR / output_name
    write_proforma_table(template_path, output_path, items)
    return items, output_path
