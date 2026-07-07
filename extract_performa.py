"""Extract Proforma line items from a purchase-order PDF."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pdf_extractor.extractor import extract_purchase_order


def _clean(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _row_cell(row: list[str | None], index: int) -> str:
    if index >= len(row):
        return ""
    return _clean(row[index])


def _is_number(value: object) -> bool:
    if value is None:
        return False
    text = str(value).strip().replace(",", "")
    if not text:
        return False
    try:
        float(text)
        return True
    except ValueError:
        return False


def _to_float(value: object) -> float:
    return float(str(value).replace(",", ""))


def extract_proforma_items(pdf_path: str | Path) -> list[dict[str, Any]]:
    """Build Proforma table rows from a purchase-order PDF."""
    po = extract_purchase_order(pdf_path)
    items: list[dict[str, Any]] = []

    item_table = next(
        (
            table
            for table in po.tables
            if len(table) >= 5 and any(_row_cell(row, 0) == "Item" for row in table)
        ),
        None,
    )
    if not item_table:
        return items

    pending: dict[str, float] | None = None

    for row in item_table:
        desc_col = _row_cell(row, 1)
        qty_col = _row_cell(row, 2)
        unit_col = _row_cell(row, 3)
        unit_price_col = _row_cell(row, 5)

        if _is_number(desc_col) and qty_col and unit_col:
            pending = {
                "days": _to_float(qty_col),
                "per_day_rate": _to_float(unit_price_col),
            }
            continue

        if pending and "Table" in desc_col and "|" in desc_col:
            items.append(
                {
                    "sno": len(items) + 1,
                    "description": desc_col,
                    "per_day_rate": pending["per_day_rate"],
                    "days": pending["days"],
                    "total": pending["per_day_rate"] * pending["days"],
                }
            )
            pending = None

    return items
