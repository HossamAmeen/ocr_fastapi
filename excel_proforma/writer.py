"""Read data.xlsm and rewrite the Proforma table from Python data."""

from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

PROFORMA_SHEET = "Proforma"
HEADER_ROW = 7
DATA_START_ROW = 8
GROSS_VALUE_ROW = 15
PROTECTED_HEADER_END_ROW = 7

# Header metadata that must never be overwritten (e.g. DT label + value).
PROTECTED_CELLS = ("D3", "E3", "D4", "E4", "D5", "E5", "D6", "E6")


def write_proforma_table(
    input_path: str | Path,
    output_path: str | Path,
    items: list[dict[str, Any]],
) -> Path:
    """Load xlsm, replace Proforma line items, and save a new workbook."""
    input_path = Path(input_path)
    output_path = Path(output_path)

    wb = load_workbook(input_path, keep_vba=True)
    ws = wb[PROFORMA_SHEET]

    protected = _snapshot_cells(ws, PROTECTED_CELLS)

    _clear_old_rows(ws, start_row=DATA_START_ROW, keep_rows=len(items))
    _write_items(ws, items)
    _update_totals(ws, item_count=len(items))
    _restore_cells(ws, protected)

    wb.save(output_path)
    return output_path


def _snapshot_cells(ws: Worksheet, addresses: tuple[str, ...]) -> dict[str, object]:
    return {addr: ws[addr].value for addr in addresses}


def _restore_cells(ws: Worksheet, values: dict[str, object]) -> None:
    for addr, value in values.items():
        cell = ws[addr]
        if isinstance(cell, MergedCell):
            continue
        cell.value = value


def _clear_old_rows(ws: Worksheet, start_row: int, keep_rows: int) -> None:
    """Clear unused data rows before the totals section (never touch header rows)."""
    if start_row <= PROTECTED_HEADER_END_ROW:
        raise ValueError("Cannot clear protected header rows.")

    for row in range(start_row + keep_rows, GROSS_VALUE_ROW):
        for col in ("A", "B", "D", "E", "F"):
            cell = ws[f"{col}{row}"]
            if isinstance(cell, MergedCell):
                continue
            cell.value = None


def _write_items(ws: Worksheet, items: list[dict[str, Any]]) -> None:
    template_row = DATA_START_ROW

    for index, item in enumerate(items):
        row = DATA_START_ROW + index
        if row >= GROSS_VALUE_ROW:
            break

        _copy_row_style(ws, template_row, row)

        ws[f"A{row}"].value = item.get("sno", index + 1)
        ws[f"B{row}"].value = item["description"]
        ws[f"D{row}"].value = float(item["per_day_rate"])
        ws[f"E{row}"].value = float(item["days"])
        ws[f"F{row}"].value = f"=D{row}*E{row}"


def _copy_row_style(ws: Worksheet, source_row: int, target_row: int) -> None:
    for col in ("A", "B", "D", "E", "F"):
        source = ws[f"{col}{source_row}"]
        target = ws[f"{col}{target_row}"]
        target.number_format = copy(source.number_format)
        target.font = copy(source.font)
        target.border = copy(source.border)
        target.fill = copy(source.fill)
        target.alignment = copy(source.alignment)


def _update_totals(ws: Worksheet, item_count: int) -> None:
    if item_count == 0:
        ws[f"F{GROSS_VALUE_ROW}"].value = 0
        return

    first_row = DATA_START_ROW
    last_row = DATA_START_ROW + item_count - 1
    ws[f"F{GROSS_VALUE_ROW}"].value = f"=SUM(F{first_row}:F{last_row})"
