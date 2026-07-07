"""Extract structured data from purchase-order PDFs."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pdfplumber


@dataclass
class SubLineItem:
    item: str
    quantity: str
    unit: str
    unit_price: str
    total_price: str
    service_title: str = ""
    service_text: str = ""


@dataclass
class LineItem:
    item: str
    description: str
    quantity: str
    unit: str
    unit_price: str
    total_price: str
    vat_percent: str
    delivery_date: str
    sub_items: list[SubLineItem] = field(default_factory=list)


@dataclass
class PurchaseOrder:
    document_type: str
    vendor_name: str
    vendor_address: str
    date: str
    so_number: str
    scope_of_work: str
    line_items_count: int
    line_items: list[LineItem]
    agreement_line_item: str
    rfq_line_item: str
    currency: str
    subtotal: str
    vat: str
    total: str
    total_in_words: str
    comments: str
    raw_text: str
    tables: list[list[list[str | None]]]


def _clean(value: str | None) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _parse_header(text: str) -> dict[str, Any]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    document_type = lines[0] if lines else ""
    vendor_name = lines[1] if len(lines) > 1 else ""

    address_lines: list[str] = []
    idx = 2
    while idx < len(lines) and not lines[idx].startswith("DATE"):
        address_lines.append(lines[idx])
        idx += 1
    vendor_address = ", ".join(address_lines)

    date_match = re.search(r"DATE:\s*([\d.]+)", text)
    so_match = re.search(r"S\.O\. NUMBER\s*(\d+)", text)

    scope_match = re.search(
        r"SCOPE OF WORK\s*\n(.+?)\nTotal number of line items",
        text,
        re.DOTALL,
    )
    count_match = re.search(r"Total number of line items\s*:\s*(\d+)", text)

    agreement_match = re.search(r"4700014396/\d+", text)
    currency_match = re.search(r"CURRENCY:\s*(\w+)", text)
    subtotal_match = re.search(r"SUBTOTAL:\s*([\d,]+\.?\d*)", text)
    vat_match = re.search(r"VAT:\s*([\d,]+\.?\d*)", text)
    total_match = re.search(r"(?<!\w)TOTAL:\s*([\d,]+\.?\d*)", text)
    words_match = re.search(
        r"Total Value in words\s*:(.+?)(?:COMMENTS|$)",
        text,
        re.DOTALL,
    )
    comments_match = re.search(
        r"COMMENTS OR SPECIAL INSTRUCTION\(PO LONG TEXT\)\s*:\s*(.+)$",
        text,
        re.DOTALL,
    )

    return {
        "document_type": document_type,
        "vendor_name": vendor_name,
        "vendor_address": vendor_address,
        "date": date_match.group(1) if date_match else "",
        "so_number": so_match.group(1) if so_match else "",
        "scope_of_work": _clean(scope_match.group(1)) if scope_match else "",
        "line_items_count": int(count_match.group(1)) if count_match else 0,
        "agreement_line_item": agreement_match.group(0) if agreement_match else "",
        "rfq_line_item": "",
        "currency": currency_match.group(1) if currency_match else "",
        "subtotal": subtotal_match.group(1) if subtotal_match else "",
        "vat": vat_match.group(1) if vat_match else "",
        "total": total_match.group(1) if total_match else "",
        "total_in_words": _clean(words_match.group(1)) if words_match else "",
        "comments": _clean(comments_match.group(1)) if comments_match else "",
    }


def _row_cell(row: list[str | None], index: int) -> str:
    if index >= len(row):
        return ""
    return _clean(row[index])


def _parse_line_items_from_tables(tables: list[list[list[str | None]]]) -> list[LineItem]:
    """Parse nested PO line items from pdfplumber table output."""
    item_table = next(
        (table for table in tables if any(_row_cell(row, 0) == "Item" for row in table)),
        None,
    )
    if not item_table:
        return []

    line_items: list[LineItem] = []
    current_main: LineItem | None = None
    pending_service_title = ""
    pending_service_text = ""

    def attach_service_to_last_sub() -> None:
        nonlocal pending_service_title, pending_service_text
        if not current_main or not current_main.sub_items:
            return
        sub = current_main.sub_items[-1]
        if pending_service_title:
            sub.service_title = pending_service_title
        if pending_service_text:
            sub.service_text = pending_service_text
        pending_service_title = ""
        pending_service_text = ""

    for row in item_table:
        cells = [_clean(cell) for cell in row]
        joined = " ".join(cell for cell in cells if cell)

        if joined.startswith("Total number of line items"):
            continue
        if joined.startswith("Item Service Description"):
            continue
        if not joined or joined.startswith("Agreement / Line Item"):
            break
        if joined.startswith("PO Material Text"):
            break

        item_col = _row_cell(row, 0)
        desc_col = _row_cell(row, 1)
        qty_col = _row_cell(row, 2)
        unit_col = _row_cell(row, 3)
        unit_price_col = _row_cell(row, 5)
        total_price_col = _row_cell(row, 6)
        vat_col = _row_cell(row, 7)
        delivery_col = _row_cell(row, 8)

        if item_col.isdigit() and desc_col and qty_col and unit_col in {"AU", "EA", "LS", "DLY", "HR", "DAY"}:
            attach_service_to_last_sub()
            current_main = LineItem(
                item=item_col,
                description=desc_col,
                quantity=qty_col,
                unit=unit_col,
                unit_price=unit_price_col,
                total_price=total_price_col,
                vat_percent=vat_col,
                delivery_date=delivery_col,
            )
            line_items.append(current_main)
            continue

        if desc_col.isdigit() and qty_col and unit_col in {"AU", "EA", "LS", "DLY", "HR", "DAY"} and current_main:
            attach_service_to_last_sub()
            current_main.sub_items.append(
                SubLineItem(
                    item=desc_col,
                    quantity=qty_col,
                    unit=unit_col,
                    unit_price=unit_price_col,
                    total_price=total_price_col,
                )
            )
            continue

        if "Table" in desc_col and "|" in desc_col:
            attach_service_to_last_sub()
            pending_service_title = desc_col.split("|", 1)[-1].strip()
            continue

        if desc_col.startswith("Service Text"):
            pending_service_text = desc_col.split(":", 1)[-1].strip()
            continue

    attach_service_to_last_sub()
    return line_items


def _parse_line_items(text: str, tables: list[list[list[str | None]]]) -> list[LineItem]:
    """Parse line items from tables, with text fallback."""
    from_tables = _parse_line_items_from_tables(tables)
    if from_tables:
        return from_tables

    body_match = re.search(
        r"% Date\s*\n(.+?)\nPO Material Text",
        text,
        re.DOTALL,
    )
    if not body_match:
        return []

    body = body_match.group(1)
    lines = [line.strip() for line in body.splitlines() if line.strip()]

    line_items: list[LineItem] = []
    current_main: LineItem | None = None
    pending_service_title = ""
    pending_service_text = ""
    collecting_service_text = False

    main_pattern = re.compile(
        r"^(\d{1,3})\s+(.+?)\s+(\d+)\s+(AU|EA|LS|DLY|HR|DAY)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+(\d+)\s+([\d.]+)$"
    )
    sub_pattern = re.compile(
        r"^(\d{1,3})\s+(\d+)\s+(DLY|LS|AU|EA|HR|DAY)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)$"
    )
    service_title_pattern = re.compile(r"^Table.+?\|\s*(.+)$")

    i = 0
    while i < len(lines):
        line = lines[i]

        if line.startswith("Service Text"):
            collecting_service_text = True
            pending_service_text = re.sub(r"^Service Text\s*:\s*", "", line)
            i += 1
            continue

        if collecting_service_text:
            if main_pattern.match(line) or sub_pattern.match(line) or service_title_pattern.match(line):
                collecting_service_text = False
            else:
                pending_service_text = f"{pending_service_text} {_clean(line)}".strip()
                i += 1
                continue

        main_match = main_pattern.match(line)
        if main_match:
            if current_main and pending_service_title and current_main.sub_items:
                current_main.sub_items[-1].service_title = pending_service_title
                current_main.sub_items[-1].service_text = pending_service_text
            current_main = LineItem(
                item=main_match.group(1),
                description=main_match.group(2),
                quantity=main_match.group(3),
                unit=main_match.group(4),
                unit_price=main_match.group(5),
                total_price=main_match.group(6),
                vat_percent=main_match.group(7),
                delivery_date=main_match.group(8),
            )
            line_items.append(current_main)
            pending_service_title = ""
            pending_service_text = ""
            i += 1
            continue

        sub_match = sub_pattern.match(line)
        if sub_match and current_main:
            if pending_service_title and current_main.sub_items:
                current_main.sub_items[-1].service_title = pending_service_title
                current_main.sub_items[-1].service_text = pending_service_text
            current_main.sub_items.append(
                SubLineItem(
                    item=sub_match.group(1),
                    quantity=sub_match.group(2),
                    unit=sub_match.group(3),
                    unit_price=sub_match.group(4),
                    total_price=sub_match.group(5),
                )
            )
            pending_service_title = ""
            pending_service_text = ""
            i += 1
            continue

        title_match = service_title_pattern.match(line)
        if title_match:
            if pending_service_title and current_main and current_main.sub_items:
                current_main.sub_items[-1].service_title = pending_service_title
                current_main.sub_items[-1].service_text = pending_service_text
            pending_service_title = _clean(title_match.group(1))
            pending_service_text = ""
            i += 1
            continue

        i += 1

    if current_main and pending_service_title and current_main.sub_items:
        current_main.sub_items[-1].service_title = pending_service_title
        current_main.sub_items[-1].service_text = pending_service_text

    return line_items


def _extract_tables(pdf_path: Path) -> list[list[list[str | None]]]:
    tables: list[list[list[str | None]]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                cleaned = [[_clean(cell) or None for cell in row] for row in table]
                tables.append(cleaned)
    return tables


def _extract_raw_text(pdf_path: Path) -> str:
    parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts).strip()


def extract_purchase_order(pdf_path: str | Path) -> PurchaseOrder:
    path = Path(pdf_path)
    raw_text = _extract_raw_text(path)
    tables = _extract_tables(path)
    header = _parse_header(raw_text)
    line_items = _parse_line_items(raw_text, tables)

    return PurchaseOrder(
        raw_text=raw_text,
        tables=tables,
        line_items=line_items,
        **header,
    )


def format_tables(tables: list[list[list[str | None]]]) -> str:
    """Render extracted PDF tables as readable terminal output."""
    sections: list[str] = []

    for index, table in enumerate(tables, start=1):
        if not table:
            continue

        col_count = max(len(row) for row in table)
        normalized: list[list[str]] = []
        widths = [0] * col_count

        for row in table:
            cells = [_clean(cell) for cell in row]
            while len(cells) < col_count:
                cells.append("")
            normalized.append(cells)
            for col, cell in enumerate(cells):
                widths[col] = max(widths[col], len(cell))

        divider = "-+-".join("-" * width for width in widths)
        lines = [f"Table {index}", divider]

        for row_index, cells in enumerate(normalized):
            lines.append(" | ".join(cell.ljust(widths[col]) for col, cell in enumerate(cells)))
            if row_index == 0:
                lines.append(divider)

        sections.append("\n".join(lines))

    return "\n\n".join(sections) if sections else "No tables found."


def format_purchase_order(po: PurchaseOrder) -> str:
    """Render extracted data as human-readable terminal output."""
    lines = [
        "=" * 72,
        "PURCHASE ORDER EXTRACTION",
        "=" * 72,
        f"Document Type : {po.document_type}",
        f"Vendor        : {po.vendor_name}",
        f"Address       : {po.vendor_address}",
        f"Date          : {po.date}",
        f"S.O. Number   : {po.so_number}",
        f"Scope of Work : {po.scope_of_work}",
        f"Line Items    : {po.line_items_count}",
        "",
        "-" * 72,
        "LINE ITEMS",
        "-" * 72,
    ]

    for item in po.line_items:
        lines.extend(
            [
                f"Item {item.item}: {item.description}",
                f"  Qty: {item.quantity} {item.unit} | Unit Price: {item.unit_price} | "
                f"Total: {item.total_price} | VAT: {item.vat_percent}% | "
                f"Delivery: {item.delivery_date}",
            ]
        )
        for sub in item.sub_items:
            lines.append(
                f"  - Sub-item {sub.item}: {sub.quantity} {sub.unit} @ "
                f"{sub.unit_price} = {sub.total_price}"
            )
            if sub.service_title:
                lines.append(f"    Service: {sub.service_title}")
            if sub.service_text:
                lines.append(f"    Text: {sub.service_text}")

    lines.extend(
        [
            "",
            "-" * 72,
            "FINANCIAL SUMMARY",
            "-" * 72,
            f"Agreement / Line Item : {po.agreement_line_item}",
            f"RFQ / Line Item       : {po.rfq_line_item or '(empty)'}",
            f"Currency              : {po.currency}",
            f"Subtotal              : {po.subtotal}",
            f"VAT                   : {po.vat}",
            f"Total                 : {po.total}",
            f"Total in Words        : {po.total_in_words}",
            "",
            "-" * 72,
            "COMMENTS",
            "-" * 72,
            po.comments or "(none)",
            "",
            "-" * 72,
            "RAW TEXT",
            "-" * 72,
            po.raw_text,
            "",
            "-" * 72,
            "TABLES (JSON)",
            "-" * 72,
            json.dumps(po.tables, indent=2, ensure_ascii=False),
            "",
            "-" * 72,
            "STRUCTURED DATA (JSON)",
            "-" * 72,
            json.dumps(asdict(po), indent=2, ensure_ascii=False),
        ]
    )
    return "\n".join(lines)
