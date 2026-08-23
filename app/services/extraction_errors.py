from __future__ import annotations


def parse_form_bool(value: object) -> bool:
    """Parse checkbox / radio booleans sent as multipart form strings."""
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().casefold() in {"1", "true", "yes", "on"}


def format_extraction_error(
    section: str,
    lines: list[str],
    *,
    mode: str | None = None,
) -> str:
    """Build a user-facing extraction error with a clear section header."""
    header = f"{section} extraction failed"
    if mode:
        header += f" ({mode})"
    return header + ".\n\n" + "\n".join(lines)


def soe_extraction_mode_label(is_summary: bool) -> str:
    return "Summary (paragraph) mode" if is_summary else "Table mode"
