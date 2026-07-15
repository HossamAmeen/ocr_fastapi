"""Extract Time Log data from SOE daily operations report PDFs."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pdfplumber

from pdf_extractor.extractor import _extract_raw_text

_TIME_LOG_ROW = re.compile(
    # Phase may be "WO", "36", or multi-token values like "12 1/4".
    r"^(\d{1,2}:\d{2})\s+([\d.]+)\s+(.+?)\s+(\w+)\s+(\w+)\s+(\w)\s+(.+)$"
)
_TIME_LOG_ROW_FROM_TO = re.compile(
    r"^(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})\s+([\d.]+)\s+(.+?)\s+(\w+)\s+(\w+)\s+(\w)\s+(.+)$"
)
_NOTE_ROW = re.compile(r"^(?:[A-Z]\s+)?\*+\s*(.+)$")
_OAMN_RANGE = re.compile(
    r"^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*(?:Hr'?s?)?\s*:?\s*(.*)$",
    re.IGNORECASE,
)
_OAMN_HEADER = re.compile(r"^O\.A\.M\.N\.?$", re.IGNORECASE)
_REPORT_PERIOD = re.compile(
    r"Rpt\. Period:\s*(.+?)\s+to\s+(.+?)(?:\n|$)",
    re.IGNORECASE,
)
_WELL_NAME = re.compile(r"Well Name:\s*(\S+)")
_REPORT_DATE = re.compile(r"DATE:\s*(.+?)(?:\n|$)", re.IGNORECASE)
_REPORT_FOR = re.compile(r"Report For:\s*(\d{1,2})/(\d{1,2})/(\d{4})", re.IGNORECASE)
# Classic daily ops: "Rig: AL HUDAIRIYAT Cost ..."
_RIG_LABEL = re.compile(r"(?im)^Rig:\s*(.+)$")
# ADNOC / inline headers: "RIG: AD-109 DOM: ASR ..."
_RIG_INLINE = re.compile(r"\bRIG\s*:?\s*([A-Z0-9][A-Z0-9\-]{1,20})\b")
# Vendor reports sometimes use "Rig Number" / "Rig Number: 125"
_RIG_NUMBER = re.compile(r"(?i)\bRig\s*Number\b\s*:?\s*([A-Za-z0-9\-]+)")
_FROM_TIME = re.compile(r"^\d{1,2}:\d{2}$")
_CLOCK_TIME = re.compile(r"^\d{1,2}:\d{2}$")
_DURATION_VALUE = re.compile(r"^\d+\.\d{1,2}$")
# Substring match for Job Time Log word windows, etc.
_TIME_LOG_HEADING = re.compile(r"time[\s\-]+log", re.IGNORECASE)
# Standalone title lines only: "Time Log", "Job Time Log" — not "Ahead/Behind Time log(Hrs)".
# Use [^\S\n] so titles cannot span lines (e.g. "/ /" above "Time Log").
_TIME_LOG_TITLE_LINE = re.compile(
    r"^(?P<title>(?:[\w/&\-]+[^\S\n]+)*time[ \t\-]+log)[^\S\n]*$",
    re.IGNORECASE | re.MULTILINE,
)
_TIME_LOG_HEADER_ROW = re.compile(
    r"to\s+duration\s+phase\s+code\s+sub\s+type\s+operation",
    re.IGNORECASE,
)
_TIME_LOG_HEADER_FROM = re.compile(
    r"from\s+duration\s+phase\s+code\s+sub\s+type\s+operation",
    re.IGNORECASE,
)
_TIME_LOG_HEADER_ANY = re.compile(
    r"(?:from|to|start)\s+(?:to\s+)?duration\s+phase\s+code\s+sub\s+type\s+operation",
    re.IGNORECASE,
)
_JOB_TIME_LOG_MARKER = re.compile(
    r"Start\s+End\s+Sum\s+of|Start\s*Time.*End\s*Time.*Dur",
    re.IGNORECASE | re.DOTALL,
)
_JOB_TIME_LOG_SECTION_END = re.compile(
    r"^(Total|Interval Problems|Mud Checks|Drill Strings)\b",
    re.IGNORECASE,
)
_PAGE_CHROME_LINE = re.compile(
    r"^(?:"
    r"Drilling Daily Operations Report|"
    r"Well Name:.*|"
    r"Well Bore:.*|"
    r"Rig:.*|"
    r"Report No:.*|"
    r"Phase:.*Rpt\.?\s*Period:.*|"
    r"\d{1,2}/\d{1,2}/\d{4}\s+\d+\s*"
    r")$",
    re.IGNORECASE,
)
_SECTION_END = re.compile(r"Phase Time and Cost Summary", re.IGNORECASE)
# Comment column content starts left of the "Comment" header label.
_JOB_COMMENT_X_MIN = 355.0
_JOB_START_TIME_X_MAX = 45.0
_JOB_END_TIME_X_MAX = 75.0
_JOB_DURATION_X_MAX = 105.0


def _to_float(value: str) -> float:
    return float(value.replace(",", ""))


def _finalize_entry(entry: dict[str, Any] | None) -> dict[str, Any] | None:
    if not entry:
        return None
    entry["operation"] = entry["operation"].strip()
    entry["notes"] = [note.strip() for note in entry["notes"] if note.strip()]
    return entry


def _is_time_log_header_line(line: str) -> bool:
    return bool(
        _TIME_LOG_HEADER_ROW.search(line)
        or _TIME_LOG_HEADER_FROM.search(line)
        or _TIME_LOG_HEADER_ANY.search(line)
    )


def _detect_time_log_time_column(header_line: str) -> str:
    """Return which header column supplies the Excel time value.

    Uses ``to`` when a To column exists, otherwise the first time column
    (From / Start).
    """
    header = header_line.casefold()
    has_to = bool(re.search(r"\bto\b", header))
    has_from = bool(re.search(r"\bfrom\b", header)) or bool(re.search(r"\bstart\b", header))
    if has_to:
        return "to"
    if has_from:
        return "from"
    return "first"


def _build_time_log_entry(
    *,
    time_column: str,
    from_time: str = "",
    to_time: str = "",
    duration: float,
    phase: str,
    code: str,
    sub: str,
    entry_type: str,
    operation: str,
    notes: list[str],
) -> dict[str, Any]:
    """Build a time-log entry, picking the Excel time from To or first column."""
    if time_column == "to":
        excel_time = to_time or from_time
    else:
        excel_time = from_time or to_time
    return {
        "from": from_time,
        "to": excel_time,
        "duration": duration,
        "phase": phase,
        "code": code,
        "sub": sub,
        "type": entry_type,
        "operation": operation,
        "notes": notes,
    }


def _parse_time_log_row(line: str, time_column: str) -> dict[str, Any] | None:
    from_to_match = _TIME_LOG_ROW_FROM_TO.match(line)
    if from_to_match:
        from_time = from_to_match.group(1)
        to_time = from_to_match.group(2)
        return _build_time_log_entry(
            time_column=time_column,
            from_time=from_time,
            to_time=to_time,
            duration=_to_float(from_to_match.group(3)),
            phase=from_to_match.group(4),
            code=from_to_match.group(5),
            sub=from_to_match.group(6),
            entry_type=from_to_match.group(7),
            operation=from_to_match.group(8).strip(),
            notes=[],
        )

    row_match = _TIME_LOG_ROW.match(line)
    if not row_match:
        return None

    time_value = row_match.group(1)
    if time_column == "to":
        from_time, to_time = "", time_value
    else:
        from_time, to_time = time_value, ""
    return _build_time_log_entry(
        time_column=time_column,
        from_time=from_time,
        to_time=to_time,
        duration=_to_float(row_match.group(2)),
        phase=row_match.group(3),
        code=row_match.group(4),
        sub=row_match.group(5),
        entry_type=row_match.group(6),
        operation=row_match.group(7).strip(),
        notes=[],
    )


_DEFAULT_TABLE_NAMES = ("Time Log", "Job Time Log")
_OPERATIONAL_TIME_SUMMARY = "Operational Time Summary"


def _normalize_table_names(table_names: list[str] | None) -> list[str]:
    if not table_names:
        return list(_DEFAULT_TABLE_NAMES)
    cleaned = [name.strip() for name in table_names if name and name.strip()]
    return cleaned or list(_DEFAULT_TABLE_NAMES)


def _compile_table_title_pattern(name: str) -> re.Pattern[str]:
    """Build a regex that matches a standalone PDF table title line."""
    tokens = [re.escape(part) for part in name.split() if part]
    if not tokens:
        return re.compile(r"a^")
    body = r"[^\S\n]+".join(tokens)
    return re.compile(rf"^{body}[^\S\n]*$", re.IGNORECASE | re.MULTILINE)


def _table_title_patterns(table_names: list[str] | None) -> list[re.Pattern[str]]:
    return [_compile_table_title_pattern(name) for name in _normalize_table_names(table_names)]


def _line_matches_table_title(line: str, patterns: list[re.Pattern[str]]) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    return any(pattern.match(stripped) for pattern in patterns)


def _is_time_log_title_line(line: str, table_names: list[str] | None = None) -> bool:
    patterns = _table_title_patterns(table_names)
    if _line_matches_table_title(line, patterns):
        return True
    if not table_names:
        return bool(re.match(
            r"^(?:[\w/&\-]+[^\S\n]+)*time[ \t\-]+log[^\S\n]*$",
            line.strip(),
            re.IGNORECASE,
        ))
    return False


def _find_time_log_start(text: str, table_names: list[str] | None = None) -> int:
    """Return the start index of the first matching table title, or -1."""
    patterns = _table_title_patterns(table_names)
    starts = [match.start() for pattern in patterns for match in pattern.finditer(text)]
    if starts:
        return min(starts)
    if not table_names:
        match = _TIME_LOG_TITLE_LINE.search(text)
        return match.start() if match else -1
    return -1


def _is_time_log_heading(line: str, table_names: list[str] | None = None) -> bool:
    """True for title-like headings (used by Job Time Log helpers)."""
    return _is_time_log_title_line(line, table_names) or bool(
        _TIME_LOG_HEADING.search(line) and len(line.strip()) <= 40
    )


def _is_job_time_log_section(section: str) -> bool:
    """True when the time-log block uses Start/End/Comment columns."""
    return bool(_JOB_TIME_LOG_MARKER.search(section[:500]))


def _is_skippable_time_log_line(line: str, table_names: list[str] | None = None) -> bool:
    """Skip reprinted titles, column headers, and page chrome across page breaks."""
    if _is_time_log_title_line(line, table_names):
        return True
    if _is_time_log_header_line(line):
        return True
    if _PAGE_CHROME_LINE.match(line):
        return True
    return False


def _iter_classic_time_log_sections(
    text: str,
    table_names: list[str] | None = None,
) -> list[tuple[str, int]]:
    """Yield (section_text, start_index) for each matching table through its summary.

    Continuation pages that reprint the table header stay inside the same
    section until the summary marker appears.
    """
    patterns = _table_title_patterns(table_names)
    title_matches = sorted(
        (
            (match.start(), match)
            for pattern in patterns
            for match in pattern.finditer(text)
        ),
        key=lambda item: item[0],
    )
    if not title_matches and not table_names:
        title_matches = [(match.start(), match) for match in _TIME_LOG_TITLE_LINE.finditer(text)]

    sections: list[tuple[str, int]] = []
    index = 0
    while index < len(title_matches):
        start = title_matches[index][0]
        end_match = _SECTION_END.search(text, start)
        end = end_match.start() if end_match else len(text)
        sections.append((text[start:end].strip(), start))
        while index < len(title_matches) and title_matches[index][0] < end:
            index += 1
    return sections


def _clean_rig_value(raw: str) -> str:
    """Normalize a Rig: line value down to the rig name/code."""
    text = raw.strip()
    if not text:
        return ""
    # Stop before other header fields that often share the same PDF line.
    text = re.split(
        r"\s{2,}"
        r"|\s+Cost\b"
        r"|\s+Number\b"
        r"|\s+Rotating\b"
        r"|\s+DOM\s*:"
        r"|\s+WELL\s*:"
        r"|\s+Well\s*:"
        r"|\s+EVENT\s*:"
        r"|\s+Event\s*:"
        r"|\s+RPT\b"
        r"|\s+Rpt\b"
        r"|\s+DATE\s*:"
        r"|\s+Date\s*:"
        r"|\s+Report\b"
        r"|\s+\d{1,3}(?:,\d{3})+(?:\.\d+)?",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()
    return text


def format_rig_for_display(value: str) -> str:
    """Return a short, human-readable rig label for UI and error messages."""
    cleaned = _clean_rig_value(value)
    return cleaned or str(value or "").strip() or "unknown"


def _normalize_rig(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.casefold())


def rigs_match(left: str, right: str) -> bool:
    """True when two rig labels refer to the same rig."""
    a = _normalize_rig(left)
    b = _normalize_rig(right)
    if not a or not b:
        return False
    return a == b or a in b or b in a


def _extract_rig_from_text(text: str) -> str:
    """Return the Rig value from page or section text."""
    rig = ""
    for match in _RIG_LABEL.finditer(text):
        cleaned = _clean_rig_value(match.group(1))
        if cleaned and not cleaned.casefold().startswith("release"):
            rig = cleaned
    if rig:
        return rig
    inline_match = _RIG_INLINE.search(text)
    if inline_match:
        cleaned = _clean_rig_value(inline_match.group(1))
        if cleaned:
            return cleaned
    number_match = re.search(r"(?is)\bRig\s*Number\b.{0,80}?\b(\d{2,5})\b", text)
    if number_match:
        return number_match.group(1)
    return ""


def _collect_rigs_from_pages(pages: list[str]) -> list[str]:
    """Return unique rig codes found across PDF pages, in order."""
    rigs: list[str] = []
    for page_text in pages:
        display = format_rig_for_display(_extract_rig_from_text(page_text))
        if display != "unknown" and display not in rigs:
            rigs.append(display)
    return rigs


def _has_classic_rig_label(text: str) -> bool:
    """True when the PDF uses a classic 'Rig:' header field."""
    for match in _RIG_LABEL.finditer(text):
        cleaned = _clean_rig_value(match.group(1))
        if cleaned and not cleaned.casefold().startswith("release"):
            return True
    return False


def _effective_rig_page_filter(pages: list[str], rig_filter: str | None) -> str | None:
    """Apply page-level Rig filtering only when the PDF has classic Rig: headers."""
    if not rig_filter:
        return None
    if any(_has_classic_rig_label(page) for page in pages):
        return rig_filter
    return None


def _iter_pdf_page_texts(pdf_path: Path) -> list[str]:
    """Return raw text for every PDF page, in order."""
    pages: list[str] = []
    with pdfplumber.open(pdf_path) as document:
        for page in document.pages:
            pages.append(page.extract_text() or "")
    return pages


def _filter_pages_by_rig(pages: list[str], rig_filter: str | None) -> tuple[str, list[int]]:
    """Keep every page whose Rig matches the Excel Rig value.

    Unlike taking only the first matching Time Log section, this keeps all
    matching pages (for example pages 6-9 for one Rig in a multi-well PDF).
    """
    if not rig_filter:
        return "\n".join(pages).strip(), list(range(len(pages)))

    matched_pages: list[str] = []
    matched_indexes: list[int] = []
    for index, page_text in enumerate(pages):
        page_rig = _extract_rig_from_text(page_text)
        if page_rig and rigs_match(page_rig, rig_filter):
            matched_pages.append(page_text)
            matched_indexes.append(index)
    return "\n".join(matched_pages).strip(), matched_indexes


def _metadata_before(text: str, position: int) -> tuple[str, str, str, str]:
    """Read well, report period, and rig from text preceding a Time Log title."""
    window = text[max(0, position - 6000) : position]
    well_match = None
    for match in _WELL_NAME.finditer(window):
        well_match = match
    period_match = None
    for match in _REPORT_PERIOD.finditer(window):
        period_match = match
    well = well_match.group(1) if well_match else ""
    period_from = period_match.group(1).strip() if period_match else ""
    period_to = period_match.group(2).strip() if period_match else ""
    rig = _extract_rig_from_text(window)
    return well, period_from, period_to, rig


def _report_for_to_dmy(text: str) -> str:
    """Parse 'Report For: M/D/Y' and return D/M/Y for the Excel writer."""
    match = _REPORT_FOR.search(text)
    if not match:
        return ""
    month, day, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
    return f"{day}/{month}/{year}"


def _group_words_into_lines(
    words: list[dict[str, Any]],
    y_tol: float = 3.5,
) -> list[list[dict[str, Any]]]:
    if not words:
        return []
    ordered = sorted(words, key=lambda word: (float(word["top"]), float(word["x0"])))
    lines: list[list[dict[str, Any]]] = [[ordered[0]]]
    for word in ordered[1:]:
        if abs(float(word["top"]) - float(lines[-1][0]["top"])) <= y_tol:
            lines[-1].append(word)
        else:
            lines.append([word])
    return [sorted(line, key=lambda word: float(word["x0"])) for line in lines]


def _split_comment_parts(parts: list[str]) -> tuple[list[str], list[str]]:
    """Split comment words into operation text and *-prefixed notes."""
    operation_parts: list[str] = []
    notes: list[str] = []
    index = 0
    while index < len(parts):
        part = parts[index]
        if part == "*" or part.startswith("*"):
            note_bits = []
            if part.startswith("*") and part != "*":
                note_bits.append(part.lstrip("*").strip())
            index += 1
            while index < len(parts) and parts[index] != "*" and not parts[index].startswith("*"):
                note_bits.append(parts[index])
                index += 1
            note = " ".join(bit for bit in note_bits if bit).strip()
            if note:
                notes.append(note)
            continue
        operation_parts.append(part)
        index += 1
    return operation_parts, notes


def _parse_job_time_log_page(page: Any) -> list[dict[str, Any]]:
    """Extract Job Time Log rows from one PDF page using word positions."""
    words = page.extract_words() or []
    if not words:
        return []

    heading_top: float | None = None
    for index in range(len(words) - 1):
        window = " ".join(word["text"] for word in words[index : index + 3])
        if not _TIME_LOG_HEADING.search(window):
            continue
        heading_top = min(float(word["top"]) for word in words[index : index + 3])
        break
    if heading_top is None:
        return []

    section_words = [word for word in words if float(word["top"]) >= heading_top - 1]
    lines = _group_words_into_lines(section_words)
    entries: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    for line in lines:
        texts = [word["text"] for word in line]
        joined = " ".join(texts)
        if current is not None and _JOB_TIME_LOG_SECTION_END.match(joined):
            break
        if _is_time_log_heading(joined) or "Dur (hr)" in joined or joined.startswith("Start End"):
            continue
        # Skip the split header lines: "Start End Sum of ..." / "Time Time Dur ..."
        if "IADC" in texts and "Comment" in texts:
            continue
        if texts[:2] == ["Start", "End"] or texts[:2] == ["Time", "Time"]:
            continue

        start_word = next(
            (
                word
                for word in line
                if _CLOCK_TIME.match(word["text"]) and float(word["x0"]) < _JOB_START_TIME_X_MAX
            ),
            None,
        )
        if start_word is not None:
            if current is not None:
                entries.append(_finalize_entry(current) or current)
            end_word = next(
                (
                    word
                    for word in line
                    if _CLOCK_TIME.match(word["text"])
                    and _JOB_START_TIME_X_MAX <= float(word["x0"]) < _JOB_END_TIME_X_MAX
                ),
                None,
            )
            duration_word = next(
                (
                    word
                    for word in line
                    if _DURATION_VALUE.match(word["text"])
                    and _JOB_END_TIME_X_MAX <= float(word["x0"]) < _JOB_DURATION_X_MAX
                ),
                None,
            )
            comment_parts = [
                word["text"] for word in line if float(word["x0"]) >= _JOB_COMMENT_X_MIN
            ]
            operation_parts, notes = _split_comment_parts(comment_parts)
            current = {
                "from": start_word["text"],
                "to": end_word["text"] if end_word else start_word["text"],
                "duration": _to_float(duration_word["text"]) if duration_word else 0.0,
                "phase": "",
                "code": "",
                "sub": "",
                "type": "",
                "operation": " ".join(operation_parts).strip(),
                "notes": notes,
            }
            continue

        if current is None:
            continue

        comment_parts = [
            word["text"] for word in line if float(word["x0"]) >= _JOB_COMMENT_X_MIN
        ]
        if not comment_parts:
            continue

        operation_parts, notes = _split_comment_parts(comment_parts)
        current["notes"].extend(notes)
        if operation_parts:
            addition = " ".join(operation_parts).strip()
            if current["operation"]:
                current["operation"] = f"{current['operation']} {addition}".strip()
            else:
                current["operation"] = addition

    if current is not None:
        entries.append(_finalize_entry(current) or current)
    return entries


def extract_job_time_log(
    pdf_path: str | Path,
    *,
    rig_filter: str | None = None,
    table_names: list[str] | None = None,
) -> dict[str, Any]:
    """Extract Start/End/Comment Job Time Log tables from vendor drilling PDFs."""
    pdf_path = Path(pdf_path)
    pages = _iter_pdf_page_texts(pdf_path)
    effective_filter = _effective_rig_page_filter(pages, rig_filter)
    _, matched_page_indexes = _filter_pages_by_rig(pages, effective_filter)

    if effective_filter and not matched_page_indexes:
        full_text = "\n".join(pages)
        return {
            "well_name": _WELL_NAME.search(full_text).group(1) if _WELL_NAME.search(full_text) else "",
            "rig": _extract_rig_from_text(full_text),
            "report_period_from": "",
            "report_period_to": "",
            "entries": [],
            "oamn_entries": [],
            "total_duration": 0.0,
            "skipped_rig_mismatch": True,
            "skip_reason": "rig_mismatch",
            "matched_pages": [],
        }

    page_indexes = matched_page_indexes or list(range(len(pages)))
    text = "\n".join(pages[index] for index in page_indexes)
    rig = _extract_rig_from_text(text)

    entries: list[dict[str, Any]] = []
    with pdfplumber.open(pdf_path) as document:
        for index in page_indexes:
            if index < len(document.pages):
                entries.extend(_parse_job_time_log_page(document.pages[index]))

    report_date = _report_for_to_dmy(text)
    well_match = _WELL_NAME.search(text)
    skip_reason = "empty_table" if not entries else ""
    return {
        "well_name": well_match.group(1) if well_match else "",
        "rig": rig,
        "report_period_from": report_date,
        "report_period_to": report_date,
        "entries": entries,
        "oamn_entries": [],
        "total_duration": round(sum(float(entry.get("duration", 0.0)) for entry in entries), 2),
        "skip_reason": skip_reason,
        "matched_pages": [index + 1 for index in page_indexes],
    }



def _parse_time_log_section(
    section: str,
    table_names: list[str] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    lines = [line.strip() for line in section.splitlines() if line.strip()]
    time_column = "to"
    if lines and _is_time_log_title_line(lines[0], table_names):
        # Skip title; also skip the column header row when present.
        if len(lines) >= 2 and _is_time_log_header_line(lines[1]):
            time_column = _detect_time_log_time_column(lines[1])
            lines = lines[2:]
        else:
            lines = lines[1:]
    elif lines and _is_time_log_header_line(lines[0]):
        time_column = _detect_time_log_time_column(lines[0])
        lines = lines[1:]

    entries: list[dict[str, Any]] = []
    oamn_entries: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    pending_notes: list[str] = []
    in_oamn = False
    pending_oamn: dict[str, Any] | None = None

    for line in lines:
        if _is_skippable_time_log_line(line, table_names):
            continue

        if _OAMN_HEADER.match(line):
            finalized = _finalize_entry(current)
            if finalized:
                entries.append(finalized)
            current = None
            pending_notes = []
            in_oamn = True
            continue
        if line.startswith("="):
            continue

        parsed_row = _parse_time_log_row(line, time_column)
        if parsed_row:
            finalized = _finalize_entry(current)
            if finalized:
                entries.append(finalized)
            parsed_row["notes"] = pending_notes.copy()
            current = parsed_row
            pending_notes = []
            continue

        note_match = _NOTE_ROW.match(line)
        if note_match:
            note = note_match.group(1).strip()
            if in_oamn:
                if pending_oamn is not None:
                    pending_oamn["notes"].append(note)
            elif current is not None:
                current["notes"].append(note)
            else:
                pending_notes.append(note)
            continue

        oamn_match = _OAMN_RANGE.match(line)
        if oamn_match and in_oamn:
            if pending_oamn:
                oamn_entries.append(pending_oamn)
            pending_oamn = {
                "from": oamn_match.group(1),
                "to": oamn_match.group(2),
                "operation": (oamn_match.group(3) or "").strip(),
                "notes": [],
            }
            continue

        if in_oamn and pending_oamn is not None:
            if pending_oamn["operation"]:
                pending_oamn["operation"] = f"{pending_oamn['operation']} {line}".strip()
            else:
                pending_oamn["operation"] = line
            continue

        if current is not None:
            current["operation"] = f"{current['operation']} {line}".strip()

    finalized = _finalize_entry(current)
    if finalized:
        entries.append(finalized)
    if pending_oamn:
        oamn_entries.append(pending_oamn)

    return entries, oamn_entries


def extract_time_log(
    pdf_path: str | Path,
    *,
    rig_filter: str | None = None,
    table_names: list[str] | None = None,
) -> dict[str, Any]:
    """Extract time-log table(s) from an SOE daily operations report PDF.

    When ``rig_filter`` is set (from the Excel Rig cell), every page whose ``Rig:``
    value matches is kept — not only the first page/section. Time Log tables
    on those pages are then extracted, including multi-page continuations.
    """
    pdf_path = Path(pdf_path)
    pages = _iter_pdf_page_texts(pdf_path)
    effective_filter = _effective_rig_page_filter(pages, rig_filter)
    text, matched_page_indexes = _filter_pages_by_rig(pages, effective_filter)

    if effective_filter and not text:
        full_text = "\n".join(pages)
        return {
            "well_name": _WELL_NAME.search(full_text).group(1) if _WELL_NAME.search(full_text) else "",
            "rig": _extract_rig_from_text(full_text),
            "report_period_from": "",
            "report_period_to": "",
            "entries": [],
            "oamn_entries": [],
            "total_duration": 0.0,
            "skipped_rig_mismatch": True,
            "skip_reason": "rig_mismatch",
            "matched_pages": [],
        }

    if not text:
        return {
            "well_name": "",
            "rig": "",
            "report_period_from": "",
            "report_period_to": "",
            "entries": [],
            "oamn_entries": [],
            "total_duration": 0.0,
            "matched_pages": [],
        }

    start = _find_time_log_start(text, table_names)
    if start < 0:
        return {
            "well_name": _WELL_NAME.search(text).group(1) if _WELL_NAME.search(text) else "",
            "rig": _extract_rig_from_text(text),
            "report_period_from": "",
            "report_period_to": "",
            "entries": [],
            "oamn_entries": [],
            "total_duration": 0.0,
            "skip_reason": "no_matching_table",
            "matched_pages": [index + 1 for index in matched_page_indexes],
        }

    end = text.find("Phase Time and Cost Summary", start)
    if end < 0:
        total_match = re.search(r"^Total\b", text[start:], re.MULTILINE)
        end = start + total_match.start() if total_match else -1
    first_section = text[start:end].strip() if end > 0 else text[start : start + 2000].strip()

    if _is_job_time_log_section(first_section):
        data = extract_job_time_log(pdf_path, rig_filter=rig_filter, table_names=table_names)
        data["matched_pages"] = [index + 1 for index in matched_page_indexes]
        return data

    sections = _iter_classic_time_log_sections(text, table_names)
    entries: list[dict[str, Any]] = []
    oamn_entries: list[dict[str, Any]] = []
    wells: list[str] = []
    rigs: list[str] = []
    report_period_from = ""
    report_period_to = ""

    for section, section_start in sections:
        well, period_from, period_to, page_rig = _metadata_before(text, section_start)
        # Pages were already filtered by Rig; still skip odd mismatches if any.
        if effective_filter and page_rig and not rigs_match(page_rig, effective_filter):
            continue
        section_entries, section_oamn = _parse_time_log_section(section, table_names)
        entries.extend(section_entries)
        oamn_entries.extend(section_oamn)
        if well and well not in wells:
            wells.append(well)
        if page_rig and page_rig not in rigs:
            rigs.append(page_rig)
        if not report_period_from and period_from:
            report_period_from = period_from
            report_period_to = period_to

    well_name = ", ".join(wells)
    rig_name = ", ".join(rigs)
    if not well_name:
        well_match = _WELL_NAME.search(text)
        well_name = well_match.group(1) if well_match else ""
    if not rig_name:
        rig_name = _extract_rig_from_text(text)
    if not report_period_from:
        period_match = _REPORT_PERIOD.search(text)
        if period_match:
            report_period_from = period_match.group(1).strip()
            report_period_to = period_match.group(2).strip()

    skip_reason = ""
    if effective_filter and not entries and not oamn_entries:
        skip_reason = "empty_table" if sections else "no_matching_table"
    elif not entries and not oamn_entries:
        skip_reason = "empty_table" if sections else "no_matching_table"

    return {
        "well_name": well_name,
        "rig": rig_name,
        "report_period_from": report_period_from,
        "report_period_to": report_period_to,
        "entries": entries,
        "oamn_entries": oamn_entries,
        "total_duration": round(sum(entry["duration"] for entry in entries), 2),
        "skipped_rig_mismatch": skip_reason == "rig_mismatch",
        "skip_reason": skip_reason,
        "matched_pages": [index + 1 for index in matched_page_indexes],
    }



def _cell_text(value: object) -> str:
    return str(value).strip() if value else ""


def _find_table_columns(cells: list[str]) -> tuple[int | None, int | None]:
    from_col: int | None = None
    details_col: int | None = None
    for index, cell in enumerate(cells):
        label = cell.lower()
        if label == "from":
            from_col = index
        if "operation details" in label:
            details_col = index
    return from_col, details_col


def _extract_report_date_from_text(text: str) -> str:
    match = _REPORT_DATE.search(text)
    return match.group(1).strip() if match else ""


def _extract_operational_tables_from_page(page: Any) -> tuple[str, list[dict[str, str]]]:
    """Extract Operational Time Summary rows from one PDF page."""
    page_text = page.extract_text() or ""
    page_date = _extract_report_date_from_text(page_text)
    entries: list[dict[str, str]] = []

    for table in page.extract_tables() or []:
        if not table or len(table) < 3:
            continue

        header_row_index: int | None = None
        from_col: int | None = None
        details_col: int | None = None
        table_date = page_date

        for row_index, row in enumerate(table):
            cells = [_cell_text(cell) for cell in row]
            joined = " ".join(cells)

            if "DATE:" in joined.upper():
                match = _REPORT_DATE.search(joined)
                if match:
                    table_date = match.group(1).strip()

            found_from, found_details = _find_table_columns(cells)
            if found_from is not None and found_details is not None:
                header_row_index = row_index
                from_col = found_from
                details_col = found_details
                break

        if header_row_index is None or from_col is None or details_col is None:
            continue

        for row in table[header_row_index + 1 :]:
            cells = [_cell_text(cell) for cell in row]
            if not any(cells):
                continue

            time_from = cells[from_col] if from_col < len(cells) else ""
            operation_details = cells[details_col] if details_col < len(cells) else ""
            if not _FROM_TIME.match(time_from):
                break

            entries.append(
                {
                    "from": time_from,
                    "operation_details": operation_details,
                    "report_date": table_date,
                }
            )

    return page_date, entries


def extract_operational_time_summary(
    pdf_path: str | Path,
    *,
    rig_filter: str | None = None,
) -> dict[str, Any]:
    """Extract Operational Time Summary rows, optionally filtered by page Rig."""
    pdf_path = Path(pdf_path)
    entries: list[dict[str, str]] = []
    matched_rigs: list[str] = []
    matched_pages: list[int] = []
    report_date = ""

    pages = _iter_pdf_page_texts(pdf_path)
    all_pdf_rigs = _collect_rigs_from_pages(pages)

    with pdfplumber.open(pdf_path) as document:
        for page_index, page in enumerate(document.pages):
            page_text = pages[page_index] if page_index < len(pages) else (page.extract_text() or "")
            page_rig = _extract_rig_from_text(page_text)

            if rig_filter and (not page_rig or not rigs_match(page_rig, rig_filter)):
                continue

            page_date, page_entries = _extract_operational_tables_from_page(page)
            if not page_entries:
                continue

            rig_display = format_rig_for_display(page_rig)
            if rig_display != "unknown" and rig_display not in matched_rigs:
                matched_rigs.append(rig_display)
            matched_pages.append(page_index + 1)
            if not report_date and page_date:
                report_date = page_date
            entries.extend(page_entries)

    skip_reason = ""
    if not entries:
        if rig_filter:
            has_matching_rig_page = any(
                rigs_match(rig_filter, rig) for rig in all_pdf_rigs
            )
            skip_reason = "no_matching_table" if has_matching_rig_page else "rig_mismatch"
        else:
            skip_reason = "empty_table" if all_pdf_rigs else "no_matching_table"

    if matched_rigs:
        rig_display = ", ".join(matched_rigs)
    elif all_pdf_rigs:
        rig_display = ", ".join(all_pdf_rigs)
    else:
        rig_display = ""

    return {
        "source": "operational_time_summary",
        "date": report_date,
        "rig": rig_display,
        "entries": entries,
        "all_rigs": all_pdf_rigs,
        "skip_reason": skip_reason,
        "skipped_rig_mismatch": skip_reason == "rig_mismatch",
        "matched_pages": matched_pages,
    }


def _wants_operational_time_summary(table_names: list[str] | None) -> bool:
    if not table_names:
        return True
    normalized = {name.strip().casefold() for name in table_names if name and name.strip()}
    return _OPERATIONAL_TIME_SUMMARY.casefold() in normalized


def extract_soe_data(
    pdf_path: str | Path,
    *,
    rig_filter: str | None = None,
    table_names: list[str] | None = None,
) -> dict[str, Any]:
    """Extract SOE rows from either Time Log or Operational Time Summary PDFs.

    When ``rig_filter`` is provided, only Time Log sections whose ``Rig:``
    value matches that filter (same as the Excel Rig cell) are returned.
    """
    text = _extract_raw_text(Path(pdf_path))
    if _wants_operational_time_summary(table_names) and "Operational Time Summary" in text:
        return extract_operational_time_summary(pdf_path, rig_filter=rig_filter)

    data = extract_time_log(pdf_path, rig_filter=rig_filter, table_names=table_names)
    data["source"] = "time_log"
    return data


if __name__ == "__main__":
    import json
    import sys

    pdf = sys.argv[1] if len(sys.argv) > 1 else "SOE/template/1.pdf"
    print(json.dumps(extract_soe_data(pdf), indent=2, ensure_ascii=False))
