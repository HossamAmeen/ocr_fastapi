"""Extract job order procedure text from completion program PDFs."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pdf_extractor.extractor import _extract_raw_text

# 1-C detailed procedure template
START_MARKERS_1C = (
    "1-C Upper Completion - Running Procedure",
    "1-C procedure",
    "1-C Upper Completion",
)
END_MARKERS_1C = (
    "The final report requested to be shared from Completion Engineer:",
    "1-D Additional Information",
    "1-D Additional Information or Contigency Plan",
)

# Running completion summary / checklist template
RUNNING_SUMMARY_START = (
    "1 Running Upper Completion",
    "Running Upper Completion",
)
RUNNING_SUMMARY_END = "1.15 Nipple Up X-Mast Tree"

# Running completion detailed steps template
RUNNING_PROCEDURE_START = re.compile(
    r"(?m)^Running Completion\s*(?:\r?\n|$)"
)
RUNNING_PROCEDURE_END = re.compile(
    r"(?m)^32 Perform final tests of TR-SCSSSV|^4 SECURE WELL AND RELEASE RIG",
    re.IGNORECASE,
)

_HEADER_LINE = re.compile(
    r"^(?:"
    r"ADNOC Classification:.*"
    r"|#"
    r"|UPPER COMPLETION PROGRAM"
    r"|COMPLETION PROGRAM"
    r"|COMP_[A-Z0-9_]+\(program\).*"
    r"|Page \d+ of \d+"
    r"|AD-\d+\s+BB-\d+\s+OP\d+/\w+"
    r"|Packer Tests, X-mass Tree installation"
    r"|Gauge Cutter Run"
    r")$",
    re.IGNORECASE,
)
_STEP_LINE = re.compile(r"^\d+(?:\.\d+)+\s+")
_NUMBERED_STEP = re.compile(r"^\d+\s+")


def _find_start_1c(text: str) -> int:
    for marker in START_MARKERS_1C:
        index = text.find(marker)
        if index >= 0:
            return index
    return -1


def _find_end_1c(text: str, start: int) -> int:
    search_from = start + 1
    for marker in END_MARKERS_1C:
        index = text.find(marker, search_from)
        if index >= 0:
            return index

    xmas_index = text.find("1.15 Nipple Up X-Mast Tree", search_from)
    if xmas_index < 0:
        return len(text)

    tail = text[xmas_index:]
    match = re.search(
        r"(?ms)^1\.15 Nipple Up X-Mast Tree\s*(.*?)(?=The final report requested|\Z)",
        tail,
    )
    if not match:
        return xmas_index + len("1.15 Nipple Up X-Mast Tree")
    return xmas_index + match.end()


def _find_running_summary_section(text: str) -> str:
    start = -1
    for marker in RUNNING_SUMMARY_START:
        index = text.find(marker)
        if index >= 0:
            start = index
            break
    if start < 0:
        return ""

    end = text.find(RUNNING_SUMMARY_END, start)
    if end < 0:
        return ""
    end += len(RUNNING_SUMMARY_END)
    return text[start:end].strip()


def _find_running_procedure_section(text: str) -> str:
    match = RUNNING_PROCEDURE_START.search(text)
    if not match:
        return ""

    start = match.start()
    end_match = RUNNING_PROCEDURE_END.search(text, match.end())
    if not end_match:
        xmas_match = re.search(
            r"(?is)31 Nipple up Xmas tree.*?(?=^32 Perform final tests|^4 SECURE WELL|\Z)",
            text[start:],
        )
        if not xmas_match:
            return ""
        return text[start : start + xmas_match.end()].strip()

    return text[start : end_match.start()].strip()


def _is_header_line(line: str) -> bool:
    return bool(_HEADER_LINE.match(line.strip()))


def _is_new_item(line: str, numbered_steps: bool = False) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if _STEP_LINE.match(stripped):
        return True
    if numbered_steps and _NUMBERED_STEP.match(stripped):
        return True
    return stripped.startswith(("*", "!", "~", "-", "•", "o "))


def _merge_wrapped_lines(lines: list[str], numbered_steps: bool = False) -> list[str]:
    merged: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if not merged or _is_new_item(stripped, numbered_steps=numbered_steps):
            merged.append(stripped)
            continue
        merged[-1] = f"{merged[-1]} {stripped}".strip()
    return merged


def _build_result(pdf_path: Path, source: str, merged_lines: list[str]) -> dict[str, Any]:
    return {
        "pdf": str(pdf_path),
        "source": source,
        "section_title": merged_lines[0] if merged_lines else "",
        "lines": [
            {"line_no": index, "text": line}
            for index, line in enumerate(merged_lines, start=1)
        ],
    }


def _extract_section_lines(
    section: str,
    *,
    numbered_steps: bool = False,
    merge_wrapped: bool = True,
) -> list[str]:
    raw_lines = [line.strip() for line in section.splitlines() if line.strip()]
    content_lines = [line for line in raw_lines if not _is_header_line(line)]
    if not merge_wrapped:
        return content_lines
    return _merge_wrapped_lines(content_lines, numbered_steps=numbered_steps)


def extract_job_order_procedure(pdf_path: str | Path) -> dict[str, Any]:
    """Extract the 1-C running procedure through the X-Mast Tree step, line by line."""
    pdf_path = Path(pdf_path)
    text = _extract_raw_text(pdf_path)

    start = _find_start_1c(text)
    if start < 0:
        return _build_result(pdf_path, "job_order_1c", [])

    end = _find_end_1c(text, start)
    section = text[start:end].strip()
    merged_lines = _extract_section_lines(section, merge_wrapped=False)
    return _build_result(pdf_path, "job_order_1c", merged_lines)


def _split_running_completion_title(lines: list[str]) -> list[str]:
    if not lines or not lines[0].startswith("Running Completion"):
        return lines

    first = lines[0]
    if "DS," not in first:
        return lines

    title, intro = first.split("DS,", 1)
    return [title.strip(), f"DS,{intro.strip()}", *lines[1:]]


def extract_running_completion(pdf_path: str | Path) -> dict[str, Any]:
    """Extract running completion steps from the summary or detailed checklist template."""
    pdf_path = Path(pdf_path)
    text = _extract_raw_text(pdf_path)

    procedure_section = _find_running_procedure_section(text)
    if procedure_section:
        merged_lines = _extract_section_lines(procedure_section, numbered_steps=True)
        merged_lines = _split_running_completion_title(merged_lines)
        if merged_lines:
            return _build_result(pdf_path, "running_completion", merged_lines)

    summary_section = _find_running_summary_section(text)
    if summary_section:
        merged_lines = _extract_section_lines(summary_section)
        if merged_lines:
            return _build_result(pdf_path, "running_completion_summary", merged_lines)

    return _build_result(pdf_path, "running_completion", [])


def detect_job_order_source(text: str) -> str:
    """Detect which job order template is present in the PDF text."""
    if _find_start_1c(text) >= 0:
        return "1c"
    if RUNNING_PROCEDURE_START.search(text):
        return "running"
    summary = _find_running_summary_section(text)
    if summary and "1.1 Conduct" in summary:
        return "running"
    if summary:
        return "running"
    return "unknown"


def extract_job_order_data(
    pdf_path: str | Path,
    source: str = "auto",
) -> dict[str, Any]:
    """Extract job order data using the requested or detected template."""
    pdf_path = Path(pdf_path)
    text = _extract_raw_text(pdf_path)

    selected = source if source != "auto" else detect_job_order_source(text)
    if selected == "1c":
        return extract_job_order_procedure(pdf_path)
    if selected == "running":
        return extract_running_completion(pdf_path)

    data = extract_running_completion(pdf_path)
    if data["lines"]:
        return data
    return extract_job_order_procedure(pdf_path)
