import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    # Packaged PyInstaller mode: paths relative to executable location
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # Standard Python execution: paths relative to source file
    BASE_DIR = Path(__file__).resolve().parent.parent

APP_VERSION = "4"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

REQUIRED_SHEETS = [
    "JOB ORDER",
    "SOE",
    "Proforma",
]


def validate_template(template_path: str | Path) -> None:
    """Validate that an Excel template workbook contains all required sheets."""
    from excel_utils.workbook_utils import validate_workbook_sheets

    validate_workbook_sheets(template_path, REQUIRED_SHEETS)


