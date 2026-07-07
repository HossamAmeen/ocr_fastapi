from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR, UPLOAD_DIR
from app.schemas.job_order import JobOrderLine, JobOrderResponse
from app.services.job_order_service import process_job_order

router = APIRouter(prefix="/api/job-order", tags=["job-order"])

_ALLOWED_SOURCES = {"auto", "1c", "running"}


@router.post("/generate", response_model=JobOrderResponse)
async def generate_job_order(
    pdf: UploadFile = File(..., description="Completion program PDF"),
    excel: UploadFile = File(..., description="Excel template (.xlsm)"),
    source: str = Form(default="auto", description="Template: auto, 1c, or running"),
) -> JobOrderResponse:
    if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="PDF file is required.")
    if not excel.filename or not excel.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Excel template (.xlsm or .xlsx) is required.")
    if source not in _ALLOWED_SOURCES:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid source {source!r}. Use auto, 1c, or running.",
        )

    job_id = uuid.uuid4().hex[:12]
    pdf_path = UPLOAD_DIR / f"{job_id}.pdf"
    excel_suffix = Path(excel.filename).suffix.lower()
    excel_path = UPLOAD_DIR / f"{job_id}{excel_suffix}"

    pdf_path.write_bytes(await pdf.read())
    excel_path.write_bytes(await excel.read())

    try:
        data, lines, output_path, appended_rows = process_job_order(
            pdf_path,
            excel_path,
            source=source,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {exc}") from exc
    finally:
        pdf_path.unlink(missing_ok=True)
        excel_path.unlink(missing_ok=True)

    filename = output_path.name
    media_type = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if filename.lower().endswith(".xlsx")
        else "application/vnd.ms-excel.sheet.macroEnabled.12"
    )

    return JobOrderResponse(
        section_title=str(data.get("section_title") or ""),
        source=str(data.get("source") or source),
        lines=[JobOrderLine(**line) for line in lines],
        line_count=appended_rows,
        download_url=f"/api/job-order/download/{filename}",
        filename=filename,
    )


@router.get("/download/{filename}")
async def download_job_order(filename: str) -> FileResponse:
    safe_name = Path(filename).name
    file_path = OUTPUT_DIR / safe_name

    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    if safe_name.lower().endswith(".xlsx"):
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        media_type = "application/vnd.ms-excel.sheet.macroEnabled.12"

    return FileResponse(
        file_path,
        media_type=media_type,
        filename=safe_name,
    )
