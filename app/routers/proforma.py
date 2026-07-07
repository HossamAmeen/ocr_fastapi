from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR, UPLOAD_DIR
from app.schemas.proforma import ProformaItem, ProformaResponse
from app.services.proforma_service import process_proforma

router = APIRouter(prefix="/api/proforma", tags=["proforma"])


@router.post("/generate", response_model=ProformaResponse)
async def generate_proforma(
    pdf: UploadFile = File(..., description="Purchase order PDF"),
    excel: UploadFile = File(..., description="Excel template (.xlsm)"),
) -> ProformaResponse:
    if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="PDF file is required.")
    if not excel.filename or not excel.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Excel template (.xlsm) is required.")

    job_id = uuid.uuid4().hex[:12]
    pdf_path = UPLOAD_DIR / f"{job_id}.pdf"
    excel_path = UPLOAD_DIR / f"{job_id}.xlsm"

    pdf_bytes = await pdf.read()
    excel_bytes = await excel.read()
    pdf_path.write_bytes(pdf_bytes)
    excel_path.write_bytes(excel_bytes)

    try:
        items, output_path = process_proforma(pdf_path, excel_path)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {exc}") from exc
    finally:
        pdf_path.unlink(missing_ok=True)
        excel_path.unlink(missing_ok=True)

    gross_total = sum(item["total"] for item in items)
    filename = output_path.name

    return ProformaResponse(
        items=[ProformaItem(**item) for item in items],
        item_count=len(items),
        gross_total=gross_total,
        download_url=f"/api/proforma/download/{filename}",
        filename=filename,
    )


@router.get("/download/{filename}")
async def download_proforma(filename: str) -> FileResponse:
    safe_name = Path(filename).name
    file_path = OUTPUT_DIR / safe_name

    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    return FileResponse(
        file_path,
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        filename=safe_name,
    )
