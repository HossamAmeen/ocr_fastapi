from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR, UPLOAD_DIR
from app.schemas.soe import SoePdfSummary, SoeResponse, SoeRow
from app.services.soe_service import process_soe

router = APIRouter(prefix="/api/soe", tags=["soe"])


@router.post("/generate", response_model=SoeResponse)
async def generate_soe(
    pdfs: list[UploadFile] = File(..., description="SOE daily operations report PDFs"),
    excel: UploadFile = File(..., description="Excel template (.xlsm)"),
) -> SoeResponse:
    if not pdfs:
        raise HTTPException(status_code=400, detail="At least one PDF file is required.")
    if not excel.filename or not excel.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Excel template (.xlsm) is required.")

    job_id = uuid.uuid4().hex[:12]
    pdf_paths: list[Path] = []
    excel_path = UPLOAD_DIR / f"{job_id}.xlsm"

    try:
        for index, pdf in enumerate(pdfs):
            if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid PDF file: {pdf.filename or 'unnamed'}",
                )
            pdf_path = UPLOAD_DIR / f"{job_id}_{index}.pdf"
            pdf_path.write_bytes(await pdf.read())
            pdf_paths.append(pdf_path)

        excel_path.write_bytes(await excel.read())
        pdf_summaries, rows, output_path, row_count = process_soe(pdf_paths, excel_path)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {exc}") from exc
    finally:
        for pdf_path in pdf_paths:
            pdf_path.unlink(missing_ok=True)
        excel_path.unlink(missing_ok=True)

    filename = output_path.name
    return SoeResponse(
        pdf_summaries=[SoePdfSummary(**summary) for summary in pdf_summaries],
        rows=[SoeRow(**row) for row in rows],
        row_count=row_count,
        pdf_count=len(pdfs),
        download_url=f"/api/soe/download/{filename}",
        filename=filename,
    )


@router.get("/download/{filename}")
async def download_soe(filename: str) -> FileResponse:
    safe_name = Path(filename).name
    file_path = OUTPUT_DIR / safe_name

    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    return FileResponse(
        file_path,
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        filename=safe_name,
    )
