from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR, UPLOAD_DIR
from app.schemas.soe import SoePdfSummary, SoeResponse, SoeRow
from app.services.soe_service import parse_table_names, process_soe

router = APIRouter(prefix="/api/soe", tags=["soe"])


@router.post("/generate", response_model=SoeResponse)
async def generate_soe(
    pdfs: list[UploadFile] = File(..., description="SOE daily operations report PDFs"),
    pdf_names: list[str] = Form(default=[], description="Display names for uploaded PDFs"),
    table_names: list[str] = Form(
        default=[],
        description="PDF table titles to extract (one per form field, or comma-separated)",
    ),
    excel: UploadFile = File(..., description="Excel template (.xlsm)"),
) -> SoeResponse:
    if not pdfs:
        raise HTTPException(status_code=400, detail="At least one PDF file is required.")
    if not excel.filename or not excel.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Excel template (.xlsm) is required.")

    job_id = uuid.uuid4().hex[:12]
    pdf_entries: list[tuple[Path, str]] = []
    excel_path = UPLOAD_DIR / f"{job_id}.xlsm"

    try:
        for index, pdf in enumerate(pdfs):
            if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid PDF file: {pdf.filename or 'unnamed'}",
                )
            display_name = (
                pdf_names[index]
                if index < len(pdf_names) and pdf_names[index].strip()
                else (pdf.filename or f"upload_{index + 1}.pdf")
            )
            pdf_path = UPLOAD_DIR / f"{job_id}_{index}.pdf"
            pdf_path.write_bytes(await pdf.read())
            pdf_entries.append((pdf_path, display_name))

        excel_path.write_bytes(await excel.read())
        parsed_table_names = parse_table_names(table_names)
        pdf_summaries, rows, output_path, row_count = process_soe(
            pdf_entries,
            excel_path,
            table_names=parsed_table_names,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {exc}") from exc
    finally:
        for pdf_path, _ in pdf_entries:
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
