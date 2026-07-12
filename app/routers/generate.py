from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR, UPLOAD_DIR
from app.schemas.combined import (
    CombinedJobOrderResult,
    CombinedProformaResult,
    CombinedResponse,
    CombinedSoeResult,
)
from app.schemas.job_order import JobOrderLine
from app.schemas.proforma import ProformaItem
from app.schemas.soe import SoePdfSummary, SoeRow
from app.services.combined_service import process_combined

router = APIRouter(prefix="/api/generate", tags=["generate"])

_ALLOWED_JOB_ORDER_SOURCES = {"auto", "1c", "running"}


@router.post("", response_model=CombinedResponse)
async def generate_workbook(
    excel: UploadFile = File(..., description="Excel template (.xlsm)"),
    proforma_pdf: UploadFile | None = File(None, description="Purchase order PDF"),
    job_order_pdf: UploadFile | None = File(None, description="Completion program PDF"),
    job_order_source: str = Form(default="auto", description="Job Order template: auto, 1c, or running"),
    soe_pdfs: list[UploadFile] = File(default=[], description="SOE report PDFs"),
    soe_pdf_names: list[str] = Form(default=[], description="Display names for SOE PDFs"),
) -> CombinedResponse:
    if not excel.filename or not excel.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Excel template (.xlsm or .xlsx) is required.")
    if job_order_source not in _ALLOWED_JOB_ORDER_SOURCES:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid job_order_source {job_order_source!r}. Use auto, 1c, or running.",
        )

    has_proforma = proforma_pdf is not None and proforma_pdf.filename
    has_job_order = job_order_pdf is not None and job_order_pdf.filename
    has_soe = bool(soe_pdfs)

    if not has_proforma and not has_job_order and not has_soe:
        raise HTTPException(
            status_code=400,
            detail="Provide at least one PDF: Proforma, SOE, or Job Order.",
        )

    if has_proforma and not proforma_pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Proforma file must be a PDF.")
    if has_job_order and not job_order_pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Job Order file must be a PDF.")

    job_id = uuid.uuid4().hex[:12]
    excel_suffix = Path(excel.filename).suffix.lower()
    excel_path = UPLOAD_DIR / f"{job_id}{excel_suffix}"
    proforma_path: Path | None = None
    job_order_path: Path | None = None
    soe_entries: list[tuple[Path, str]] = []

    try:
        excel_path.write_bytes(await excel.read())

        if has_proforma:
            proforma_path = UPLOAD_DIR / f"{job_id}_proforma.pdf"
            proforma_path.write_bytes(await proforma_pdf.read())

        if has_job_order:
            job_order_path = UPLOAD_DIR / f"{job_id}_job_order.pdf"
            job_order_path.write_bytes(await job_order_pdf.read())

        for index, pdf in enumerate(soe_pdfs):
            if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid SOE PDF: {pdf.filename or 'unnamed'}",
                )
            display_name = (
                soe_pdf_names[index]
                if index < len(soe_pdf_names) and soe_pdf_names[index].strip()
                else (pdf.filename or f"soe_{index + 1}.pdf")
            )
            pdf_path = UPLOAD_DIR / f"{job_id}_soe_{index}.pdf"
            pdf_path.write_bytes(await pdf.read())
            soe_entries.append((pdf_path, display_name))

        data, output_path = process_combined(
            excel_path,
            proforma_pdf=proforma_path,
            soe_pdfs=soe_entries or None,
            job_order_pdf=job_order_path,
            job_order_source=job_order_source,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {exc}") from exc
    finally:
        if proforma_path:
            proforma_path.unlink(missing_ok=True)
        if job_order_path:
            job_order_path.unlink(missing_ok=True)
        for pdf_path, _ in soe_entries:
            pdf_path.unlink(missing_ok=True)
        excel_path.unlink(missing_ok=True)

    filename = output_path.name
    proforma_result = None
    if data.get("proforma"):
        proforma_result = CombinedProformaResult(
            items=[ProformaItem(**item) for item in data["proforma"]["items"]],
            item_count=data["proforma"]["item_count"],
            gross_total=data["proforma"]["gross_total"],
        )

    soe_result = None
    if data.get("soe"):
        soe_result = CombinedSoeResult(
            pdf_summaries=[SoePdfSummary(**summary) for summary in data["soe"]["pdf_summaries"]],
            rows=[SoeRow(**row) for row in data["soe"]["rows"]],
            row_count=data["soe"]["row_count"],
            pdf_count=data["soe"]["pdf_count"],
            rig_filter=data["soe"].get("rig_filter", ""),
        )

    job_order_result = None
    if data.get("job_order"):
        job_order_result = CombinedJobOrderResult(
            section_title=data["job_order"]["section_title"],
            source=data["job_order"]["source"],
            lines=[JobOrderLine(**line) for line in data["job_order"]["lines"]],
            line_count=data["job_order"]["line_count"],
        )

    return CombinedResponse(
        processed_sections=data["processed_sections"],
        proforma=proforma_result,
        soe=soe_result,
        job_order=job_order_result,
        download_url=f"/api/generate/download/{filename}",
        filename=filename,
    )


@router.get("/download/{filename}")
async def download_workbook(filename: str) -> FileResponse:
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
