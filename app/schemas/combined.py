from pydantic import BaseModel, Field

from app.schemas.job_order import JobOrderLine
from app.schemas.proforma import ProformaItem
from app.schemas.soe import SoePdfSummary, SoeRow


class CombinedProformaResult(BaseModel):
    items: list[ProformaItem]
    item_count: int
    gross_total: float


class CombinedJobOrderResult(BaseModel):
    section_title: str = ""
    source: str = ""
    lines: list[JobOrderLine]
    line_count: int


class CombinedSoeResult(BaseModel):
    pdf_summaries: list[SoePdfSummary]
    rows: list[SoeRow]
    row_count: int
    pdf_count: int
    rig_filter: str = ""


class CombinedResponse(BaseModel):
    processed_sections: list[str] = Field(description="Sections written to the workbook")
    proforma: CombinedProformaResult | None = None
    soe: CombinedSoeResult | None = None
    job_order: CombinedJobOrderResult | None = None
    download_url: str
    filename: str
