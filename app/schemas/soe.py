from pydantic import BaseModel, Field


class SoePdfSummary(BaseModel):
    filename: str
    source: str = ""
    well_name: str = ""
    rig: str = ""
    report_date: str = ""
    report_period_from: str = ""
    report_period_to: str = ""
    row_count: int
    skipped: bool = False
    skip_reason: str = ""


class SoeRow(BaseModel):
    date: str
    time: str
    event: str


class SoeResponse(BaseModel):
    pdf_summaries: list[SoePdfSummary]
    rows: list[SoeRow]
    row_count: int = Field(description="Total time-log rows appended")
    pdf_count: int = Field(description="Number of PDFs processed")
    download_url: str
    filename: str
