from pydantic import BaseModel, Field


class JobOrderLine(BaseModel):
    line_no: int
    text: str


class JobOrderResponse(BaseModel):
    section_title: str = ""
    source: str = Field(description="Detected or selected extraction template")
    lines: list[JobOrderLine]
    line_count: int = Field(description="Procedure lines appended to JOB ORDER sheet")
    download_url: str
    filename: str
