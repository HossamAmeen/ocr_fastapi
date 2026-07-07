from pydantic import BaseModel, Field


class JobOverLine(BaseModel):
    line_no: int
    text: str


class JobOverResponse(BaseModel):
    section_title: str = ""
    source: str = Field(description="Detected or selected extraction template")
    lines: list[JobOverLine]
    line_count: int = Field(description="Procedure lines appended to JOB ORDER sheet")
    download_url: str
    filename: str
