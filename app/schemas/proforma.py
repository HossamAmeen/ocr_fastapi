from pydantic import BaseModel, Field


class ProformaItem(BaseModel):
    sno: int
    description: str
    per_day_rate: float
    days: float
    total: float = Field(description="per_day_rate * days")


class ProformaResponse(BaseModel):
    items: list[ProformaItem]
    item_count: int
    gross_total: float
    download_url: str
    filename: str
