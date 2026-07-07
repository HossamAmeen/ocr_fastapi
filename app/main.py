from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.routers import job_over, proforma, soe

APP_DIR = Path(__file__).resolve().parent

app = FastAPI(
    title="OCR Excel Generator",
    description="Upload PDFs and Excel templates to generate Proforma, SOE, and Job Over workbooks.",
    version="1.1.0",
)

app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")
templates = Jinja2Templates(directory=APP_DIR / "templates")

app.include_router(proforma.router)
app.include_router(soe.router)
app.include_router(job_over.router)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "index.html")


@app.get("/soe", response_class=HTMLResponse)
async def soe_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "soe.html")


@app.get("/job-over", response_class=HTMLResponse)
async def job_over_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "job_over.html")
