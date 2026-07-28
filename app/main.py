from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import APP_VERSION
from app.routers import generate, job_order, proforma, soe

APP_DIR = Path(__file__).resolve().parent

app = FastAPI(
    title="OCR Excel Generator",
    description="Upload PDFs and Excel templates to generate Proforma, SOE, and Job Order workbooks.",
    version=APP_VERSION,
)

app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")
templates = Jinja2Templates(directory=APP_DIR / "templates")
templates.env.globals["app_version"] = APP_VERSION

app.include_router(generate.router)
app.include_router(proforma.router)
app.include_router(soe.router)
app.include_router(job_order.router)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "index.html")


@app.get("/proforma", response_class=HTMLResponse)
async def proforma_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "proforma.html")


@app.get("/soe", response_class=HTMLResponse)
async def soe_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "soe.html")


@app.get("/job-order", response_class=HTMLResponse)
async def job_order_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "job_order.html")
