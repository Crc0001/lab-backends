import uuid
import os
from pathlib import Path
import numpy as np
from fastapi import APIRouter, HTTPException, Query, Depends
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession

from database import get_db
from models import ReportPayload, ReportResponse, LinearityInput, LinearityResult, ValidationReportDB
from services.document_renderer import render_report, render_preview_html, render_pdf

router = APIRouter(prefix="/api/v1/report", tags=["report"])

REPORTS_DIR = Path(os.environ.get("REPORTS_DIR", "static/reports"))
REPORTS_DIR.mkdir(parents=True, exist_ok=True)


@router.post("/generate", response_model=ReportResponse)
async def generate_report(payload: ReportPayload, fmt: str = Query("docx", description="docx 或 pdf"),
                           db: AsyncSession = Depends(get_db)):
    uid = uuid.uuid4().hex[:8]
    if fmt == "pdf":
        filename = f"rep_{uid}.pdf"
        report_path = REPORTS_DIR / filename
        try:
            render_pdf(str(report_path), payload)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"PDF生成失败: {e}")
    else:
        filename = f"rep_{uid}.docx"
        report_path = REPORTS_DIR / filename
        try:
            render_report(str(report_path), payload)
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    record = ValidationReportDB(
        report_title=getattr(payload, "report_title", ""),
        doc_number=getattr(payload, "doc_number", ""),
        project_name=getattr(payload, "project_name", ""),
        user_id=getattr(payload, "user_id", "default_user"),
        form_data=payload.model_dump(),
        docx_file_path=str(report_path) if fmt != "pdf" else None,
        pdf_file_path=str(report_path) if fmt == "pdf" else None,
    )
    db.add(record)
    try:
        await db.commit()
    except Exception:
        report_path.unlink(missing_ok=True)
        raise
    return ReportResponse(status="success", download_url=f"/api/v1/report/download/{filename}")


@router.post("/preview")
async def preview_report(payload: ReportPayload):
    try:
        html_content = render_preview_html(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return {"html": html_content}


@router.get("/download/{filename}")
async def download_report(filename: str):
    if filename != Path(filename).name or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    path = REPORTS_DIR / filename
    if not path.resolve().is_relative_to(REPORTS_DIR.resolve()):
        raise HTTPException(status_code=400, detail="非法文件名")
    if not path.exists():
        raise HTTPException(status_code=404, detail="报告文件不存在")
    media_type = "application/pdf" if filename.endswith(".pdf") else \
        "text/html; charset=utf-8" if filename.endswith(".html") else \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path=str(path), media_type=media_type, filename=filename)


@router.post("/calc/linearity", response_model=LinearityResult)
async def calc_linearity(data: LinearityInput):
    x = np.array(data.concentrations)
    y = np.array(data.peak_areas)
    coeffs = np.polyfit(x, y, 1)
    slope, intercept = float(coeffs[0]), float(coeffs[1])
    r = float(np.corrcoef(x, y)[0, 1])
    r2 = r ** 2
    sign = "+" if intercept >= 0 else "-"
    equation = f"y = {slope:.2f}x {sign} {abs(intercept):.2f}"
    return LinearityResult(equation=equation, r=f"{r:.4f}", r2=f"{r2:.4f}")
