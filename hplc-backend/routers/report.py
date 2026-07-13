from fastapi import APIRouter, Depends, HTTPException
from pathlib import Path
from sqlalchemy.ext.asyncio import AsyncSession
from models import InspectionReport, ValidationReport, InspectionReportDB, ValidationReportDB
from services.renderer import (render_inspection_word, render_inspection_html,
                                render_validation_word, render_validation_html,
                                render_inspection_pdf, render_validation_pdf)
from database import get_db
from audit_client import log_action
import aiofiles
import uuid, os

router = APIRouter(prefix="/report")
REPORTS_DIR = Path("static/reports").resolve()

async def _save(content: bytes, ext: str) -> str:
    name = f"{uuid.uuid4().hex}.{ext}"
    path = REPORTS_DIR / name
    async with aiofiles.open(path, "wb") as f:
        await f.write(content)
    return f"/static/reports/{name}"

@router.post("/inspection/generate")
async def inspection_word(data: InspectionReport, db: AsyncSession = Depends(get_db)):
    buf = render_inspection_word(data)
    url = await _save(buf, "docx")
    db.add(InspectionReportDB(
        report_title=data.product_name or "检验记录",
        doc_number=data.doc_number,
        form_data=data.model_dump(),
        docx_file_path=url,
    ))
    await db.commit()
    await log_action("default_user", f"生成检验记录 DOCX：{data.product_name} {data.doc_number}")
    return {"download_url": url}

@router.post("/inspection/generate-pdf")
async def inspection_pdf(data: InspectionReport, db: AsyncSession = Depends(get_db)):
    buf = render_inspection_pdf(data)
    url = await _save(buf, "pdf")
    db.add(InspectionReportDB(
        report_title=data.product_name or "检验记录",
        doc_number=data.doc_number,
        form_data=data.model_dump(),
        pdf_file_path=url,
    ))
    await db.commit()
    await log_action("default_user", f"生成检验记录 PDF：{data.product_name} {data.doc_number}")
    return {"download_url": url}

@router.post("/inspection/preview")
async def inspection_preview(data: InspectionReport):
    return {"html": render_inspection_html(data)}

@router.post("/validation/generate")
async def validation_word(data: ValidationReport, db: AsyncSession = Depends(get_db)):
    buf = render_validation_word(data)
    url = await _save(buf, "docx")
    db.add(ValidationReportDB(
        report_title=data.project_name or "方法学验证报告",
        doc_number=data.doc_number,
        form_data=data.model_dump(),
        docx_file_path=url,
    ))
    await db.commit()
    await log_action("default_user", f"生成验证报告 DOCX：{data.project_name} {data.doc_number}")
    return {"download_url": url}

@router.post("/validation/generate-pdf")
async def validation_pdf(data: ValidationReport, db: AsyncSession = Depends(get_db)):
    buf = render_validation_pdf(data)
    url = await _save(buf, "pdf")
    db.add(ValidationReportDB(
        report_title=data.project_name or "方法学验证报告",
        doc_number=data.doc_number,
        form_data=data.model_dump(),
        pdf_file_path=url,
    ))
    await db.commit()
    await log_action("default_user", f"生成验证报告 PDF：{data.project_name} {data.doc_number}")
    return {"download_url": url}

@router.post("/validation/preview")
async def validation_preview(data: ValidationReport):
    return {"html": render_validation_html(data)}
