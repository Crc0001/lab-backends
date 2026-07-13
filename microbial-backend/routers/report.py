import uuid
import os
from datetime import datetime
from pathlib import Path
from fastapi import APIRouter, HTTPException, UploadFile, File, Query, Depends
from fastapi.responses import FileResponse, HTMLResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, desc, func
import aiofiles

from database import get_db
from models import ReportPayload, UploadMediaResponse, ReportGenerationResponse, MicrobialReportDB
from services.document_renderer import render_report, render_preview_html, render_pdf
from audit_client import log_action

router = APIRouter(prefix="/api/v1/report", tags=["report"])

UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "static/uploads"))
REPORTS_DIR = Path(os.environ.get("REPORTS_DIR", "static/reports"))
PERMANENT_DIR = Path("reports/microbial")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
PERMANENT_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/tiff", "image/bmp"}


@router.post("/upload-media", response_model=UploadMediaResponse)
async def upload_media(file: UploadFile = File(...)):
    if file.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(status_code=400, detail=f"不支持的文件类型: {file.content_type}")
    ext = Path(file.filename).suffix or ".png"
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}{ext}"
    async with aiofiles.open(save_path, "wb") as out:
        await out.write(await file.read())
    return UploadMediaResponse(file_id=f"{file_id}{ext}", url=f"/static/uploads/{file_id}{ext}")


@router.post("/upload-txt", response_model=UploadMediaResponse)
async def upload_txt(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".txt"):
        raise HTTPException(status_code=400, detail="仅支持 .txt 文件")
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}.txt"
    async with aiofiles.open(save_path, "wb") as out:
        await out.write(await file.read())
    return UploadMediaResponse(file_id=f"{file_id}.txt", url=str(save_path))


@router.post("/generate", response_model=ReportGenerationResponse)
async def generate_report(payload: ReportPayload, fmt: str = Query("docx"), db: AsyncSession = Depends(get_db)):
    uid = uuid.uuid4().hex[:8]
    docx_path = pdf_path = docx_url = pdf_url = None

    if fmt in ("docx", "both"):
        filename = f"microbial_{uid}.docx"
        out_path = PERMANENT_DIR / filename
        try:
            render_report(str(out_path), payload)
            docx_path = str(out_path)
            docx_url = f"/api/v1/report/download/{filename}"
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    if fmt in ("pdf", "both"):
        filename_pdf = f"microbial_{uid}.pdf"
        out_path_pdf = PERMANENT_DIR / filename_pdf
        try:
            render_pdf(str(out_path_pdf), payload)
            pdf_path = str(out_path_pdf)
            pdf_url = f"/api/v1/report/download/{filename_pdf}"
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"PDF生成失败: {e}")

    record = MicrobialReportDB(
        report_title=payload.report_title,
        doc_number=payload.doc_number,
        project_name=payload.project_name,
        user_id=payload.user_id,
        form_data=payload.model_dump(),
        docx_file_path=docx_path,
        pdf_file_path=pdf_path,
        docx_download_url=docx_url,
        pdf_download_url=pdf_url,
    )
    db.add(record)
    await db.commit()
    await db.refresh(record)

    await log_action(payload.user_id, f"生成微生物报告({fmt}) - {payload.report_title}")

    download_url = docx_url or pdf_url
    return ReportGenerationResponse(status="success", download_url=download_url, report_id=record.id,
                                    docx_url=docx_url, pdf_url=pdf_url)


@router.get("/history")
async def get_history(user_id: str = Query("default_user"), page: int = 1, page_size: int = 20,
                      db: AsyncSession = Depends(get_db)):
    offset = (page - 1) * page_size
    total = await db.scalar(select(func.count()).select_from(MicrobialReportDB).where(MicrobialReportDB.user_id == user_id))
    result = await db.execute(
        select(MicrobialReportDB).where(MicrobialReportDB.user_id == user_id)
        .order_by(desc(MicrobialReportDB.created_at)).offset(offset).limit(page_size)
    )
    reports = result.scalars().all()
    return {
        "total": total, "page": page, "page_size": page_size,
        "reports": [
            {
                "id": r.id, "report_title": r.report_title, "doc_number": r.doc_number,
                "project_name": r.project_name, "created_at": r.created_at,
                "docx_url": r.docx_download_url, "pdf_url": r.pdf_download_url,
            } for r in reports
        ]
    }


@router.post("/preview")
async def preview_report(payload: ReportPayload):
    try:
        html_content = render_preview_html(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return HTMLResponse(content=html_content)


@router.get("/download/{filename}")
async def download_report(filename: str):
    if filename != Path(filename).name or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    report_path = PERMANENT_DIR / filename
    if not report_path.exists():
        report_path = REPORTS_DIR / filename
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="报告文件不存在")
    if not report_path.resolve().is_relative_to(PERMANENT_DIR.resolve()) and \
       not report_path.resolve().is_relative_to(REPORTS_DIR.resolve()):
        raise HTTPException(status_code=400, detail="非法文件名")
    media_type = "application/pdf" if filename.endswith(".pdf") else \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path=str(report_path), media_type=media_type, filename=filename)
