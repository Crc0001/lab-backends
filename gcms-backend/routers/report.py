import uuid
import os
from pathlib import Path
from fastapi import APIRouter, HTTPException, Query, Depends, UploadFile, File
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, desc, func
import aiofiles

from services.gcms_renderer import render_gcms_report, render_gcms_preview, render_gcms_pdf
from database import get_db
from models import GcmsReportDB, GcmsPayload
from audit_client import log_action

router = APIRouter(prefix="/api/v1/gcms", tags=["gcms"])

REPORTS_DIR = Path(os.environ.get("REPORTS_DIR", "static/reports"))
PERMANENT_REPORTS_DIR = Path("reports/gcms")
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "static/uploads"))
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
PERMANENT_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/tiff", "image/bmp"}


@router.post("/upload-media")
async def upload_gcms_media(file: UploadFile = File(...)):
    if file.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(status_code=400, detail=f"不支持的文件类型: {file.content_type}")
    ext = Path(file.filename).suffix or ".png"
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}{ext}"
    async with aiofiles.open(save_path, "wb") as out:
        await out.write(await file.read())
    return {"file_id": f"{file_id}{ext}", "url": f"/static/uploads/{file_id}{ext}"}


@router.post("/generate")
async def generate_gcms_report(payload: GcmsPayload, fmt: str = Query("docx"),
                                db: AsyncSession = Depends(get_db)):
    uid = uuid.uuid4().hex[:8]
    docx_path = pdf_path = docx_url = pdf_url = None

    if fmt in ("docx", "both"):
        filename = f"gcms_{uid}.docx"
        out_path = PERMANENT_REPORTS_DIR / filename
        render_gcms_report(str(out_path), payload.model_dump())
        docx_path = str(out_path)
        docx_url = f"/api/v1/gcms/download/{filename}"

    if fmt in ("pdf", "both"):
        filename_pdf = f"gcms_{uid}.pdf"
        out_path_pdf = PERMANENT_REPORTS_DIR / filename_pdf
        try:
            render_gcms_pdf(str(out_path_pdf), payload.model_dump())
            pdf_path = str(out_path_pdf)
            pdf_url = f"/api/v1/gcms/download/{filename_pdf}"
        except Exception as ex:
            raise HTTPException(status_code=500, detail=f"PDF生成失败: {ex}")

    record = GcmsReportDB(
        report_title=payload.report_title,
        doc_number=payload.doc_number,
        project_name=payload.project_name,
        user_id=payload.user_id,
        form_data=payload.model_dump(),
        docx_file_path=docx_path,
        pdf_file_path=pdf_path,
    )
    db.add(record)
    await db.commit()
    await db.refresh(record)

    await log_action(payload.user_id, f"生成气质报告({fmt}) - {payload.report_title}")

    return {
        "status": "success",
        "report_id": record.id,
        "download_url": docx_url or pdf_url,
        "docx_url": docx_url,
        "pdf_url": pdf_url,
    }


@router.get("/history")
async def get_gcms_history(user_id: str = Query("default_user"), page: int = 1,
                            page_size: int = 20, db: AsyncSession = Depends(get_db)):
    offset = (page - 1) * page_size
    total = await db.scalar(
        select(func.count()).select_from(GcmsReportDB).where(GcmsReportDB.user_id == user_id)
    )
    result = await db.execute(
        select(GcmsReportDB).where(GcmsReportDB.user_id == user_id)
        .order_by(desc(GcmsReportDB.created_at)).offset(offset).limit(page_size)
    )
    reports = result.scalars().all()
    return {
        "total": total, "page": page, "page_size": page_size,
        "reports": [
            {
                "id": r.id, "report_title": r.report_title, "doc_number": r.doc_number,
                "project_name": r.project_name, "created_at": r.created_at,
                "docx_url": f"/api/v1/gcms/download/{Path(r.docx_file_path).name}" if r.docx_file_path else None,
                "pdf_url": f"/api/v1/gcms/download/{Path(r.pdf_file_path).name}" if r.pdf_file_path else None,
            } for r in reports
        ]
    }


@router.get("/download/{filename}")
async def download_gcms_report(filename: str):
    if filename != Path(filename).name or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    path = PERMANENT_REPORTS_DIR / filename
    if not path.exists():
        path = REPORTS_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="报告文件不存在")
    if not path.resolve().is_relative_to(PERMANENT_REPORTS_DIR.resolve()) and \
       not path.resolve().is_relative_to(REPORTS_DIR.resolve()):
        raise HTTPException(status_code=400, detail="非法文件名")
    media_type = "application/pdf" if filename.endswith(".pdf") else \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path=str(path), media_type=media_type, filename=filename)


@router.post("/preview")
async def preview_gcms_report(payload: GcmsPayload):
    try:
        html_content = render_gcms_preview(payload.model_dump())
    except Exception as ex:
        raise HTTPException(status_code=500, detail=str(ex))
    return {"html": html_content}
