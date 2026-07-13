import uuid
import os
from pathlib import Path
from fastapi import APIRouter, HTTPException, Query, Depends, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, desc, func
import aiofiles

from services.detection_renderer import render_detection_report, render_detection_preview, render_detection_pdf
from database import get_db
from models import DetectionReportDB

router = APIRouter(prefix="/api/v1/detection", tags=["detection"])

REPORTS_DIR = Path(os.environ.get("REPORTS_DIR", "static/reports"))
PERMANENT_REPORTS_DIR = Path("reports/detection")
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "static/uploads"))
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
PERMANENT_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/tiff", "image/bmp"}


class SignatureAuditEntry(BaseModel):
    previous_file_id: Optional[str] = None
    new_file_id: Optional[str] = None
    reason: str = ""
    changed_at: str = ""


class DetectionPayload(BaseModel):
    report_title: str = ""
    doc_number: str = ""
    security_level: str = ""
    project_name: str = ""
    user_id: str = "default_user"
    personnel: List[dict] = []
    approver_signatures: Optional[List[Optional[str]]] = None
    approver_signature_audits: List[SignatureAuditEntry] = []
    instruments: List[List[str]] = []
    instrument_headers: List[str] = []
    reagents: List[List[str]] = []
    references: List[List[str]] = []
    s21_rows: List[List[str]] = []
    s22_rows: List[List[str]] = []
    s23_rows: List[List[str]] = []
    s24_rows: List[List[str]] = []
    s25_rows: List[List[str]] = []
    s26_rows: List[List[str]] = []
    s27_rows: List[List[str]] = []
    ms_values: List[str] = []
    ms_methods: List[List[str]] = []
    hplc_methods: List[dict] = []
    s41_rows: List[List[str]] = []
    s42_rows: List[List[str]] = []
    s43_headers: List[str] = []
    s43_cols: List[List[str]] = []
    s44_content: str = ""
    storage_path: str = ""


@router.post("/upload-media")
async def upload_detection_media(file: UploadFile = File(...)):
    if file.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(status_code=400, detail=f"不支持的文件类型: {file.content_type}")
    ext = Path(file.filename).suffix or ".png"
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}{ext}"
    async with aiofiles.open(save_path, "wb") as out:
        await out.write(await file.read())
    return {"file_id": f"{file_id}{ext}", "url": f"/static/uploads/{file_id}{ext}"}


@router.post("/generate")
async def generate_detection_report(payload: DetectionPayload, fmt: str = Query("docx"),
                                     db: AsyncSession = Depends(get_db)):
    uid = uuid.uuid4().hex[:8]
    docx_path = pdf_path = docx_url = pdf_url = None

    if fmt in ("docx", "both"):
        filename = f"det_{uid}.docx"
        out_path = PERMANENT_REPORTS_DIR / filename
        render_detection_report(str(out_path), payload.model_dump())
        docx_path = str(out_path)
        docx_url = f"/api/v1/detection/download/{filename}"

    if fmt in ("pdf", "both"):
        filename_pdf = f"det_{uid}.pdf"
        out_path_pdf = PERMANENT_REPORTS_DIR / filename_pdf
        try:
            render_detection_pdf(str(out_path_pdf), payload.model_dump())
            pdf_path = str(out_path_pdf)
            pdf_url = f"/api/v1/detection/download/{filename_pdf}"
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"PDF生成失败: {e}")

    record = DetectionReportDB(
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

    return {
        "status": "success",
        "report_id": record.id,
        "download_url": docx_url or pdf_url,
        "docx_url": docx_url,
        "pdf_url": pdf_url,
    }


@router.get("/history")
async def get_detection_history(user_id: str = Query("default_user"), page: int = 1, page_size: int = 20,
                                 db: AsyncSession = Depends(get_db)):
    offset = (page - 1) * page_size
    total = await db.scalar(
        select(func.count()).select_from(DetectionReportDB).where(DetectionReportDB.user_id == user_id)
    )
    result = await db.execute(
        select(DetectionReportDB).where(DetectionReportDB.user_id == user_id)
        .order_by(desc(DetectionReportDB.created_at)).offset(offset).limit(page_size)
    )
    reports = result.scalars().all()
    return {
        "total": total, "page": page, "page_size": page_size,
        "reports": [
            {
                "id": r.id, "report_title": r.report_title, "doc_number": r.doc_number,
                "project_name": r.project_name, "created_at": r.created_at,
                "docx_url": f"/api/v1/detection/download/{Path(r.docx_file_path).name}" if r.docx_file_path else None,
                "pdf_url": f"/api/v1/detection/download/{Path(r.pdf_file_path).name}" if r.pdf_file_path else None,
            } for r in reports
        ]
    }


@router.get("/download/{filename}")
async def download_detection_report(filename: str):
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
async def preview_detection_report(payload: DetectionPayload):
    try:
        html_content = render_detection_preview(payload.model_dump())
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return {"html": html_content}
