import uuid
import os
from pathlib import Path
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Query
from fastapi.responses import FileResponse, HTMLResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, desc
import aiofiles

from database import get_db
from models import (
    FermentationReportDB,
    ReportGenerationPayload,
    ReportGenerationResponse,
    EchoConfigResponse,
    UploadMediaResponse,
    MaterialPrep,
    FermentationInduction,
    ExpressionForm,
    PurificationResults,
    ReportPayloadV2,
)
from services.document_renderer import render_report
from services.document_renderer_v2 import render_report_v2, render_preview_html, render_pdf
from audit_client import log_action

router = APIRouter(prefix="/api/v1/report", tags=["report"])

UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "static/uploads"))
REPORTS_DIR = Path(os.environ.get("REPORTS_DIR", "static/reports"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/tiff", "image/bmp"}
ALLOWED_TEXT_TYPES = {"text/plain", "application/octet-stream"}


@router.post("/upload-txt", response_model=UploadMediaResponse)
async def upload_txt(file: UploadFile = File(...)):
    """上传附录 TXT 文件，返回文件ID和服务器存储路径。"""
    ext = Path(file.filename).suffix.lower()
    if ext != ".txt":
        raise HTTPException(status_code=400, detail="仅支持 .txt 文件")
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}.txt"
    async with aiofiles.open(save_path, "wb") as out:
        await out.write(await file.read())
    return UploadMediaResponse(
        file_id=f"{file_id}.txt",
        url=str(save_path),  # 返回服务器绝对路径，供后端 render 使用
    )


@router.get("/latest-config", response_model=EchoConfigResponse)
async def get_latest_config(
    user_id: str = Query("default_user", description="用户ID"),
    db: AsyncSession = Depends(get_db),
):
    """获取该用户最近一次成功保存的实验配置，用于表单数据回显。"""
    result = await db.execute(
        select(FermentationReportDB)
        .where(FermentationReportDB.user_id == user_id)
        .order_by(desc(FermentationReportDB.created_at))
        .limit(1)
    )
    record = result.scalar_one_or_none()

    if record is None:
        raise HTTPException(status_code=404, detail="该用户暂无历史实验记录")

    return EchoConfigResponse(
        materials=MaterialPrep(**record.materials_data),
        fermentation=FermentationInduction(**record.fermentation_data),
        expression=ExpressionForm(**record.expression_data),
        purification=PurificationResults(**record.purification_data),
    )


@router.post("/upload-media", response_model=UploadMediaResponse)
async def upload_media(file: UploadFile = File(...)):
    """上传 SDS-PAGE 电泳图或质谱图，返回文件ID和访问URL。"""
    if file.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的文件类型: {file.content_type}，仅支持 JPEG/PNG/TIFF/BMP"
        )

    ext = Path(file.filename).suffix or ".png"
    file_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{file_id}{ext}"

    async with aiofiles.open(save_path, "wb") as out:
        content = await file.read()
        await out.write(content)

    return UploadMediaResponse(
        file_id=f"{file_id}{ext}",
        url=f"/static/uploads/{file_id}{ext}",
    )


@router.post("/generate", response_model=ReportGenerationResponse)
async def generate_report(
    payload: ReportGenerationPayload,
    db: AsyncSession = Depends(get_db),
):
    """接收完整实验数据，持久化至数据库并生成 Word 报告。"""
    record = FermentationReportDB(
        report_title=payload.report_title,
        user_id=payload.user_id,
        materials_data=payload.materials.model_dump(),
        fermentation_data=payload.fermentation.model_dump(),
        expression_data=payload.expression.model_dump(),
        purification_data=payload.purification.model_dump(),
        sds_page_file_id=payload.sds_page_file_id,
        mass_spec_file_id=payload.mass_spec_file_id,
    )
    db.add(record)
    await db.flush()
    await db.refresh(record)

    report_filename = f"rep_{record.id}.docx"
    report_path = REPORTS_DIR / report_filename

    sds_path = UPLOAD_DIR / payload.sds_page_file_id if payload.sds_page_file_id else None
    mass_path = UPLOAD_DIR / payload.mass_spec_file_id if payload.mass_spec_file_id else None

    try:
        render_report(
            output_path=str(report_path),
            payload=payload,
            sds_page_path=str(sds_path) if sds_path and sds_path.exists() else None,
            mass_spec_path=str(mass_path) if mass_path and mass_path.exists() else None,
        )
    except Exception as e:
        await db.rollback()
        raise HTTPException(status_code=500, detail=str(e))

    try:
        await db.commit()
    except Exception as e:
        report_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"数据库保存失败: {e}")

    await log_action(payload.user_id, f"生成蛋白发酵报告(v1) - {payload.report_title}")

    return ReportGenerationResponse(
        status="success",
        report_id=record.id,
        download_url=f"/api/v1/report/download/{report_filename}",
    )


@router.get("/download/{filename}")
async def download_report(filename: str):
    """下载已生成的 Word 或 PDF 报告。"""
    # 防路径穿越：只允许纯文件名
    if filename != Path(filename).name or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    report_path = REPORTS_DIR / filename
    # 确认最终路径仍在报告目录内
    if not report_path.resolve().is_relative_to(REPORTS_DIR.resolve()):
        raise HTTPException(status_code=400, detail="非法文件名")
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="报告文件不存在")
    media_type = "application/pdf" if filename.endswith(".pdf") else \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path=str(report_path), media_type=media_type, filename=filename)


@router.post("/generate-v2", response_model=ReportGenerationResponse)
async def generate_report_v2(payload: ReportPayloadV2, fmt: str = Query("docx", description="docx 或 pdf"),
               db: AsyncSession = Depends(get_db)):
    """新架构：接收表格数据，生成 Word 或 PDF 报告，并持久化到数据库。"""
    import uuid as _uuid
    uid = _uuid.uuid4().hex[:8]
    docx_path = pdf_path = docx_url = pdf_url = None

    docx_filename = f"rep_{uid}.docx"
    docx_out = REPORTS_DIR / docx_filename
    try:
        render_report_v2(str(docx_out), payload)
        docx_path = str(docx_out)
        docx_url = f"/api/v1/report/download/{docx_filename}"
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    if fmt == "pdf":
        pdf_filename = f"rep_{uid}.pdf"
        pdf_out = REPORTS_DIR / pdf_filename
        try:
            render_pdf(str(pdf_out), payload)
            pdf_path = str(pdf_out)
            pdf_url = f"/api/v1/report/download/{pdf_filename}"
        except Exception as e:
            docx_out.unlink(missing_ok=True)
            raise HTTPException(status_code=500, detail=f"PDF生成失败: {e}")

    record = FermentationReportDB(
        report_title=payload.report_title,
        user_id=payload.user_id,
        materials_data=payload.model_dump(),
        fermentation_data={
            "approver_signatures": payload.approver_signatures,
            "approver_signature_audits": [e.model_dump() for e in payload.approver_signature_audits],
        },
        expression_data={},
        purification_data={},
        form_data=payload.model_dump(),
        docx_file_path=docx_path,
        pdf_file_path=pdf_path,
    )
    db.add(record)
    try:
        await db.commit()
    except Exception as e:
        docx_out.unlink(missing_ok=True)
        if fmt == "pdf":
            pdf_out.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"数据库保存失败: {e}")
    await db.refresh(record)

    await log_action(payload.user_id, f"生成蛋白发酵报告(v2) - {payload.report_title}")

    download_url = pdf_url or docx_url
    return ReportGenerationResponse(status="success", report_id=record.id, download_url=download_url)


@router.get("/history")
async def get_history(user_id: str = Query("default_user"), page: int = 1, page_size: int = 20,
                      db: AsyncSession = Depends(get_db)):
    from sqlalchemy import func
    offset = (page - 1) * page_size
    total = await db.scalar(
        select(func.count()).select_from(FermentationReportDB).where(FermentationReportDB.user_id == user_id)
    )
    result = await db.execute(
        select(FermentationReportDB).where(FermentationReportDB.user_id == user_id)
        .order_by(desc(FermentationReportDB.created_at)).offset(offset).limit(page_size)
    )
    reports = result.scalars().all()
    return {
        "total": total, "page": page, "page_size": page_size,
        "reports": [
            {
                "id": r.id, "report_title": r.report_title,
                "created_at": r.created_at,
                "docx_url": f"/api/v1/report/download/{Path(r.docx_file_path).name}" if r.docx_file_path else None,
                "pdf_url": f"/api/v1/report/download/{Path(r.pdf_file_path).name}" if r.pdf_file_path else None,
            } for r in reports
        ]
    }


@router.post("/preview")
async def preview_report(payload: ReportPayloadV2):
    """返回 HTML 预览页面。"""
    try:
        html_content = render_preview_html(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return HTMLResponse(content=html_content)





