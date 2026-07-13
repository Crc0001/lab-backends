from fastapi import APIRouter, Depends, Query
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, func
from models import InspectionReportDB, ValidationReportDB, ReportHistoryItem, ReportHistoryResponse
from database import get_db

router = APIRouter(prefix="/history")

@router.get("/inspection", response_model=ReportHistoryResponse)
async def inspection_history(
    limit: int = Query(50, le=200),
    offset: int = Query(0),
    db: AsyncSession = Depends(get_db),
):
    total = (await db.execute(select(func.count()).select_from(InspectionReportDB))).scalar()
    rows = (await db.execute(
        select(InspectionReportDB).order_by(InspectionReportDB.id.desc()).limit(limit).offset(offset)
    )).scalars().all()
    return {"total": total, "reports": [ReportHistoryItem.model_validate(r) for r in rows]}

@router.get("/validation", response_model=ReportHistoryResponse)
async def validation_history(
    limit: int = Query(50, le=200),
    offset: int = Query(0),
    db: AsyncSession = Depends(get_db),
):
    total = (await db.execute(select(func.count()).select_from(ValidationReportDB))).scalar()
    rows = (await db.execute(
        select(ValidationReportDB).order_by(ValidationReportDB.id.desc()).limit(limit).offset(offset)
    )).scalars().all()
    return {"total": total, "reports": [ReportHistoryItem.model_validate(r) for r in rows]}
