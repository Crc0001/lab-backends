from typing import Optional

from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel, ConfigDict
from sqlalchemy import or_, select
from sqlalchemy.ext.asyncio import AsyncSession

from database import get_db
from models import EquipmentDB


router = APIRouter(prefix="/api/v1/equipment", tags=["equipment"])


class EquipmentItem(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    sequence: int
    equipment_code: Optional[str] = None
    instrument_name: str
    model: Optional[str] = None
    manufacturer: Optional[str] = None
    manufacture_date: Optional[str] = None
    installation_date: Optional[str] = None
    verification_type: Optional[str] = None
    purchase_date: Optional[str] = None
    earliest_calibration_date: Optional[str] = None
    latest_calibration_date: Optional[str] = None
    next_calibration_date: Optional[str] = None
    next_verification_date: Optional[str] = None
    verification_cycle: Optional[str] = None
    remarks: Optional[str] = None


@router.get("", response_model=list[EquipmentItem])
async def list_equipment(
    q: Optional[str] = Query(default=None, max_length=100),
    limit: int = Query(default=200, ge=1, le=500),
    db: AsyncSession = Depends(get_db),
):
    stmt = select(EquipmentDB)
    if q and (term := q.strip()):
        pattern = f"%{term}%"
        stmt = stmt.where(
            or_(
                EquipmentDB.instrument_name.ilike(pattern),
                EquipmentDB.equipment_code.ilike(pattern),
                EquipmentDB.model.ilike(pattern),
                EquipmentDB.manufacturer.ilike(pattern),
            )
        )
    stmt = stmt.order_by(
        EquipmentDB.instrument_name,
        EquipmentDB.equipment_code,
        EquipmentDB.sequence,
    ).limit(limit)
    result = await db.execute(stmt)
    return list(result.scalars())
