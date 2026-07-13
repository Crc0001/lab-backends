from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import get_db
from models import DraftDB


router = APIRouter(prefix="/draft")


class DraftPayload(BaseModel):
    user_id: str
    data: dict


@router.get("/{kind}/{draft_key}")
async def get_draft(kind: str, draft_key: str, user_id: str = Query(..., min_length=1), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(DraftDB).where(DraftDB.kind == kind, DraftDB.draft_key == draft_key, DraftDB.user_id == user_id))
    draft = result.scalar_one_or_none()
    if draft is None:
        return {"exists": False, "data": None, "updated_at": None}
    return {"exists": True, "data": draft.form_data, "updated_at": draft.updated_at.isoformat()}


@router.put("/{kind}/{draft_key}")
async def save_draft(kind: str, draft_key: str, payload: DraftPayload, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(DraftDB).where(DraftDB.kind == kind, DraftDB.draft_key == draft_key, DraftDB.user_id == payload.user_id))
    draft = result.scalar_one_or_none()
    if draft is None:
        draft = DraftDB(kind=kind, draft_key=draft_key, user_id=payload.user_id, form_data=payload.data)
        db.add(draft)
    else:
        draft.form_data = payload.data
    await db.commit()
    await db.refresh(draft)
    return {"saved": True, "updated_at": draft.updated_at.isoformat()}
