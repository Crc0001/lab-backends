from fastapi import APIRouter, Query, Header, HTTPException
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel
from audit_db import insert_log, query_logs
import os

router = APIRouter(prefix="/api/v1/audit", tags=["audit"])

_AUDIT_SECRET = os.getenv("AUDIT_SECRET", "")


def _check_secret(x_audit_secret: str | None):
    if _AUDIT_SECRET and x_audit_secret != _AUDIT_SECRET:
        raise HTTPException(status_code=403, detail="Forbidden")


class LogEntry(BaseModel):
    time: str
    user: str
    type: str
    desc: str


@router.post("/log", status_code=201)
async def post_log(entry: LogEntry, x_audit_secret: str | None = Header(default=None)):
    _check_secret(x_audit_secret)
    await run_in_threadpool(insert_log, entry.time, entry.user, entry.type, entry.desc)
    return {"status": "ok"}


@router.get("/logs")
async def get_logs(
    user: str | None = Query(None),
    limit: int = Query(200, le=1000),
    offset: int = Query(0),
    x_audit_secret: str | None = Header(default=None),
):
    _check_secret(x_audit_secret)
    return await run_in_threadpool(query_logs, user, limit, offset)
