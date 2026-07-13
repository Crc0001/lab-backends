import hashlib
import hmac
import os
import uuid
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import AsyncSessionLocal, get_db
from models import UserDB, UserSessionDB

router = APIRouter(prefix="/api/v1/auth", tags=["auth"])

_DEFAULT_USERS = {
    "admin": ["*"],
    "zym": ["protein"],
    "ltf": ["chemistry"],
    "dwh": ["microbiology"],
    "ymx": ["gcms", "hplc"],
}


def _hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or os.urandom(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=16384, r=8, p=1)
    return f"{salt.hex()}${digest.hex()}"


def _verify_password(password: str, stored: str) -> bool:
    try:
        salt_hex, expected = stored.split("$", 1)
        actual = _hash_password(password, bytes.fromhex(salt_hex)).split("$", 1)[1]
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


async def seed_default_users() -> None:
    async with AsyncSessionLocal() as db:
        for username, permissions in _DEFAULT_USERS.items():
            result = await db.execute(select(UserDB).where(UserDB.username == username))
            if result.scalar_one_or_none() is None:
                db.add(UserDB(
                    username=username,
                    password_hash=_hash_password("123456"),
                    permissions=permissions,
                    is_active=1,
                ))
        await db.commit()


class LoginRequest(BaseModel):
    username: str
    password: str


class SessionStatusRequest(BaseModel):
    username: str
    session_id: str


@router.post("/login")
async def login(req: LoginRequest, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(UserDB).where(UserDB.username == req.username))
    user = result.scalar_one_or_none()
    if user is None or not user.is_active or not _verify_password(req.password, user.password_hash):
        raise HTTPException(status_code=401, detail="用户名或密码错误")
    session_id = uuid.uuid4().hex
    result = await db.execute(select(UserSessionDB).where(UserSessionDB.username == req.username))
    session = result.scalar_one_or_none()
    if session is None:
        db.add(UserSessionDB(username=req.username, session_id=session_id))
    else:
        session.session_id = session_id
    await db.commit()
    return {
        "status": "ok",
        "username": req.username,
        "session_id": session_id,
        "permissions": user.permissions,
    }


@router.post("/session/status")
async def session_status(req: SessionStatusRequest, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(UserSessionDB).where(UserSessionDB.username == req.username))
    session = result.scalar_one_or_none()
    active = session is not None and session.session_id == req.session_id
    return {"active": active, "code": "OK" if active else "SESSION_REPLACED"}


@router.post("/logout")
async def logout(req: SessionStatusRequest, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(UserSessionDB).where(UserSessionDB.username == req.username))
    session = result.scalar_one_or_none()
    if session is not None and session.session_id == req.session_id:
        await db.delete(session)
        await db.commit()
    return {"status": "ok"}
