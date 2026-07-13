import os
import httpx
from datetime import datetime

AUDIT_BACKEND_URL = os.getenv("AUDIT_BACKEND_URL", "http://192.168.31.200:8001")

async def log_action(user: str, action: str, log_type: str = "action"):
    entry = {
        "time": datetime.now().isoformat(),
        "user": user,
        "type": log_type,
        "desc": action,
    }
    try:
        async with httpx.AsyncClient(timeout=5) as client:
            await client.post(f"{AUDIT_BACKEND_URL}/api/v1/audit/log", json=entry)
    except Exception:
        pass
