"""
审计日志客户端：将蛋白发酵后端产生的审计记录转发到化学报告后端的统一 audit.db。
化学报告后端地址可通过环境变量 AUDIT_BACKEND_URL 覆盖。
"""
import os
import httpx
from datetime import datetime

AUDIT_BACKEND_URL = os.getenv("AUDIT_BACKEND_URL", "http://127.0.0.1:8001")


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
        pass  # 审计失败不阻断主流程
