from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from contextlib import asynccontextmanager
from dotenv import load_dotenv

load_dotenv()

from routers.report import router as report_router
from routers.detection import router as detection_router
from routers.audit import router as audit_router
from routers.auth import router as auth_router, seed_default_users
from routers.draft import router as draft_router
from routers.equipment import router as equipment_router
from audit_db import init_db as init_audit_db
from database import init_db as init_reports_db
from equipment_seed import seed_equipment


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    init_audit_db()
    await init_reports_db()
    await seed_equipment()
    await seed_default_users()
    yield
    # Shutdown (if needed)


app = FastAPI(title="化学检测报告自动生成系统", version="1.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://192.168.31.200", "http://localhost", "http://127.0.0.1"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

Path("static/uploads").mkdir(parents=True, exist_ok=True)
Path("static/reports").mkdir(parents=True, exist_ok=True)
Path("reports").mkdir(parents=True, exist_ok=True)  # 永久保存报告的目录
app.mount("/static", StaticFiles(directory="static"), name="static")

app.include_router(report_router)
app.include_router(detection_router)
app.include_router(audit_router)
app.include_router(auth_router)
app.include_router(draft_router)
app.include_router(equipment_router)


@app.get("/health")
async def health():
    return {"status": "ok"}
