from contextlib import asynccontextmanager
from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path

load_dotenv()

from database import init_db
from routers.report import router as report_router
from routers.draft import router as draft_router


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    yield


app = FastAPI(
    title="蛋白发酵报告自动生成系统",
    description="支持数据库持久化与历史配置回显的蛋白发酵实验报告生成平台",
    version="2.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://192.168.1.150", "http://localhost", "http://127.0.0.1"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

Path("static/uploads").mkdir(parents=True, exist_ok=True)
Path("static/reports").mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

app.include_router(report_router)
app.include_router(draft_router)


@app.get("/health")
async def health():
    return {"status": "ok"}
