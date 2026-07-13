import datetime as dt
from typing import List, Optional
from pydantic import BaseModel
from sqlalchemy import Column, Integer, String, DateTime, JSON
from database import Base


class GcmsReportDB(Base):
    """气质报告数据库模型"""
    __tablename__ = "gcms_reports"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    project_name = Column(String(255), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class DraftDB(Base):
    __tablename__ = "drafts"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    kind = Column(String(50), nullable=False, index=True)
    draft_key = Column(String(100), nullable=False, index=True)
    user_id = Column(String(50), index=True, default="default_user")
    updated_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc), onupdate=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)


class SignatureAuditEntry(BaseModel):
    previous_file_id: Optional[str] = None
    new_file_id: Optional[str] = None
    reason: str = ""
    changed_at: str = ""


class GcmsPayload(BaseModel):
    report_title: str = ""
    doc_number: str = ""
    security_level: str = ""
    project_name: str = ""
    user_id: str = "default_user"
    personnel: List[dict] = []
    approver_signatures: Optional[List[Optional[str]]] = None
    approver_signature_audits: List[SignatureAuditEntry] = []
    instruments: List[List[str]] = []
    instrument_headers: List[str] = []
    reagents: List[List[str]] = []
    references: List[List[str]] = []
    s21_rows: List[List[str]] = []
    s22_rows: List[List[str]] = []
    s23_rows: List[List[str]] = []
    s24_rows: List[List[str]] = []
    s25_rows: List[List[str]] = []
    s26_rows: List[List[str]] = []
    s27_rows: List[List[str]] = []
    gcms_param_values: List[str] = []
    gc_methods: List[dict] = []
    s41_rows: List[List[str]] = []
    s42_rows: List[List[str]] = []
    s43_headers: List[str] = []
    s43_cols: List[List[str]] = []
    s44_content: str = ""
    storage_path: str = ""
