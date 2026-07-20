import datetime as dt
from pydantic import BaseModel
from typing import List, Optional
from sqlalchemy import Column, Integer, String, DateTime, JSON
from database import Base


class PersonnelItem(BaseModel):
    editor: str = ""
    reviewer: str = ""
    approver: str = ""
    date: str = ""
    version: str = ""


class SignatureAuditEntry(BaseModel):
    previous_file_id: Optional[str] = None
    new_file_id: Optional[str] = None
    reason: str = ""
    changed_at: str = ""


# ── Chapter 4 models ──────────────────────────────────────────────────────────

class CompoundItem(BaseModel):
    name: str = ""
    weight: str = ""       # 4.1 称量量
    s44_desc: str = ""     # 4.4 对照品储备液描述


class TableDef(BaseModel):
    headers: List[str] = []
    rows: List[List[str]] = []     # 用户填写的实际数据行


class LinearityTableDef(BaseModel):
    col_headers: List[str] = []          # 定量限均值 + S1~S5
    concentrations: List[str] = []
    peak_areas: List[str] = []
    equation: str = ""
    r2: str = ""


class RecoveryTableDef(BaseModel):
    headers: List[str] = []
    group_label: str = ""
    group_name: str = ""
    added_amount: str = ""
    sample_content: str = ""
    avg_recovery: str = ""
    rsd: str = ""
    sub_rows: List[List[str]] = []
    rows: List[List[str]] = []


class Chapter4Data(BaseModel):
    compounds: List[CompoundItem] = []
    # 4.1
    s41_items: List[str] = []
    s41_table_title: str = "表4.1-1 亚硝胺杂质专属性实验结果"
    s41_table: TableDef = TableDef()
    s41_conclusion: str = ""
    # 4.2
    s42_items: List[str] = []
    s42_loq_titles: List[str] = []
    s42_loq_tables: List[TableDef] = []
    s42_lod_title: str = "表4.2 亚硝胺杂质检测限实验结果"
    s42_lod_table: TableDef = TableDef()
    s42_conclusion: str = ""
    # 4.3
    s43_items: List[str] = []
    s43_titles: List[str] = []
    s43_tables: List[LinearityTableDef] = []
    s43_conclusion: str = ""
    # 4.4
    s44_items: List[str] = []
    s44_titles: List[str] = []
    s44_tables: List[RecoveryTableDef] = []
    s44_conclusion: str = ""
    # 4.5
    s45_items: List[str] = []
    s45_table_title: str = "表4.5 亚硝胺杂质检测结果"
    s45_table: TableDef = TableDef()
    s45_conclusion: str = ""


class LinearityInput(BaseModel):
    concentrations: List[float]
    peak_areas: List[float]


class LinearityResult(BaseModel):
    equation: str
    r: str
    r2: str


class ReportPayload(BaseModel):
    chapter4: Chapter4Data = Chapter4Data()
    report_title: str = ""
    doc_number: str = ""
    security_level: str = ""
    project_name: str = ""
    user_id: str = "default_user"
    personnel: List[PersonnelItem] = []
    approver_signatures: Optional[List[Optional[str]]] = None
    approver_signature_audits: List[SignatureAuditEntry] = []
    detection_method: str = ""
    impurities: List[List[str]] = []
    # 液相条件
    hplc_column: str = ""
    hplc_column_det: str = ""
    hplc_temp: str = ""
    hplc_temp_det: str = ""
    hplc_mobile_phase_desc: str = ""
    mobile_phase_rows: List[List[str]] = []
    hplc_flow_rate: str = ""
    hplc_flow_rate_det: str = ""
    hplc_inj_vol: str = ""
    hplc_inj_vol_det: str = ""
    hplc_inj_temp: str = ""
    hplc_inj_temp_det: str = ""
    # 质谱条件
    ms_instrument: str = ""
    ms_instrument_det: str = ""
    ms_ion_source: str = ""
    ms_ion_source_det: str = ""
    ms_scan_mode: str = ""
    ms_scan_mode_det: str = ""
    monitoring_rows: List[List[str]] = []
    ion_source_rows: List[List[str]] = []
    ion_source_det: str = ""
    # 底部单行
    limit: str = ""
    limit_det: str = ""
    switch_valve: str = ""
    switch_valve_det: str = ""
    detection_time: str = ""
    detection_time_det: str = ""
    detection_note: str = ""
    # 方法学验证
    validation_rows: List[List[str]] = []
    # 仪器与试剂
    instruments: List[List[str]] = []
    reagents: List[List[str]] = []
    references: List[List[str]] = []
    # 5. 修订历史
    revision_rows: List[List[str]] = []
    # 附表1：人员培训表
    training_content: str = ""
    training_instructor: str = ""
    training_date: str = ""
    training_attendees: str = ""
    # 附表2：偏差处理表
    deviation_rows: List[List[str]] = []


class ReportResponse(BaseModel):
    status: str
    download_url: str


# ==================== SQLAlchemy 数据库模型 ====================

class DetectionReportDB(Base):
    """化学检测报告数据库模型"""
    __tablename__ = "detection_reports"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    project_name = Column(String(255), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))
    
    # 存储完整表单数据 JSON
    form_data = Column(JSON, nullable=False)
    
    # 生成的文件路径
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class ValidationReportDB(Base):
    """化学验证报告数据库模型"""
    __tablename__ = "validation_reports"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    project_name = Column(String(255), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))
    
    # 存储完整表单数据 JSON
    form_data = Column(JSON, nullable=False)
    
    # 生成的文件路径
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class UserSessionDB(Base):
    __tablename__ = "user_sessions"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    username = Column(String(50), nullable=False, unique=True, index=True)
    session_id = Column(String(100), nullable=False)
    updated_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc), onupdate=lambda: dt.datetime.now(dt.timezone.utc))


class UserDB(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    username = Column(String(50), nullable=False, unique=True, index=True)
    password_hash = Column(String(255), nullable=False)
    permissions = Column(JSON, nullable=False, default=list)
    is_active = Column(Integer, nullable=False, default=1)
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))


class EquipmentDB(Base):
    __tablename__ = "equipment"

    sequence = Column(Integer, primary_key=True, autoincrement=False)
    equipment_code = Column(String(100), nullable=True, index=True)
    instrument_name = Column(String(255), nullable=False, index=True)
    model = Column(String(255), nullable=True)
    manufacturer = Column(String(255), nullable=True)
    manufacture_date = Column(String(50), nullable=True)
    installation_date = Column(String(50), nullable=True)
    verification_type = Column(String(100), nullable=True)
    purchase_date = Column(String(50), nullable=True)
    earliest_calibration_date = Column(String(50), nullable=True)
    latest_calibration_date = Column(String(50), nullable=True)
    next_calibration_date = Column(String(50), nullable=True)
    next_verification_date = Column(String(50), nullable=True)
    verification_cycle = Column(String(100), nullable=True)
    remarks = Column(String(500), nullable=True)


class DraftDB(Base):
    __tablename__ = "drafts"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    kind = Column(String(50), nullable=False, index=True)
    draft_key = Column(String(100), nullable=False, index=True)
    user_id = Column(String(50), index=True, default="default_user")
    updated_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc), onupdate=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)


# ==================== 历史记录响应模型 ====================

class ReportHistoryItem(BaseModel):
    id: int
    report_title: str
    doc_number: Optional[str]
    project_name: Optional[str]
    created_at: dt.datetime
    docx_file_path: Optional[str]
    pdf_file_path: Optional[str]

    class Config:
        from_attributes = True


class ReportHistoryResponse(BaseModel):
    total: int
    reports: List[ReportHistoryItem]
