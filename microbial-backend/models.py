from datetime import datetime
from typing import List, Optional
from pydantic import BaseModel
from sqlalchemy import Column, Integer, String, DateTime, JSON
from database import Base


class MicrobialReportDB(Base):
    __tablename__ = "microbial_reports"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    project_name = Column(String(255), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=datetime.utcnow)
    form_data = Column(JSON, nullable=False)
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)
    docx_download_url = Column(String(500), nullable=True)
    pdf_download_url = Column(String(500), nullable=True)


class DraftDB(Base):
    __tablename__ = "drafts"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    kind = Column(String(50), nullable=False, index=True)
    draft_key = Column(String(100), nullable=False, index=True)
    user_id = Column(String(50), index=True, default="default_user")
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    form_data = Column(JSON, nullable=False)


class PersonnelRow(BaseModel):
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


class ResultRowData(BaseModel):
    seq: str = ""
    name: str = ""
    extra_cols: List[str] = []
    image_file_id: Optional[str] = None


class ReportPayload(BaseModel):
    report_title: str = "微生物实验报告"
    user_id: str = "default_user"
    doc_number: str = ""
    security_level: str = ""
    project_name: str = ""
    first_page_footer: str = ""

    personnel: List[PersonnelRow] = []
    approver_signatures: Optional[List[Optional[str]]] = None
    approver_signature_audits: List[SignatureAuditEntry] = []

    # 一、材料试剂与仪器
    strains_headers: Optional[List[str]] = None       # 菌株信息
    strains: List[List[str]] = []
    media_headers: Optional[List[str]] = None         # 培养基
    media: List[List[str]] = []
    reagents_headers: Optional[List[str]] = None
    reagents: List[List[str]] = []
    instruments_headers: Optional[List[str]] = None
    instruments: List[List[str]] = []
    consumables_headers: Optional[List[str]] = None
    consumables: List[List[str]] = []

    # 二、操作步骤（3个固定 + 动态）
    procedure_headers: Optional[List[List[str]]] = None
    procedures: List[List[List[str]]] = []
    dynamic_section_titles: List[str] = []
    dynamic_section_headers: List[List[str]] = []
    dynamic_section_tables: List[List[List[str]]] = []

    # 三、实验结果与分析
    result_extra_headers: List[str] = []
    results: List[ResultRowData] = []

    # 四、实验结果（通用表格/文本模式，来自Flutter前端Step4）
    results_mode: str = "table"          # 'table' | 'text'
    results_text_content: str = ""
    results_headers: List[str] = []
    results_rows: List[List[str]] = []

    # 四、实验结论
    conclusion_table_headers: Optional[List[str]] = None
    conclusion_table_rows: List[List[str]] = []

    # 附录：菌落计数/OD标准曲线
    qc_data_headers: Optional[List[str]] = None
    qc_concentrations: Optional[List[str]] = None     # X轴（稀释度或浓度）
    qc_od_values: Optional[List[str]] = None          # Y轴（OD600或菌落数）
    qc_linear_equation: Optional[str] = None
    qc_correlation_coeff: Optional[float] = None
    qc_conclusion_headers: Optional[List[str]] = None
    qc_conclusion_row_labels: Optional[List[str]] = None
    qc_conclusion_rows: Optional[List[List[str]]] = None

    appendix_file_paths: List[str] = []


class UploadMediaResponse(BaseModel):
    file_id: str
    url: str


class ReportGenerationResponse(BaseModel):
    status: str
    download_url: str
    report_id: int = 0
    docx_url: str | None = None
    pdf_url: str | None = None
