import datetime as dt
from pydantic import BaseModel
from typing import List, Optional
from sqlalchemy import Column, Integer, String, DateTime, JSON
from database import Base

# ── 通用 ──────────────────────────────────────────────────────────────────────

class InstrumentRow(BaseModel):
    name: str = ""
    model: str = ""
    number: str = ""
    manufacturer: str = ""
    expiry: str = ""

class ReagentRow(BaseModel):
    name: str = ""
    batch: str = ""
    grade: str = ""
    manufacturer: str = ""

class ApprovalRow(BaseModel):
    role: str = ""       # 起草/审核/批准
    position: str = ""
    signature_date: str = ""

# ── 检验记录 ──────────────────────────────────────────────────────────────────

class InspectionItem(BaseModel):
    """单个检验项目（性状、有关物质、含量测定…）"""
    title: str = ""                        # 项目名称
    temperature: str = ""
    humidity: str = ""
    instruments: List[List[str]] = []      # 每行：[名称, 型号, 编号, 厂家, 有效期]
    reagents: List[List[str]] = []         # 每行：[名称, 批号, 级别/含量, 来源]
    standard: str = ""                     # 标准
    method: str = ""                       # 检验方法/色谱条件/溶液配制
    procedure: str = ""                    # 操作过程
    calculation: str = ""                  # 计算过程及结果
    result: str = ""                       # 结果
    conclusion: str = "符合规定"
    inspector: str = ""
    inspector_date: str = ""
    reviewer: str = ""
    reviewer_date: str = ""

class InspectionReport(BaseModel):
    report_type: str = "api"               # api | solid | liquid
    # 基本信息
    product_name: str = ""
    doc_number: str = ""
    batch_number: str = ""
    quantity: str = ""
    specification: str = ""
    source: str = ""
    request_date: str = ""
    basis: str = ""                        # 检验依据
    sample_nature: str = ""               # 样品性质（固体制剂用）
    company: str = "石家庄凯瑞德医药科技发展有限公司"
    # 审核批准
    approvals: List[ApprovalRow] = []
    # 检验项目列表（动态）
    items: List[InspectionItem] = []
    # 异常情况
    abnormal: str = ""

# ── 验证报告 ──────────────────────────────────────────────────────────────────

class ValidationConditionRow(BaseModel):
    condition: str = ""    # 检测条件
    basis: str = ""        # 确定依据

class ValidationConditions(BaseModel):
    column: ValidationConditionRow = ValidationConditionRow()
    temp: ValidationConditionRow = ValidationConditionRow()
    mobile_phase: ValidationConditionRow = ValidationConditionRow()
    wavelength: ValidationConditionRow = ValidationConditionRow()
    flow_rate: ValidationConditionRow = ValidationConditionRow()
    injection_vol: ValidationConditionRow = ValidationConditionRow()
    limit: ValidationConditionRow = ValidationConditionRow()
    run_time: ValidationConditionRow = ValidationConditionRow()

class ValidationSection(BaseModel):
    """通用验证章节：段落列表 + 动态表格 + 结论"""
    paragraphs: List[str] = []
    table_headers: List[str] = []
    table_rows: List[List[str]] = []
    conclusion: str = ""

class ValidationReport(BaseModel):
    doc_number: str = ""
    project_name: str = ""
    author: str = ""
    author_date: str = ""
    reviewer: str = ""
    reviewer_date: str = ""
    approver: str = ""
    approver_date: str = ""
    purpose: str = ""
    scope: str = ""
    basis: str = ""
    method_summary: str = ""
    # 检测条件
    conditions: ValidationConditions = ValidationConditions()
    validation_summary: List[List[str]] = []  # [项目, 可接受标准, 结果]
    # 仪器与试剂
    instruments: List[List[str]] = []      # [名称, 编号, 型号, 厂家]
    reagents: List[List[str]] = []         # [名称, 批号, 级别, 厂家]
    references: List[List[str]] = []       # [名称, 批号, 含量, 厂家]
    samples: List[List[str]] = []          # [名称, 批号, 厂家]
    excipients: List[List[str]] = []       # [名称, 批号, 厂家]
    # 8个验证章节（固定）
    s51_specificity: ValidationSection = ValidationSection()
    s52_linearity: ValidationSection = ValidationSection()
    s53_loq_lod: ValidationSection = ValidationSection()
    s54_repeatability: ValidationSection = ValidationSection()
    s55_intermediate: ValidationSection = ValidationSection()
    s56_accuracy: ValidationSection = ValidationSection()
    s57_stability: ValidationSection = ValidationSection()
    s58_robustness: ValidationSection = ValidationSection()
    # 修订历史
    revisions: List[List[str]] = []        # [编号, 更改原因, 生效日期]


# ── ORM 模型 ──────────────────────────────────────────────────────────────────

class InspectionReportDB(Base):
    __tablename__ = "inspection_reports"
    id = Column(Integer, primary_key=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class ValidationReportDB(Base):
    __tablename__ = "validation_reports"
    id = Column(Integer, primary_key=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    doc_number = Column(String(100), nullable=True)
    user_id = Column(String(50), index=True, default="default_user")
    created_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class DraftDB(Base):
    __tablename__ = "drafts"
    id = Column(Integer, primary_key=True, autoincrement=True)
    kind = Column(String(50), nullable=False, index=True)
    draft_key = Column(String(100), nullable=False, index=True)
    user_id = Column(String(50), index=True, default="default_user")
    updated_at = Column(DateTime, default=lambda: dt.datetime.now(dt.timezone.utc), onupdate=lambda: dt.datetime.now(dt.timezone.utc))
    form_data = Column(JSON, nullable=False)


# ── 历史记录响应 ───────────────────────────────────────────────────────────────

class ReportHistoryItem(BaseModel):
    id: int
    report_title: str
    doc_number: Optional[str]
    created_at: dt.datetime
    docx_file_path: Optional[str]
    pdf_file_path: Optional[str]

    class Config:
        from_attributes = True


class ReportHistoryResponse(BaseModel):
    total: int
    reports: List[ReportHistoryItem]
