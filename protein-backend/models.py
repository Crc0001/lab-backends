from datetime import datetime
from typing import List, Optional, Any, Dict
from pydantic import BaseModel, Field
from sqlalchemy import Column, Integer, String, Float, DateTime, JSON
from database import Base


# ==================== SQLAlchemy 数据库模型 ====================

class FermentationReportDB(Base):
    __tablename__ = "fermentation_reports"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    report_title = Column(String(255), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    user_id = Column(String(50), index=True)

    materials_data = Column(JSON, nullable=False)
    fermentation_data = Column(JSON, nullable=False)
    expression_data = Column(JSON, nullable=False)
    purification_data = Column(JSON, nullable=False)

    sds_page_file_id = Column(String(100), nullable=True)
    mass_spec_file_id = Column(String(100), nullable=True)

    # V2 扩展字段：完整表单数据和文件路径
    form_data = Column(JSON, nullable=True)
    docx_file_path = Column(String(500), nullable=True)
    pdf_file_path = Column(String(500), nullable=True)


class DraftDB(Base):
    __tablename__ = "drafts"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    kind = Column(String(50), nullable=False, index=True)
    draft_key = Column(String(100), nullable=False, index=True)
    user_id = Column(String(50), index=True, default="default_user")
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    form_data = Column(JSON, nullable=False)


# ==================== Pydantic 请求与响应模型 ====================

class MaterialPrep(BaseModel):
    reagents: List[Dict[str, Any]] = Field(..., description="试剂信息：名称、厂家、批号")
    instruments: List[Dict[str, Any]] = Field(..., description="仪器信息：名称、厂家、编号")
    vector_name: str = Field(..., example="pET-28a")
    host_cell: str = Field(..., example="BL21(DE3)")
    target_gene_size_kda: float = Field(..., description="目的基因大小 (kDa)")
    tags: List[str] = Field(..., description="序列标签，如 His-tag, GST-tag")


class FermentationInduction(BaseModel):
    media_type: str = Field("LB", description="培养基类型：LB 或 TB")
    inoculation_volume_percent: float = Field(..., description="接种量(%)")
    culture_temperature: float = Field(..., description="培养温度(℃)")
    shaker_speed_rpm: int = Field(..., description="摇床转速(rpm)")
    induction_od600: float = Field(..., description="诱导时OD600nm值")
    iptg_concentration_mm: float = Field(..., description="IPTG终浓度(mM)")
    induction_temperature: float = Field(..., description="诱导温度(℃)")


class ExpressionForm(BaseModel):
    is_soluble: bool = Field(..., description="是否为可溶性表达")
    inclusion_body_treatment: Optional[str] = Field(None, description="包涵体表达处理工艺（若有）")


class PurificationResults(BaseModel):
    purification_method: str = Field("Ni-NTA", description="纯化方法")
    assay_method: str = Field("BCA", description="测定方法")
    protein_concentration_mg_ml: float = Field(..., description="蛋白浓度(mg/mL)")
    protein_purity_percent: float = Field(..., description="蛋白纯度(%)")


class ReportGenerationPayload(BaseModel):
    report_title: str
    user_id: str = Field("default_user", description="创建人ID")
    materials: MaterialPrep
    fermentation: FermentationInduction
    expression: ExpressionForm
    purification: PurificationResults
    sds_page_file_id: Optional[str] = None
    mass_spec_file_id: Optional[str] = None


class EchoConfigResponse(BaseModel):
    materials: MaterialPrep
    fermentation: FermentationInduction
    expression: ExpressionForm
    purification: PurificationResults


class ReportGenerationResponse(BaseModel):
    status: str
    report_id: int
    download_url: str


class UploadMediaResponse(BaseModel):
    file_id: str
    url: str


# ==================== V2 新架构模型 ====================

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


class ReportPayloadV2(BaseModel):
    report_title: str
    user_id: str = "default_user"
    doc_number: str = ""
    security_level: str = ""
    project_name: str = ""
    first_page_footer: str = ""
    personnel: List[PersonnelRow] = []
    approver_signatures: Optional[List[Optional[str]]] = None
    approver_signature_audits: List[SignatureAuditEntry] = []
    strains_headers: Optional[List[str]] = None
    strains: List[List[str]] = []
    reagents_headers: Optional[List[str]] = None
    reagents: List[List[str]] = []
    instruments_headers: Optional[List[str]] = None
    instruments: List[List[str]] = []
    reagent_config_headers: Optional[List[str]] = None
    reagent_configs: List[List[str]] = []
    consumables_headers: Optional[List[str]] = None
    consumables: List[List[str]] = []
    procedure_section_sub_titles: Optional[List[List[str]]] = None
    procedure_headers: Optional[List[List[List[str]]]] = None
    procedures: List[List[List[List[str]]]] = []
    dynamic_section_titles: List[str] = []
    dynamic_section_sub_titles: List[List[str]] = []
    dynamic_section_headers: List[List[List[str]]] = []
    dynamic_section_tables: List[List[List[List[str]]]] = []
    result_extra_headers: List[str] = []
    results: List[ResultRowData] = []
    qc_data_headers: Optional[List[str]] = None
    qc_concentrations: Optional[List[str]] = None
    qc_od_values: Optional[List[str]] = None
    qc_linear_equation: Optional[str] = None
    qc_correlation_coeff: Optional[float] = None
    qc_conclusion_headers: Optional[List[str]] = None
    qc_conclusion_row_labels: Optional[List[str]] = None
    qc_conclusion_rows: Optional[List[List[str]]] = None
    conclusion_table_headers: Optional[List[str]] = None
    conclusion_table_rows: List[List[str]] = []
    conclusion_items: List[str] = []
    appendix_file_paths: List[str] = []
