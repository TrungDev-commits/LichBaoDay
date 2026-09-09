from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field

@dataclass
class LessonBlock:
    """
    Block nội dung chi tiết của 1 tiết dạy.
    Dùng dataclass thuần (không Pydantic) vì elements chứa docx objects.
    """
    mon: str
    tiet_ppct: int
    ten_bai: str
    elements: List[Any] = field(default_factory=list)  # docx Paragraph/Table objects

class TKBSlot(BaseModel):
    thu: str = Field(..., description="Thứ trong tuần (VD: Hai, Ba, Tư, Năm, Sáu)")
    buoi: str = Field("Sáng", description="Buổi học (Sáng / Chiều)")
    tiet_tkb: int = Field(..., description="Tiết theo TKB (1-5)")
    mon: str = Field(..., description="Tên môn học")
    lop: Optional[str] = Field("5/5", description="Tên lớp học")

class BaiHoc(BaseModel):
    mon: str = Field(..., description="Tên môn học")
    tiet_ppct: int = Field(..., description="Tiết thứ bao nhiêu theo PPCT")
    ten_bai: str = Field(..., description="Tên bài dạy / Tên bài học")
    thiet_bi: Optional[str] = Field("", description="Đồ dùng / Thiết bị dạy học")
    ghi_chu: Optional[str] = Field("", description="Ghi chú thêm / Nội dung điều chỉnh")

class ScheduleRow(BaseModel):
    thu: str
    buoi: str
    tiet_tkb: int
    mon: str
    lop: str = "5/5"
    tiet_ppct: str
    ten_bai: str
    thiet_bi: str = ""
    ghi_chu: str = ""

class SubjectFileItem(BaseModel):
    subject_name: str
    filename: str

class ParseTkbResponse(BaseModel):
    unique_subjects: List[str]
    slots: List[TKBSlot]

