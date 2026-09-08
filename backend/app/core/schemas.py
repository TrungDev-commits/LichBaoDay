from pydantic import BaseModel, Field
from typing import Optional, List

class TKBSlot(BaseModel):
    thu: str = Field(..., description="Thứ trong tuần (VD: Hai, Ba, Tư, Năm, Sáu)")
    buoi: str = Field("Sáng", description="Buổi học (Sáng / Chiều)")
    tiet_tkb: int = Field(..., description="Tiết theo TKB (1-5)")
    mon: str = Field(..., description="Tên môn học")
    lop: Optional[str] = Field("3A", description="Tên lớp học")

class BaiHoc(BaseModel):
    mon: str = Field(..., description="Tên môn học")
    tiet_ppct: int = Field(..., description="Tiết thứ bao nhiêu theo PPCT")
    ten_bai: str = Field(..., description="Tên bài dạy / Tên bài học")
    thiet_bi: Optional[str] = Field("", description="Đồ dùng / Thiết bị dạy học")
    ghi_chu: Optional[str] = Field("", description="Ghi chú thêm (VD: Tích hợp Stem...)")

class ScheduleRow(BaseModel):
    thu: str
    buoi: str
    tiet_tkb: int
    mon: str
    tiet_ppct: str
    ten_bai: str
    thiet_bi: str
    ghi_chu: str
