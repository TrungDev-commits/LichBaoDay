import os
from typing import List, Dict, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from app.core.schemas import TKBSlot, BaiHoc, ScheduleRow
from app.core.engine import generate_multi_week_schedule
from app.parsers.tkb_parser import auto_parse_tkb
from app.parsers.ppct_parser import auto_parse_ppct
from app.exporters.docx_export import export_to_docx, export_multi_week_to_docx, export_weeks_to_zip

app = FastAPI(
    title="Auto Lịch Báo Dạy Engine API",
    description="API tự động đọc TKB + PPCT/Giáo án và xuất file Word Lịch báo dạy chuẩn định dạng",
    version="1.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

frontend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "frontend"))

@app.get("/")
def read_root():
    index_path = os.path.join(frontend_dir, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {
        "status": "online",
        "service": "Auto Lịch Báo Dạy Engine",
        "docs_url": "/docs"
    }

class ExportDocxRequest(BaseModel):
    tuan: int = 1
    schedule: List[ScheduleRow]

class ExportMultiWeekRequest(BaseModel):
    multi_schedule: Dict[int, List[ScheduleRow]]

@app.post("/api/generate-preview")
async def generate_preview(
    start_tuan: int = Form(1),
    end_tuan: int = Form(1),
    file_tkb: UploadFile = File(..., description="File Thời khóa biểu (.xlsx hoặc .docx)"),
    file_ppct: UploadFile = File(..., description="File Phân phối chương trình (.docx, .xlsx, hoặc .pdf)")
):
    """
    Nhận 2 file TKB và PPCT, tự động ghép lịch cho 1 hoặc nhiều tuần (từ start_tuan đến end_tuan).
    """
    try:
        tkb_bytes = await file_tkb.read()
        tkb_slots = auto_parse_tkb(tkb_bytes, file_tkb.filename)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi đọc file TKB '{file_tkb.filename}': {str(e)}")

    if not tkb_slots:
        raise HTTPException(status_code=400, detail=f"Không bóc tách được tiết học nào từ file TKB '{file_tkb.filename}'. Vui lòng kiểm tra lại cấu trúc bảng TKB.")

    try:
        ppct_bytes = await file_ppct.read()
        ppct_items = auto_parse_ppct(ppct_bytes, file_ppct.filename)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi đọc file PPCT '{file_ppct.filename}': {str(e)}")

    if not ppct_items:
        raise HTTPException(status_code=400, detail=f"Không tìm thấy bài học nào trong file PPCT '{file_ppct.filename}'. Vui lòng kiểm tra lại cấu trúc bảng bài học.")

    try:
        multi_schedule = generate_multi_week_schedule(tkb_slots, ppct_items, start_tuan=start_tuan, end_tuan=end_tuan)
        return multi_schedule
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi trong quá trình ghép lịch: {str(e)}")

@app.post("/api/export-docx")
async def export_docx_endpoint(req: ExportDocxRequest):
    """
    Xuất 1 tuần thành file Word.
    """
    try:
        docx_stream = export_to_docx(req.schedule, tuan=req.tuan)
        return StreamingResponse(
            docx_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename=Lich_Bao_Day_Tuan_{req.tuan}.docx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xuất file Word: {str(e)}")

@app.post("/api/export-multi-docx")
async def export_multi_docx_endpoint(req: ExportMultiWeekRequest):
    """
    Xuất nhiều tuần vào 1 file Word duy nhất.
    """
    try:
        docx_stream = export_multi_week_to_docx(req.multi_schedule)
        weeks_str = f"Tuan_{min(req.multi_schedule.keys())}_den_{max(req.multi_schedule.keys())}"
        return StreamingResponse(
            docx_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename=Lich_Bao_Day_{weeks_str}.docx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xuất file Word: {str(e)}")

@app.post("/api/export-zip")
async def export_zip_endpoint(req: ExportMultiWeekRequest):
    """
    Xuất nhiều tuần thành nén ZIP chứa từng file Word riêng lẻ.
    """
    try:
        zip_stream = export_weeks_to_zip(req.multi_schedule)
        weeks_str = f"Tuan_{min(req.multi_schedule.keys())}_den_{max(req.multi_schedule.keys())}"
        return StreamingResponse(
            zip_stream,
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=Lich_Bao_Day_{weeks_str}.zip"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xuất file ZIP: {str(e)}")
