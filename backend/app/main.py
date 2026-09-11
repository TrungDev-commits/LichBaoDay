import os
import sys
import io

# Cấu hình UTF-8 cho stdout/stderr trên Windows để tránh lỗi UnicodeEncodeError khi print()
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Ensure backend directory is in sys.path
backend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

from typing import List, Dict, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware

from app.core.schemas import TKBSlot, ScheduleRow, ParseTkbResponse, LessonBlock
from app.core.engine import generate_multi_week_schedule, normalize_subject_name
from app.parsers.tkb_parser import auto_parse_tkb, extract_unique_subjects
from app.parsers.giaoan_parser import parse_giaoan_docx
from app.exporters.docx_export import (
    export_combined_week_docx,
    export_multi_week_combined_docx,
    fill_template_docx,
)

app = FastAPI(
    title="Auto Lịch Báo Dạy Engine API",
    description="API tự động đọc TKB + Phân tích Giáo án từng môn và xuất file Word Gộp (Bìa + Lịch + Giáo án chi tiết)",
    version="3.0.0"
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
    return {"status": "online", "service": "Auto Lịch Báo Dạy Engine v3.0", "docs_url": "/docs"}

@app.get("/logo.png")
def get_logo():
    logo_path = os.path.join(frontend_dir, "logo.png")
    if os.path.exists(logo_path):
        return FileResponse(logo_path, media_type="image/png")
    orig_logo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "DOCS", "Logo.png"))
    return FileResponse(orig_logo, media_type="image/png")

@app.get("/favicon.ico")
def get_favicon():
    fav_path = os.path.join(frontend_dir, "favicon.ico")
    if os.path.exists(fav_path):
        return FileResponse(fav_path, media_type="image/x-icon")
    return get_logo()



from app.parsers.tkb_parser import auto_parse_tkb, extract_unique_subjects, extract_unique_classes

@app.post("/api/parse-tkb", response_model=ParseTkbResponse)
async def parse_tkb_endpoint(file_tkb: UploadFile = File(...)):
    """
    Đọc file TKB → trả danh sách tiết + danh sách môn học độc bản (unique_subjects) + danh sách lớp (unique_classes).
    """
    try:
        tkb_bytes = await file_tkb.read()
        slots = auto_parse_tkb(tkb_bytes, file_tkb.filename)
        if not slots:
            raise HTTPException(status_code=400, detail="Không tìm thấy tiết học nào từ file TKB.")
        unique_subjects = extract_unique_subjects(slots)
        unique_classes = extract_unique_classes(slots)
        return ParseTkbResponse(unique_subjects=unique_subjects, unique_classes=unique_classes, slots=slots)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi khi đọc file TKB: {str(e)}")


@app.post("/api/generate-preview")
async def generate_preview(request: Request):
    """
    Preview Bảng Lịch Báo Dạy (JSON) để hiển thị trước khi xuất file.
    """
    try:
        form = await request.form()

        file_tkb = form.get("file_tkb")
        if not file_tkb or not hasattr(file_tkb, "read"):
            raise HTTPException(status_code=400, detail="Thiếu file_tkb!")

        tuan_raw = form.get("tuan", form.get("start_tuan", 1))
        tuan = int(tuan_raw)

        if hasattr(file_tkb, "seek"):
            await file_tkb.seek(0)
        tkb_bytes = await file_tkb.read()
        slots = _get_tkb_slots(form, tkb_bytes, getattr(file_tkb, "filename", "TKB.docx"))
        if not slots:
            raise HTTPException(status_code=400, detail="Không trích xuất được dữ liệu từ TKB.")

        lesson_blocks: List[LessonBlock] = await _parse_all_giaoan(form)

        multi_data = generate_multi_week_schedule(slots, lesson_blocks, start_tuan=tuan, end_tuan=tuan)
        schedule, _ = multi_data.get(tuan, ([], []))

        return {tuan: [row.dict() for row in schedule]}
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Lỗi khi xem trước: {str(e)}")


@app.post("/api/export-docx")
async def export_docx_endpoint(request: Request):
    """
    Xuất File Word Gộp cho 1 tuần:
    Bìa + Bảng LBD + Header & Nội dung Giáo án chi tiết theo thứ tự TKB.
    """
    try:
        form = await request.form()

        tuan_raw = form.get("tuan", form.get("start_tuan", 1))
        tuan = int(tuan_raw)
        gv_name = form.get("gv_name", "Lâm Huệ Trí")
        lop = form.get("lop", "5/5")
        nam_hoc = form.get("nam_hoc", "2026 - 2027")
        start_monday = form.get("start_monday", "07/09/2026")

        file_tkb = form.get("file_tkb")
        if not file_tkb or not hasattr(file_tkb, "read"):
            raise HTTPException(status_code=400, detail="Thiếu file_tkb!")

        if hasattr(file_tkb, "seek"):
            await file_tkb.seek(0)
        tkb_bytes = await file_tkb.read()
        slots = _get_tkb_slots(form, tkb_bytes, getattr(file_tkb, "filename", "TKB.docx"))
        if not slots:
            raise HTTPException(status_code=400, detail="Không trích xuất được dữ liệu từ TKB.")

        lesson_blocks: List[LessonBlock] = await _parse_all_giaoan(form)

        multi_data = generate_multi_week_schedule(slots, lesson_blocks, start_tuan=tuan, end_tuan=tuan)
        schedule, ordered_blocks = multi_data.get(tuan, ([], []))

        # Áp dụng các chỉnh sửa tên bài dạy / ghi chú từ frontend nếu có
        edited_multi_raw = form.get("edited_multi_schedule")
        edited_single_raw = form.get("edited_schedule")
        _apply_user_edits(schedule, tuan, edited_multi_raw, edited_single_raw)

        docx_stream = export_combined_week_docx(
            schedule=schedule,
            ordered_blocks=ordered_blocks,
            tuan=tuan,
            gv_name=gv_name,
            lop=lop,
            nam_hoc=nam_hoc,
            start_monday=start_monday,
        )

        filename = f"Lich_Bao_Day_Gop_Tuan_{tuan}.docx"
        return StreamingResponse(
            docx_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Lỗi xuất file Word: {str(e)}")


# Giữ backward-compat endpoint cũ trỏ vào endpoint mới
@app.post("/api/export-multi-docx")
async def export_multi_docx_endpoint(request: Request):
    """
    Backward-compatible: xuất file giống /api/export-docx.
    Đọc tuan từ start_tuan (cũ) hoặc tuan (mới).
    """
    return await export_docx_endpoint(request)


# ─────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────
def _get_tkb_slots(form, tkb_bytes: bytes, filename: str) -> List[TKBSlot]:
    import json
    edited_raw = form.get("edited_tkb_slots")
    if edited_raw:
        try:
            data = json.loads(edited_raw)
            if isinstance(data, list) and len(data) > 0:
                res = []
                for s in data:
                    if isinstance(s, dict) and s.get("mon"):
                        res.append(TKBSlot(
                            thu=s.get("thu", "Hai"),
                            buoi=s.get("buoi", "Sáng"),
                            tiet_tkb=int(s.get("tiet_tkb", 1)),
                            mon=str(s.get("mon", "")),
                            lop=str(s.get("lop", "5/5"))
                        ))
                if res:
                    return res
        except Exception as ex:
            print(f"[get_tkb_slots] Error parsing edited_tkb_slots: {ex}")
    return auto_parse_tkb(tkb_bytes, filename)
async def _parse_all_giaoan(form) -> List[LessonBlock]:
    """
    Đọc tất cả các file giáo án từ form (key bắt đầu bằng 'file_giaoan_').
    Trả về List[LessonBlock] tổng hợp tất cả môn.
    """
    all_blocks: List[LessonBlock] = []
    for key, val in form.items():
        if key.startswith("file_giaoan_") and hasattr(val, "read"):
            subj_name = key.replace("file_giaoan_", "").strip()
            try:
                if hasattr(val, "seek"):
                    await val.seek(0)
                g_bytes = await val.read()
                if not g_bytes:
                    continue
                blocks = parse_giaoan_docx(g_bytes, subj_name)
                all_blocks.extend(blocks)
                print(f"[parse_giaoan] Mon '{subj_name}' -> {len(blocks)} block(s) tim thay")
            except Exception as ex:
                print(f"[parse_giaoan] Loi parse '{subj_name}': {ex}")
    return all_blocks


def _apply_user_edits(
    schedule: List[ScheduleRow],
    tuan: int,
    edited_multi_raw: Optional[str] = None,
    edited_single_raw: Optional[str] = None,
):
    """
    Áp dụng các chỉnh sửa của người dùng từ frontend (ten_bai, ghi_chu) vào schedule trước khi xuất file Word.
    """
    import json
    edited_list = None
    if edited_multi_raw:
        try:
            data = json.loads(edited_multi_raw)
            if isinstance(data, dict):
                edited_list = data.get(str(tuan)) or data.get(tuan)
            elif isinstance(data, list):
                edited_list = data
        except Exception as ex:
            print(f"[apply_user_edits] Error parsing edited_multi_schedule: {ex}")

    if not edited_list and edited_single_raw:
        try:
            edited_list = json.loads(edited_single_raw)
        except Exception as ex:
            print(f"[apply_user_edits] Error parsing edited_schedule: {ex}")

    if not edited_list or not isinstance(edited_list, list):
        return

    # Map theo key (thu, buoi, tiet_tkb, mon)
    edited_map = {}
    for item in edited_list:
        if isinstance(item, dict):
            k = (
                str(item.get("thu", "")).strip(),
                str(item.get("buoi", "")).strip(),
                str(item.get("tiet_tkb", "")).strip(),
                str(item.get("mon", "")).strip(),
            )
            edited_map[k] = item

    for idx, row in enumerate(schedule):
        k = (
            str(row.thu or "").strip(),
            str(row.buoi or "").strip(),
            str(row.tiet_tkb or "").strip(),
            str(row.mon or "").strip(),
        )
        item = edited_map.get(k)
        if not item and idx < len(edited_list) and isinstance(edited_list[idx], dict):
            item = edited_list[idx]

        if item:
            if "ten_bai" in item and item["ten_bai"] is not None:
                row.ten_bai = str(item["ten_bai"])
            if "ghi_chu" in item and item["ghi_chu"] is not None:
                row.ghi_chu = str(item["ghi_chu"])


