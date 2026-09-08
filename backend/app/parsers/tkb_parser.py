import io
from typing import List
import openpyxl
from docx import Document
from app.core.schemas import TKBSlot

def parse_tkb_excel(file_bytes: bytes) -> List[TKBSlot]:
    """
    Parse thời khóa biểu từ file Excel (.xlsx).
    Hỗ trợ cả dạng bảng dọc (Thứ, Buổi, Tiết, Môn) và dạng ma trận (Cột Thứ 2..6, Dòng Tiết 1..5).
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    
    slots: List[TKBSlot] = []
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return slots

    header_row_idx = -1
    day_cols = {}

    days_map = {
        "hai": "Hai", "2": "Hai", "thứ hai": "Hai", "thứ 2": "Hai", "t2": "Hai",
        "ba": "Ba", "3": "Ba", "thứ ba": "Ba", "thứ 3": "Ba", "t3": "Ba",
        "tư": "Tư", "4": "Tư", "thứ tư": "Tư", "thứ 4": "Tư", "t4": "Tư",
        "năm": "Năm", "5": "Năm", "thứ năm": "Năm", "thứ 5": "Năm", "t5": "Năm",
        "sáu": "Sáu", "6": "Sáu", "thứ sáu": "Sáu", "thứ 6": "Sáu", "t6": "Sáu",
        "bảy": "Bảy", "7": "Bảy", "thứ bảy": "Bảy", "thứ 7": "Bảy", "t7": "Bảy"
    }

    for r_idx, row in enumerate(rows[:10]):
        for c_idx, val in enumerate(row):
            if val is not None:
                val_str = str(val).strip().lower()
                for k, v in days_map.items():
                    if k == val_str or f"thứ {k}" in val_str or f"t{k}" in val_str:
                        day_cols[c_idx] = v
                        header_row_idx = r_idx

    if day_cols and header_row_idx != -1:
        current_buoi = "Sáng"
        tiet_counter = 1
        
        for r_idx in range(header_row_idx + 1, len(rows)):
            row = rows[r_idx]
            first_cell = str(row[0]).strip().lower() if row[0] is not None else ""
            
            if "chiều" in first_cell:
                current_buoi = "Chiều"
                tiet_counter = 1
                continue
            elif "sáng" in first_cell:
                current_buoi = "Sáng"
                tiet_counter = 1
                continue
                
            has_data = False
            for c_idx, thu_name in day_cols.items():
                if c_idx < len(row) and row[c_idx] is not None:
                    mon_name = str(row[c_idx]).strip()
                    if mon_name and mon_name != "-":
                        has_data = True
                        slots.append(TKBSlot(
                            thu=thu_name,
                            buoi=current_buoi,
                            tiet_tkb=tiet_counter,
                            mon=mon_name
                        ))
            if has_data:
                tiet_counter += 1
        return slots

    return slots

def parse_tkb_docx(file_bytes: bytes) -> List[TKBSlot]:
    """
    Parse thời khóa biểu từ bảng trong file Word (.docx).
    """
    doc = Document(io.BytesIO(file_bytes))
    slots: List[TKBSlot] = []

    for table in doc.tables:
        current_buoi = "Sáng"
        tiet_counter = 1
        day_cols = {}

        for r_idx, row in enumerate(table.rows):
            cells_text = [c.text.strip() for c in row.cells]
            
            if not day_cols:
                for c_idx, text in enumerate(cells_text):
                    t_lower = text.lower()
                    if "hai" in t_lower or "2" in t_lower:
                        day_cols[c_idx] = "Hai"
                    elif "ba" in t_lower or "3" in t_lower:
                        day_cols[c_idx] = "Ba"
                    elif "tư" in t_lower or "4" in t_lower:
                        day_cols[c_idx] = "Tư"
                    elif "năm" in t_lower or "5" in t_lower:
                        day_cols[c_idx] = "Năm"
                    elif "sáu" in t_lower or "6" in t_lower:
                        day_cols[c_idx] = "Sáu"
                continue

            row_joined = " ".join(cells_text).lower()
            if "chiều" in row_joined:
                current_buoi = "Chiều"
                tiet_counter = 1
                continue
            elif "sáng" in row_joined:
                current_buoi = "Sáng"
                tiet_counter = 1
                continue

            has_mon = False
            for c_idx, thu_name in day_cols.items():
                if c_idx < len(cells_text):
                    mon_name = cells_text[c_idx]
                    if mon_name and not mon_name.isdigit() and len(mon_name) > 1:
                        has_mon = True
                        slots.append(TKBSlot(
                            thu=thu_name,
                            buoi=current_buoi,
                            tiet_tkb=tiet_counter,
                            mon=mon_name
                        ))
            if has_mon:
                tiet_counter += 1

    return slots

def auto_parse_tkb(file_bytes: bytes, filename: str) -> List[TKBSlot]:
    """
    Tự động thử các parser TKB để đọc file TKB bất kể đuôi file bị đặt sai.
    """
    fn = filename.lower()
    parsers = []

    if fn.endswith(".xlsx") or fn.endswith(".xls"):
        parsers = [parse_tkb_excel, parse_tkb_docx]
    else:
        parsers = [parse_tkb_docx, parse_tkb_excel]

    for parser in parsers:
        try:
            res = parser(file_bytes)
            if res:
                return res
        except Exception:
            continue

    return []
