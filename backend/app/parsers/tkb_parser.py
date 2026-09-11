import io
import re
from typing import List, Tuple
import openpyxl
from docx import Document
from app.core.schemas import TKBSlot


def _extract_mon_and_lop(raw_text: str) -> Tuple[str, str]:
    """
    Tách tên môn và tên lớp từ nội dung ô TKB.
    Ví dụ:
      'Toán (Lớp 3/1)' -> mon='Toán', lop='3/1'
      '3/1 - Toán'     -> mon='Toán', lop='3/1'
      'Toán 3/1'       -> mon='Toán', lop='3/1'
      'Toán'           -> mon='Toán', lop=''
    """
    if not raw_text:
        return "", ""
    txt = raw_text.strip()

    # Match pattern ngoặc đơn: "Toán (Lớp 3/1)" hoặc "Toán (3/1)"
    m2 = re.search(r'^(.*?)\s*\((?:lớp\s*)?([^)]+)\)$', txt, re.IGNORECASE)
    if m2:
        return m2.group(1).strip(), m2.group(2).strip()

    # Match pattern lớp dạng 3/1, 3/2, 4/5, 5A, 3B
    m = re.search(r'\b([1-5]/[1-9]|[1-5][A-H])\b', txt, re.IGNORECASE)
    if m:
        lop = m.group(1)
        mon = txt.replace(m.group(0), "").replace("Lớp", "").replace("lớp", "").strip(" -():,")
        return mon or txt, lop

    return txt, ""


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
                    raw_val = str(row[c_idx]).strip()
                    if raw_val and raw_val != "-":
                        has_data = True
                        mon_name, lop_name = _extract_mon_and_lop(raw_val)
                        slots.append(TKBSlot(
                            thu=thu_name,
                            buoi=current_buoi,
                            tiet_tkb=tiet_counter,
                            mon=mon_name,
                            lop=lop_name or "5/5"
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

            first_cell = cells_text[0].lower() if len(cells_text) > 0 else ""
            second_cell = cells_text[1].strip() if len(cells_text) > 1 else ""

            if "chiều" in first_cell and not second_cell.isdigit():
                current_buoi = "Chiều"
                tiet_counter = 1
                continue
            elif "sáng" in first_cell and not second_cell.isdigit():
                current_buoi = "Sáng"
                tiet_counter = 1
                continue

            if "chiều" in first_cell:
                current_buoi = "Chiều"
            elif "sáng" in first_cell:
                current_buoi = "Sáng"

            tiet_num = tiet_counter
            if second_cell.isdigit():
                tiet_num = int(second_cell)

            has_mon = False
            for c_idx, thu_name in day_cols.items():
                if c_idx < len(cells_text):
                    raw_val = cells_text[c_idx]
                    if raw_val and not raw_val.isdigit() and len(raw_val) > 1 and "NGHỈ GIẢI LAO" not in raw_val.upper():
                        has_mon = True
                        mon_name, lop_name = _extract_mon_and_lop(raw_val)
                        slots.append(TKBSlot(
                            thu=thu_name,
                            buoi=current_buoi,
                            tiet_tkb=tiet_num,
                            mon=mon_name,
                            lop=lop_name or "5/5"
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


def extract_unique_subjects(slots: List[TKBSlot]) -> List[str]:
    """
    Trích xuất danh sách môn học độc bản từ danh sách slot TKB.
    """
    seen = set()
    unique = []
    for s in slots:
        mon_clean = s.mon.strip()
        if mon_clean and mon_clean not in seen and not mon_clean.isdigit():
            seen.add(mon_clean)
            unique.append(mon_clean)
    return unique


def extract_unique_classes(slots: List[TKBSlot]) -> List[str]:
    """
    Trích xuất danh sách các lớp độc bản từ danh sách slot TKB.
    """
    seen = set()
    unique = []
    for s in slots:
        if s.lop:
            l_clean = s.lop.strip()
            if l_clean and l_clean not in seen and l_clean != "5/5":
                seen.add(l_clean)
                unique.append(l_clean)
    return unique

