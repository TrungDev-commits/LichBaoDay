import io
import os
import sys
import copy
import datetime
from typing import List, Dict, Optional
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docxcompose.composer import Composer

from app.core.schemas import ScheduleRow, LessonBlock


def get_template_path() -> str:
    """Xác định đường dẫn file mẫu template_final.docx (tự tương thích PyInstaller frozen & local dev)."""
    possible_paths = []
    if getattr(sys, 'frozen', False):
        base_dir = getattr(sys, '_MEIPASS', '')
        if base_dir:
            possible_paths.extend([
                os.path.join(base_dir, "backend", "app", "templates", "template_final.docx"),
                os.path.join(base_dir, "app", "templates", "mau_lich.docx"),
                os.path.join(base_dir, "app", "templates", "template_final.docx"),
            ])
    
    curr_dir = os.path.dirname(os.path.abspath(__file__))
    possible_paths.extend([
        os.path.abspath(os.path.join(curr_dir, "..", "templates", "template_final.docx")),
        os.path.abspath(os.path.join(curr_dir, "..", "..", "app", "templates", "mau_lich.docx")),
        os.path.abspath(os.path.join(curr_dir, "..", "..", "..", "app", "templates", "mau_lich.docx")),
    ])
    
    for p in possible_paths:
        if p and os.path.exists(p):
            return p
            
    return os.path.abspath(os.path.join(curr_dir, "..", "templates", "template_final.docx"))


TEMPLATE_PATH = get_template_path()

# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────
_THU_TO_DAY_OFFSET = {
    "Hai": 0, "Ba": 1, "Tư": 2, "Năm": 3, "Sáu": 4, "Bảy": 5
}

_THU_FULL = {
    "Hai": "hai", "Ba": "ba", "Tư": "tư", "Năm": "năm", "Sáu": "sáu", "Bảy": "bảy"
}


def _parse_start_monday(start_monday: str) -> datetime.datetime:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(start_monday, fmt)
        except ValueError:
            continue
    return datetime.datetime.now()


def _date_for_thu(start_monday: str, thu: str, tuan: int = 1) -> datetime.datetime:
    """Tính ngày cụ thể của thứ X trong tuần `tuan` bắt đầu từ start_monday (Thứ 2 của Tuần 1)."""
    d = _parse_start_monday(start_monday)
    week_offset = (tuan - 1) * 7
    day_offset = _THU_TO_DAY_OFFSET.get(thu, 0)
    return d + datetime.timedelta(days=week_offset + day_offset)


def _add_page_break(doc: Document):
    """Thêm page break vào cuối document."""
    para = doc.add_paragraph()
    run = para.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


def _copy_element(elem):
    """Deep copy một xml element."""
    return copy.deepcopy(elem)


def _build_header_doc(
    tuan: int,
    thu: str,
    date: datetime.datetime,
    gv_name: str,
    mon: str,
    tiet_ppct: str,
) -> Document:
    """
    Tạo Document chứa 4 dòng header cho một tiết dạy:
      Tuần: 01
      Thứ hai, ngày 07 tháng 9 năm 2026
      Họ và tên người thực hiện: Lâm Huệ Trí
      Toán – Tiết 1
    """
    doc = Document()
    tuan_str = f"{tuan:02d}"
    thu_full = _THU_FULL.get(thu, thu.lower())
    date_str = f"Thứ {thu_full}, ngày {date.day:02d} tháng {date.month} năm {date.year}"
    tiet_str = f"Tiết {tiet_ppct}" if tiet_ppct else ""
    mon_tiet = f"{mon} – {tiet_str}" if tiet_str else mon

    lines = [
        f"Tuần: {tuan_str}",
        date_str,
        f"Họ và tên người thực hiện: {gv_name}",
        mon_tiet,
    ]
    for line in lines:
        p = doc.add_paragraph(line)
        if p.runs:
            p.runs[0].bold = True
    return doc


def _append_lesson_block(composer: Composer, header_doc: Document, block: Optional[LessonBlock]):
    """
    Append header + nội dung lesson block vào composer.
    
    Strategy:
    - Header doc: tạo Document mới sạch, chỉ chứa 4 dòng text → safe to append
    - Block elements: nếu block có docx source document, append toàn bộ source doc
      nhưng chỉ giữ lại phần nội dung của block (bằng cách crop về tập elements cần).
    
    Do limitation của docxcompose với relationships, chúng ta append header_doc
    rồi append block.source_doc trực tiếp (đã được clone và trim).
    """
    # Append header (Document sạch, không có relationships phức tạp)
    composer.append(header_doc)

    if block and block.elements:
        # Tạo document mới chứa chỉ các elements của block này
        # Bằng cách copy deep-clone toàn bộ parent document rồi giữ lại body elements
        try:
            _append_block_elements_safe(composer, block)
        except Exception as e:
            print(f"[docx_export] Warning: không thể append block '{block.ten_bai}': {e}")


def _append_block_elements_safe(composer: Composer, block: LessonBlock):
    """
    Tạo Document tạm chứa nội dung text của block theo cách an toàn.
    Dùng python-docx API để copy text/tables, tránh vấn đề relationship IDs.
    """
    from docx import Document as DocxDocument
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    tmp_doc = DocxDocument()
    # Xóa paragraph mặc định
    for p in list(tmp_doc.paragraphs):
        p._element.getparent().remove(p._element)

    has_content = False

    for elem in block.elements:
        try:
            if hasattr(elem, '_element') and elem.__class__.__name__ == 'Paragraph':
                # Paragraph: copy text với formatting cơ bản
                para = elem
                txt = para.text
                if not txt.strip():
                    tmp_doc.add_paragraph("")
                else:
                    new_p = tmp_doc.add_paragraph()
                    # Copy alignment
                    if para.alignment is not None:
                        new_p.alignment = para.alignment
                    # Copy paragraph format (indent, spacing)
                    if para.style and para.style.name:
                        try:
                            new_p.style = tmp_doc.styles[para.style.name]
                        except Exception:
                            pass
                    # Copy runs với bold/italic/underline
                    for run in para.runs:
                        new_run = new_p.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.underline = run.underline
                        if run.font.size:
                            new_run.font.size = run.font.size
                has_content = True

            elif hasattr(elem, '_tbl') or elem.__class__.__name__ == 'Table':
                # Table: copy text content cell by cell
                table = elem
                col_count = max(len(r.cells) for r in table.rows) if table.rows else 1
                new_table = tmp_doc.add_table(rows=len(table.rows), cols=col_count)
                for r_i, row in enumerate(table.rows):
                    for c_i, cell in enumerate(row.cells):
                        if c_i < col_count:
                            try:
                                new_table.cell(r_i, c_i).text = cell.text
                            except Exception:
                                pass
                has_content = True

        except Exception as ex:
            # Log nhưng không dừng
            txt_preview = getattr(elem, 'text', '')[:30] if hasattr(elem, 'text') else '?'
            print(f"  [skip elem] {type(elem).__name__}: {txt_preview!r} — {ex}")
            continue

    if has_content:
        composer.append(tmp_doc)




def _strip_external_refs(xml_elem):
    """
    Xóa các reference đến relationships phức tạp (hình ảnh, embedded)
    để tránh KeyError khi docxcompose copy relationship.
    Chỉ giữ lại text content.
    """
    from lxml import etree
    # Tags cần xóa (hình ảnh, drawing, OLE objects)
    TAGS_TO_REMOVE = {
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing',
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict',
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object',
    }
    for tag in TAGS_TO_REMOVE:
        for elem in xml_elem.iter(tag):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)




# ─────────────────────────────────────────
# FILL TEMPLATE
# ─────────────────────────────────────────
def fill_template_docx(
    schedule: List[ScheduleRow],
    tuan: int = 1,
    gv_name: str = "Lâm Huệ Trí",
    lop: str = "5/5",
    nam_hoc: str = "2026 - 2027",
    start_monday: str = "07/09/2026",
    template_path: str = TEMPLATE_PATH,
) -> Document:
    """
    Điền thông tin Bìa, Header và Bảng Lịch Báo Dạy vào file mẫu template_final.docx.
    """
    if not template_path or not os.path.exists(template_path):
        template_path = get_template_path()

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Không tìm thấy file mẫu Word template tại: {template_path}")

    doc = Document(template_path)
    d_start = _parse_start_monday(start_monday)
    d_tuan_start = d_start + datetime.timedelta(days=(tuan - 1) * 7)  # Thứ 2 của Tuần tuan
    d_tuan_end = d_tuan_start + datetime.timedelta(days=4)

    ngay_tu = d_tuan_start.strftime("%d/%m/%Y")
    ngay_den = d_tuan_end.strftime("%d/%m/%Y")
    tuan_str = f"{tuan:02d}"

    # 1. Cập nhật Bìa (paragraphs)
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt == "LỚP" or txt.startswith("LỚP "):
            _replace_para_text(p, f"LỚP {lop}")
        elif txt == "TUẦN" or txt.startswith("TUẦN "):
            _replace_para_text(p, f"TUẦN {tuan_str}")
        elif "Tuần học thứ" in txt:
            _replace_para_text(p, f"Tuần học thứ {tuan_str} ( Từ {ngay_tu} Đến {ngay_den})")

    # 2. Cập nhật Table chứa thông tin GV (duyệt tất cả các bảng để tìm đúng cell)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_txt = cell.text
                if "Họ, tên giáo viên" in cell_txt or "GVCN Lớp" in cell_txt or "Họ tên giáo viên" in cell_txt:
                    _fill_teacher_info_cell(cell, gv_name, lop, nam_hoc)

    # 3. Cập nhật Table 2 — Bảng Lịch Báo Dạy
    # Tìm bảng có cột "Thứ" ở đầu (thường là bảng thứ 2)
    lbd_table = _find_lbd_table(doc)
    if lbd_table and schedule:
        _fill_lbd_table(lbd_table, schedule, lop, tuan_str)

    return doc


def _fill_teacher_info_cell(cell, gv_name: str, lop: str, nam_hoc: str):
    """
    Điền đầy đủ thông tin giáo viên vào cell.
    Preserve 5 dòng/paragraphs gốc:
      - Họ, tên giáo viên: ...
      - GVCN Lớp: ...
      - Trường: Tiểu học Tiểu Cần 
      - Xã Tiểu Cần
      - Năm học: ...
    """
    lines = [
        f"- Họ, tên giáo viên: {gv_name}",
        f"- GVCN Lớp: {lop}",
        f"- Trường: Tiểu học Tiểu Cần ",
        f"- Xã Tiểu Cần",
        f"- Năm học: {nam_hoc}",
    ]
    if len(cell.paragraphs) >= 4:
        for idx, text in enumerate(lines):
            if idx < len(cell.paragraphs):
                _replace_para_text(cell.paragraphs[idx], text)
            else:
                p = cell.add_paragraph(text)
                if p.runs:
                    p.runs[0].bold = True
    else:
        _replace_cell_text(cell, "\n".join(lines))



def _replace_para_text(para, new_text: str):
    """Thay text của paragraph, giữ nguyên run đầu tiên để preserve formatting."""
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_text
    else:
        para.text = new_text


def _replace_cell_text(cell, new_text: str):
    """Thay text của cell an toàn."""
    text_str = str(new_text or "")
    if not cell.paragraphs:
        cell.add_paragraph(text_str)
        return
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
    if cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = text_str
    else:
        cell.paragraphs[0].text = text_str


def _find_lbd_table(doc: Document):
    """Tìm bảng Lịch Báo Dạy (có cột THỨ / BUỔI / TIẾT)."""
    for table in doc.tables:
        if len(table.rows) < 2:
            continue
        header_text = " ".join(c.text.upper() for c in table.rows[0].cells)
        if "THỨ" in header_text or "BUỔI" in header_text or "MÔN" in header_text:
            return table
    # Fallback: bảng cuối cùng có nhiều cột (>= 6)
    for table in reversed(doc.tables):
        if table.rows and len(table.rows[0].cells) >= 6:
            return table
    return None


def _fill_lbd_table(table, schedule: List[ScheduleRow], lop: str, tuan_str: str):
    """
    Điền dữ liệu vào Bảng Lịch Báo Dạy.
    Cấu trúc cột (theo template_final.docx):
      [0] Thứ | [1] Buổi | [2] Tiết | [3] Môn | [4] Lớp | [5] Tiết PPCT | [6] Tên bài | [7] Ghi chú
    
    Đọc từng hàng của bảng template, map (Thứ, Buổi, Tiết) với schedule.
    """
    # Build lookup: (thu_norm, buoi_upper, tiet_str) → ScheduleRow
    slot_map: Dict = {}
    for r in schedule:
        thu_str = (r.thu or "").strip()
        buoi_str = (r.buoi or "").strip().capitalize()
        key = (thu_str, buoi_str, str(r.tiet_tkb or ""))
        slot_map[key] = r

    for r_idx in range(1, len(table.rows)):
        row = table.rows[r_idx]
        cells = row.cells
        if len(cells) < 6:
            continue

        c_thu = cells[0].text.strip()
        c_buoi = cells[1].text.strip().capitalize()
        c_tiet = cells[2].text.strip()

        key = (c_thu, c_buoi, c_tiet)
        if key in slot_map:
            data = slot_map[key]
            _replace_cell_text(cells[3], data.mon or "")
            _replace_cell_text(cells[4], data.lop or lop)
            _replace_cell_text(cells[5], str(data.tiet_ppct or ""))
            _replace_cell_text(cells[6], data.ten_bai or "")
            if len(cells) > 7:
                _replace_cell_text(cells[7], data.ghi_chu or "")


# ─────────────────────────────────────────
# EXPORT COMBINED WEEK (V3)
# ─────────────────────────────────────────
def export_combined_week_docx(
    schedule: List[ScheduleRow],
    ordered_blocks: List[LessonBlock],
    tuan: int = 1,
    gv_name: str = "Lâm Huệ Trí",
    lop: str = "5/5",
    nam_hoc: str = "2026 - 2027",
    start_monday: str = "07/09/2026",
) -> io.BytesIO:
    """
    Xuất 1 tuần Lịch báo dạy dạng File Gộp:
      Bìa + Bảng LBD + [Header + Nội dung Giáo án chi tiết] theo thứ tự TKB.
    
    ordered_blocks phải được sắp xếp sẵn theo thứ tự TKB (do engine.generate_schedule trả về).
    Với môn chuyên trách không có giáo án: chỉ tạo Header block, không append nội dung.
    """
    doc_base = fill_template_docx(
        schedule=schedule,
        tuan=tuan,
        gv_name=gv_name,
        lop=lop,
        nam_hoc=nam_hoc,
        start_monday=start_monday,
    )

    composer = Composer(doc_base)

    # Duyệt song song schedule và ordered_blocks
    # ordered_blocks có thể ngắn hơn schedule (môn không có giáo án không có block)
    # Cần match đúng: block[i] tương ứng với schedule row có giáo án
    block_iter = iter(ordered_blocks)

    for row in schedule:
        # Tính ngày của tiết này theo Tuần tuan
        date = _date_for_thu(start_monday, row.thu, tuan=tuan)

        if row.tiet_ppct:  # có giáo án
            block = next(block_iter, None)
        else:
            block = None  # môn chuyên trách

        header_doc = _build_header_doc(
            tuan=tuan,
            thu=row.thu,
            date=date,
            gv_name=gv_name,
            mon=row.mon,
            tiet_ppct=row.tiet_ppct,
        )
        _append_lesson_block(composer, header_doc, block)

    stream = io.BytesIO()
    composer.save(stream)
    stream.seek(0)
    return stream


# ─────────────────────────────────────────
# EXPORT MULTI-WEEK
# ─────────────────────────────────────────
def export_multi_week_combined_docx(
    multi_data: Dict[int, tuple],  # tuan → (schedule, ordered_blocks)
    gv_name: str = "Lâm Huệ Trí",
    lop: str = "5/5",
    nam_hoc: str = "2026 - 2027",
    start_monday: str = "07/09/2026",
) -> io.BytesIO:
    """
    Xuất nhiều tuần liên tiếp vào 1 file Word gộp duy nhất.
    """
    sorted_weeks = sorted(multi_data.keys())
    if not sorted_weeks:
        raise ValueError("Không có dữ liệu tuần nào để xuất.")

    first_week = sorted_weeks[0]
    first_schedule, first_blocks = multi_data[first_week]

    doc_base = fill_template_docx(
        schedule=first_schedule,
        tuan=first_week,
        gv_name=gv_name,
        lop=lop,
        nam_hoc=nam_hoc,
        start_monday=start_monday,
    )
    composer = Composer(doc_base)
    _append_giaoan_blocks(composer, first_schedule, first_blocks, first_week, gv_name, start_monday)

    for w_num in sorted_weeks[1:]:
        w_schedule, w_blocks = multi_data[w_num]
        w_doc = fill_template_docx(
            schedule=w_schedule,
            tuan=w_num,
            gv_name=gv_name,
            lop=lop,
            nam_hoc=nam_hoc,
            start_monday=start_monday,
        )
        composer.append(w_doc)
        _append_giaoan_blocks(composer, w_schedule, w_blocks, w_num, gv_name, start_monday)

    stream = io.BytesIO()
    composer.save(stream)
    stream.seek(0)
    return stream


def _append_giaoan_blocks(
    composer: Composer,
    schedule: List[ScheduleRow],
    ordered_blocks: List[LessonBlock],
    tuan: int,
    gv_name: str,
    start_monday: str,
):
    """Helper: append tất cả header + block giáo án cho một tuần vào composer."""
    block_iter = iter(ordered_blocks)
    for row in schedule:
        date = _date_for_thu(start_monday, row.thu, tuan=tuan)
        if row.tiet_ppct:
            block = next(block_iter, None)
        else:
            block = None
        header_doc = _build_header_doc(
            tuan=tuan,
            thu=row.thu,
            date=date,
            gv_name=gv_name,
            mon=row.mon,
            tiet_ppct=row.tiet_ppct,
        )
        _append_lesson_block(composer, header_doc, block)
