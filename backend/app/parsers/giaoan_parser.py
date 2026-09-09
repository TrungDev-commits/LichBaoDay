import io
import re
from typing import List
from docx import Document

from app.core.schemas import LessonBlock


# ─────────────────────────────────────────
# PATTERN nhận diện heading tiết dạy
# ─────────────────────────────────────────
# Tiết heading (cấp 2): tạo block nội dung
_TIET_HEADING_RE = re.compile(
    r'^(?:'
    r'Tuần\s*\d+\s*[:\-\u2013]?\s*Tiết\s*[\d\+]+|'  # Tuần X: Tiết Y
    r'Tiết\s*[\d\+]+|'                            # Tiết X
    r'SHL\s*[:\-\u2013]'                               # SHL:
    r')',
    re.IGNORECASE
)

# Bài heading (cấp 1): chỉ match Bài có số thứ tự (Bài 1, Bài 01...)
_BAI_HEADING_RE = re.compile(
    r'^Bài\s*\d+\s*[:\-\u2013]',
    re.IGNORECASE
)

_CHUDE_HEADING_RE = re.compile(
    r'^\s*(?:TUẦN\s*\d+[\s\t:]*)?(?:CHỦ ĐỀ|PHẦN)\s*\d*',
    re.IGNORECASE
)


def clean_ten_bai(raw_title: str) -> str:
    """
    Rút trích tên bài dạy cốt lõi từ dòng heading.
    Chỉ lấy dòng đầu tiên của heading, loại bỏ tiền tố.
    """
    if not raw_title:
        return ""
    # Chỉ lấy dòng đầu (phòng trường hợp heading multi-line)
    t = raw_title.strip().split("\n")[0].strip()

    # Bỏ "Tuần X:" ở đầu
    t = re.sub(r'^\s*Tuần\s*\d+\s*[:\-–]\s*', '', t, flags=re.IGNORECASE).strip()
    # Bỏ "Tiết X(+Y):" ở đầu
    t = re.sub(r'^\s*Tiết\s*[\d\+]+\s*[:\-–]?\s*', '', t, flags=re.IGNORECASE).strip()
    # Bỏ "Bài X:" ở đầu
    t = re.sub(r'^\s*Bài\s*\d*\s*[:\-–]?\s*', '', t, flags=re.IGNORECASE).strip()
    # Bỏ "SHL:" ở đầu
    t = re.sub(r'^\s*SHL\s*[:\-–]?\s*', '', t, flags=re.IGNORECASE).strip()
    # Bỏ dấu gạch đầu
    t = re.sub(r'^\s*[-–]\s*', '', t).strip()
    # Bỏ dấu ":" đầu
    t = re.sub(r'^\s*:\s*', '', t).strip()

    # Nếu còn tiền tố dạng "Sinh hoạt dưới cờ: XYZ" → giữ nguyên phần sau ":"
    # nhưng bỏ phần "CHÀO NĂM HỌC MỚI" kiểu full-caps → viết hoa thường
    if t.isupper() and len(t) > 3:
        t = t.capitalize()

    return t


def _is_chude_heading(paragraph) -> bool:
    """Kiểm tra paragraph có phải tiêu đề Chủ đề / Phần học không."""
    txt = paragraph.text.strip()
    if not txt or len(txt) > 250:
        return False
    return bool(_CHUDE_HEADING_RE.match(txt))



def _is_tiet_heading(paragraph) -> bool:
    """Kiểm tra paragraph có phải heading tiết dạy (cấp 2 - tạo block) không."""
    txt = paragraph.text.strip()
    if not txt or len(txt) > 250:
        return False
    return bool(_TIET_HEADING_RE.match(txt))


def _is_bai_heading(paragraph) -> bool:
    """Kiểm tra paragraph có phải heading Bài (cấp 1 - không tạo block) không."""
    txt = paragraph.text.strip()
    if not txt or len(txt) > 250:
        return False
    return bool(_BAI_HEADING_RE.match(txt))


def parse_giaoan_docx(file_bytes: bytes, subject_name: str) -> List[LessonBlock]:
    """
    Parse file giáo án .docx của một môn học.
    Trả về List[LessonBlock], mỗi block gồm:
      - tiet_ppct: số thứ tự tiết (1-based, đếm tuần tự theo file)
      - ten_bai: tên bài ngắn gọn
      - elements: list các paragraph/table docx object của tiết đó

    Tự động phát hiện format:
      - Format A (TV): có heading "Tiết X:" cấp 2 → block theo Tiết
      - Format B (Toán, Khoa, SĐ...): chỉ có "Bài X:" → block theo Bài
    """
    doc = Document(io.BytesIO(file_bytes))

    # Tạo mapping xml_element → object để tra cứu nhanh
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}
    body = doc.element.body

    # Pass 1: Kiểm tra có Tiết heading không
    has_tiet_heading = any(
        _is_tiet_heading(para_map[child])
        for child in body.iterchildren()
        if child.tag.split('}')[-1] == 'p' and child in para_map
    )

    if has_tiet_heading:
        # Format A: TV-like — block theo Tiết, Bài là cấp cao hơn (không tạo block)
        return _parse_by_tiet(doc, para_map, table_map, body, subject_name)
    else:
        # Format B: Toán/Khoa-like — block theo Bài
        return _parse_by_bai(doc, para_map, table_map, body, subject_name)


def _format_combined_title(last_bai_para, tiet_para, next_para=None) -> str:
    """
    Tạo tên bài học kết hợp giữa Bài lớn + Tiết nhỏ (ví dụ: Thanh âm của gió (T1 đọc)).
    """
    raw_tiet = tiet_para.text.strip() if tiet_para else ""
    raw_bai = last_bai_para.text.strip() if last_bai_para else ""
    raw_next = next_para.text.strip() if next_para else ""

    clean_tiet = clean_ten_bai(raw_tiet)
    clean_bai = clean_ten_bai(raw_bai)
    clean_bai = re.sub(r'\(\s*\d+\s*tiết\s*\)', '', clean_bai, flags=re.IGNORECASE).strip()

    # Nếu dòng liền sau Tiết heading là "Bài: LUYỆN TẬP..." → dùng tên bài chi tiết đó
    if raw_next.startswith("Bài:") or raw_next.startswith("Bài :"):
        clean_next = clean_ten_bai(raw_next)
        if clean_next:
            return clean_next

    tiet_lower = clean_tiet.lower()
    if tiet_lower in ["đọc", "viết", "luyện từ và câu", "đọc mở rộng"]:
        m = re.search(r'Tiết\s*([\d\+]+)', raw_tiet, re.IGNORECASE)
        tiet_num = f"T{m.group(1)} " if m else ""
        if clean_bai:
            if tiet_lower == "đọc":
                return f"{clean_bai} ({tiet_num.strip()} đọc)"
            elif tiet_lower == "đọc mở rộng":
                return f"{clean_bai} - Đọc mở rộng"
            return f"{clean_bai} - {clean_tiet}"
        return clean_tiet

    if "chào năm học mới" in clean_tiet.lower():
        return "Chào năm học mới"
    if "chúng mình đã lớn" in clean_tiet.lower():
        return "Sinh hoạt chủ đề: CHÚNG MÌNH ĐÃ LỚN"
    if "bậc thang trưởng thành" in clean_tiet.lower():
        return "SHL: Bậc thang trưởng thành"

    return clean_tiet or clean_bai


def _parse_by_tiet(doc, para_map, table_map, body, subject_name: str) -> List[LessonBlock]:
    """Parse theo cấp Tiết (Bài là cấp cha, không tạo block)."""
    blocks = []
    current_ten_bai = ""
    current_elements = []
    ppct_counter = 0
    in_block = False

    last_chude_para = None
    last_bai_para = None

    # Lấy danh sách paragraph để tra cứu next paragraph
    children_p = [para_map[c] for c in body.iterchildren() if c.tag.endswith('p') and c in para_map]

    for child in body.iterchildren():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            para = para_map.get(child)
            if para is None:
                continue

            if _is_chude_heading(para):
                last_chude_para = para
            elif _is_bai_heading(para):
                last_bai_para = para
                # Kết thúc block tiết hiện tại (nếu có) khi sang Bài mới
                if in_block and current_elements:
                    blocks.append(LessonBlock(
                        mon=subject_name, tiet_ppct=ppct_counter,
                        ten_bai=current_ten_bai, elements=list(current_elements)
                    ))
                    current_elements = []
                    in_block = False

            elif _is_tiet_heading(para):
                # Kết thúc block tiết trước, bắt đầu block mới
                if in_block and current_elements:
                    blocks.append(LessonBlock(
                        mon=subject_name, tiet_ppct=ppct_counter,
                        ten_bai=current_ten_bai, elements=list(current_elements)
                    ))
                ppct_counter += 1

                # Tra cứu next paragraph để kiểm tra sub-title
                next_p = None
                if para in children_p:
                    curr_idx = children_p.index(para)
                    if curr_idx + 1 < len(children_p):
                        next_p = children_p[curr_idx + 1]

                current_ten_bai = _format_combined_title(last_bai_para, para, next_p)
                current_elements = []
                in_block = True

                # Bảo tồn paragraph Chủ đề và Bài heading vào đầu block tiết
                if last_chude_para and last_chude_para not in current_elements:
                    current_elements.append(last_chude_para)
                if last_bai_para and last_bai_para not in current_elements and last_bai_para != para:
                    current_elements.append(last_bai_para)
                current_elements.append(para)
            else:
                if in_block:
                    current_elements.append(para)

        elif tag == 'tbl':
            table = table_map.get(child)
            if table and in_block:
                current_elements.append(table)

    if in_block and current_elements:
        blocks.append(LessonBlock(
            mon=subject_name, tiet_ppct=ppct_counter,
            ten_bai=current_ten_bai, elements=list(current_elements)
        ))
    return blocks



def _parse_by_bai(doc, para_map, table_map, body, subject_name: str) -> List[LessonBlock]:
    """Parse theo cấp Bài (không có Tiết heading — mỗi Bài = 1 tiết PPCT)."""
    blocks = []
    current_ten_bai = ""
    current_elements = []
    ppct_counter = 0
    in_block = False

    last_chude_para = None

    for child in body.iterchildren():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            para = para_map.get(child)
            if para is None:
                continue

            if _is_chude_heading(para):
                last_chude_para = para

            elif _is_bai_heading(para):
                # Bài heading = block heading chính
                if in_block and current_elements:
                    blocks.append(LessonBlock(
                        mon=subject_name, tiet_ppct=ppct_counter,
                        ten_bai=current_ten_bai, elements=list(current_elements)
                    ))
                ppct_counter += 1
                current_ten_bai = clean_ten_bai(para.text.strip())
                current_elements = []
                in_block = True

                # Bảo tồn paragraph Chủ đề vào đầu block bài
                if last_chude_para and last_chude_para not in current_elements:
                    current_elements.append(last_chude_para)
                current_elements.append(para)
            else:
                if in_block:
                    current_elements.append(para)

        elif tag == 'tbl':
            table = table_map.get(child)
            if table and in_block:
                current_elements.append(table)

    if in_block and current_elements:
        blocks.append(LessonBlock(
            mon=subject_name, tiet_ppct=ppct_counter,
            ten_bai=current_ten_bai, elements=list(current_elements)
        ))
    return blocks


