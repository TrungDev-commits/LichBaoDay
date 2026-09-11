from collections import defaultdict
from typing import List, Dict, Optional, Tuple
import re

from app.core.schemas import TKBSlot, BaiHoc, ScheduleRow, LessonBlock


# ─────────────────────────────────────────────────────────
# SUBJECT NORMALIZER — map tên TKB → key chuẩn hóa
# ─────────────────────────────────────────────────────────
# Tất cả alias về cùng 1 key lowercase không dấu (để tra dict)
_SUBJECT_ALIAS: Dict[str, str] = {
    # HĐTN (Hoạt động trải nghiệm) — nhiều cách viết trên TKB
    "chào cờ - hđtn":   "hđtn",
    "chào cờ-hđtn":     "hđtn",
    "chào cờ hđtn":     "hđtn",
    "shtt (hđtn)":      "hđtn",
    "shtt(hđtn)":       "hđtn",
    "shtt - hđtn":      "hđtn",
    "shtt-hđtn":        "hđtn",
    "hđtn (thắng)":     "hđtn",
    "hđtn(thắng)":      "hđtn",
    "hđtn":             "hđtn",
    "hoạt động trải nghiệm": "hđtn",
    # Lịch sử - Địa lý
    "lịch sử-địa lí":   "sử địa",
    "lịch sử - địa lí": "sử địa",
    "l.sử-đ.lí":        "sử địa",
    "ls-đl":            "sử địa",
    "ls&đl":            "sử địa",
    "sử địa":           "sử địa",
    # Toán
    "toán":             "toán",
    "tăng cường toán":  "tăng cường toán",
    # Tiếng Việt
    "tiếng việt":       "tiếng việt",
    "t.việt":           "tiếng việt",
    "tv":               "tiếng việt",
    "tăng cường tv":    "tăng cường tiếng việt",
    "tăng cường tiếng việt": "tăng cường tiếng việt",
    # Khoa học
    "khoa học":         "khoa học",
    "khoa":             "khoa học",
    # Công nghệ
    "công nghệ":        "công nghệ",
    "cn":               "công nghệ",
    # Tin học
    "tin học":          "tin học",
    "tin":              "tin học",
    # Giáo dục thể chất
    "gdtc":             "thể dục",
    "thể dục":          "thể dục",
    "giáo dục thể chất": "thể dục",
    # Đạo đức
    "đạo đức":          "đạo đức",
    "đđ":               "đạo đức",
    # Âm nhạc
    "âm nhạc":          "âm nhạc",
    "ân":               "âm nhạc",
    # Mĩ thuật
    "mĩ thuật":         "mĩ thuật",
    "mth":              "mĩ thuật",
    # Tiếng Anh
    "tiếng anh":        "tiếng anh",
    "t.anh":            "tiếng anh",
    # Tự nhiên và xã hội
    "tnxh":             "tnxh",
    "tự nhiên và xã hội": "tnxh",
    # ─── Alias tên file không dấu (tên file trên disk) ───
    "hdtn":             "hđtn",       # HDTN.docx → hđtn
    "sd":               "sử địa",     # SD.docx   → sử địa
    "sđ":               "sử địa",     # SĐ.docx   → sử địa (có dấu)
    "toan":             "toán",       # TOAN.docx → toán
    "khoa":             "khoa học",   # KHOA.docx → khoa học (đã có nhưng thêm rõ)
    "cong nghe":        "công nghệ",  # CONG NGHE.docx
    "tin hoc":          "tin học",
    "the duc":          "thể dục",
    "dao duc":          "đạo đức",
    "am nhac":          "âm nhạc",
    "mi thuat":         "mĩ thuật",
    "tieng anh":        "tiếng anh",
}


def normalize_subject_name(name: str) -> str:
    """
    Chuẩn hóa tên môn học để ghép chính xác giữa TKB và giáo án.
    Tra bảng alias trước, nếu không khớp thì lowercase + strip.
    """
    if not name:
        return ""
    raw = name.strip().lower()
    # Bỏ thông tin lớp trong ngoặc đơn như "(Lớp 1/1)"
    raw = re.sub(r'\(lớp\s*[\d/]+\)', '', raw).strip()

    # Tra bảng alias chính xác
    if raw in _SUBJECT_ALIAS:
        return _SUBJECT_ALIAS[raw]

    # Tra bảng alias dạng "bắt đầu bằng" (cho phép có suffix)
    for alias, canonical in _SUBJECT_ALIAS.items():
        if raw.startswith(alias) or alias.startswith(raw):
            return canonical

    # Fallback: trả về chính nó sau khi lowercase
    return raw


def _build_lesson_block_map(
    lesson_blocks: List[LessonBlock]
) -> Dict[str, List[LessonBlock]]:
    """
    Nhóm lesson blocks theo tên môn đã chuẩn hóa.
    key = normalize_subject_name(mon)
    value = List[LessonBlock] theo thứ tự tiet_ppct tăng dần
    """
    result: Dict[str, List[LessonBlock]] = defaultdict(list)
    for block in lesson_blocks:
        key = normalize_subject_name(block.mon)
        result[key].append(block)
    # Sort theo tiet_ppct
    for key in result:
        result[key].sort(key=lambda b: b.tiet_ppct)
    return dict(result)


def _calculate_initial_cursors(
    tkb_slots: List[TKBSlot],
    block_map: Dict[str, List[LessonBlock]],
    start_tuan: int
) -> Dict[str, int]:
    """
    Tính con trỏ ban đầu cho từng môn.
    Tự động phát hiện File lẻ (chỉ có bài của 1 tuần) vs File gộp (chứa nhiều tuần):
    - Nếu tổng số bài của môn trong block_map <= số tiết/tuần * (start_tuan - 1)
      hoặc tổng số bài <= số tiết/tuần * 1.5:
      => Đây là File Giáo án lẻ của riêng tuần đang chọn -> cursor = 0 (không skip).
    - Ngược lại:
      => File Giáo án gộp cả năm/nhiều tuần -> cursor = count * (start_tuan - 1).
    """
    cursors: Dict[str, int] = defaultdict(int)
    if start_tuan <= 1:
        return cursors

    slots_per_week: Dict[str, int] = defaultdict(int)
    for slot in tkb_slots:
        slots_per_week[normalize_subject_name(slot.mon)] += 1

    for mon_key, count_per_week in slots_per_week.items():
        if count_per_week <= 0:
            continue
        blocks = block_map.get(mon_key, [])
        needed_skip = count_per_week * (start_tuan - 1)

        if len(blocks) <= needed_skip or len(blocks) <= int(count_per_week * 1.5):
            cursors[mon_key] = 0
        else:
            cursors[mon_key] = needed_skip

    return cursors


def _format_display_ppct(block: LessonBlock, slot_idx: int, count_per_week: int, tuan: int) -> str:
    """
    Tính số tiết PPCT hiển thị tự động cộng dồn cho tuần `tuan`.
    - Tiết PPCT chuẩn cho slot_idx trong tuần tuan = count_per_week * (tuan - 1) + slot_idx + 1.
    - Nếu trong file đã ghi sẵn số tiết lớn hơn hoặc bằng mốc của tuần (VD: 5 >= 5) -> dùng số trong file.
    - Ngược lại (file lẻ 1, 2, 3... hoặc rỗng) -> dùng calculated_ppct.
    """
    calculated_ppct = count_per_week * (tuan - 1) + slot_idx + 1
    if block and block.tiet_ppct:
        try:
            val = int(block.tiet_ppct)
            min_tuan_ppct = count_per_week * (tuan - 1) + 1
            if val >= min_tuan_ppct:
                return str(val)
        except (ValueError, TypeError):
            pass
    return str(calculated_ppct)


def generate_schedule(
    tkb_slots: List[TKBSlot],
    lesson_blocks: List[LessonBlock],
    tuan: int = 1,
    start_tuan: int = 1,
) -> Tuple[List[ScheduleRow], List[LessonBlock]]:
    """
    Sinh Bảng Lịch Báo Dạy và danh sách LessonBlock đã sắp xếp theo TKB.

    Trả về:
      - schedule: List[ScheduleRow] — dữ liệu cho Bảng LBD
      - ordered_blocks: List[LessonBlock] — block nội dung xếp đúng thứ tự TKB
        (dùng để append vào file Word gộp)
    """
    block_map = _build_lesson_block_map(lesson_blocks)
    target_tuan = start_tuan if start_tuan > 1 else tuan
    cursors = _calculate_initial_cursors(tkb_slots, block_map, target_tuan)

    slots_per_week: Dict[str, int] = defaultdict(int)
    for slot in tkb_slots:
        slots_per_week[normalize_subject_name(slot.mon)] += 1

    # Sắp xếp TKB: Thứ 2→6, Sáng trước Chiều, Tiết nhỏ trước lớn
    thu_order = {"Hai": 0, "Ba": 1, "Tư": 2, "Năm": 3, "Sáu": 4, "Bảy": 5}
    buoi_order = {"Sáng": 0, "Chiều": 1}
    sorted_slots = sorted(
        tkb_slots,
        key=lambda s: (
            thu_order.get(s.thu, 9),
            buoi_order.get(s.buoi, 9),
            s.tiet_tkb
        )
    )

    schedule: List[ScheduleRow] = []
    ordered_blocks: List[LessonBlock] = []
    subject_slot_counters: Dict[str, int] = defaultdict(int)

    for slot in sorted_slots:
        norm_mon = normalize_subject_name(slot.mon)
        blocks_for_mon = block_map.get(norm_mon, [])
        cursor = cursors[norm_mon]

        slot_idx = subject_slot_counters[norm_mon]
        subject_slot_counters[norm_mon] += 1

        if cursor < len(blocks_for_mon):
            block = blocks_for_mon[cursor]
            cursors[norm_mon] += 1

            count_pw = slots_per_week.get(norm_mon, 1)
            display_ppct = _format_display_ppct(block, slot_idx, count_pw, target_tuan)

            schedule.append(ScheduleRow(
                thu=slot.thu,
                buoi=slot.buoi,
                tiet_tkb=slot.tiet_tkb,
                mon=slot.mon,
                lop=slot.lop or "5/5",
                tiet_ppct=display_ppct,
                ten_bai=block.ten_bai,
                thiet_bi="",
                ghi_chu=""
            ))
            ordered_blocks.append(block)
        else:
            # Môn không có giáo án (chuyên trách hoặc hết block)
            schedule.append(ScheduleRow(
                thu=slot.thu,
                buoi=slot.buoi,
                tiet_tkb=slot.tiet_tkb,
                mon=slot.mon,
                lop=slot.lop or "5/5",
                tiet_ppct="",
                ten_bai="",
                thiet_bi="",
                ghi_chu=""
            ))

    return schedule, ordered_blocks


def generate_multi_week_schedule(
    tkb_slots: List[TKBSlot],
    lesson_blocks: List[LessonBlock],
    start_tuan: int = 1,
    end_tuan: int = 1,
) -> Dict[int, Tuple[List[ScheduleRow], List[LessonBlock]]]:
    """
    Sinh lịch cho nhiều tuần liên tiếp.
    Trả về dict: tuan_num → (schedule, ordered_blocks)
    """
    result = {}
    block_map = _build_lesson_block_map(lesson_blocks)
    cursors = _calculate_initial_cursors(tkb_slots, block_map, start_tuan)

    slots_per_week: Dict[str, int] = defaultdict(int)
    for slot in tkb_slots:
        slots_per_week[normalize_subject_name(slot.mon)] += 1

    thu_order = {"Hai": 0, "Ba": 1, "Tư": 2, "Năm": 3, "Sáu": 4, "Bảy": 5}
    buoi_order = {"Sáng": 0, "Chiều": 1}
    sorted_slots = sorted(
        tkb_slots,
        key=lambda s: (thu_order.get(s.thu, 9), buoi_order.get(s.buoi, 9), s.tiet_tkb)
    )

    for t in range(start_tuan, end_tuan + 1):
        week_schedule: List[ScheduleRow] = []
        week_ordered_blocks: List[LessonBlock] = []
        subject_slot_counters: Dict[str, int] = defaultdict(int)

        for slot in sorted_slots:
            norm_mon = normalize_subject_name(slot.mon)
            blocks_for_mon = block_map.get(norm_mon, [])
            cursor = cursors[norm_mon]

            slot_idx = subject_slot_counters[norm_mon]
            subject_slot_counters[norm_mon] += 1

            if cursor < len(blocks_for_mon):
                block = blocks_for_mon[cursor]
                cursors[norm_mon] += 1
                count_pw = slots_per_week.get(norm_mon, 1)
                display_ppct = _format_display_ppct(block, slot_idx, count_pw, t)
                week_schedule.append(ScheduleRow(
                    thu=slot.thu,
                    buoi=slot.buoi,
                    tiet_tkb=slot.tiet_tkb,
                    mon=slot.mon,
                    lop=slot.lop or "5/5",
                    tiet_ppct=display_ppct,
                    ten_bai=block.ten_bai,
                    thiet_bi="",
                    ghi_chu=""
                ))
                week_ordered_blocks.append(block)
            else:
                week_schedule.append(ScheduleRow(
                    thu=slot.thu,
                    buoi=slot.buoi,
                    tiet_tkb=slot.tiet_tkb,
                    mon=slot.mon,
                    lop=slot.lop or "5/5",
                    tiet_ppct="",
                    ten_bai="",
                    thiet_bi="",
                    ghi_chu=""
                ))

        result[t] = (week_schedule, week_ordered_blocks)

    return result
