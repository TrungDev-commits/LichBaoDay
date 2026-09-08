from collections import defaultdict
from typing import List, Dict, Optional
import re

from app.core.schemas import TKBSlot, BaiHoc, ScheduleRow

def normalize_subject_name(name: str) -> str:
    """
    Chuẩn hóa tên môn học để ghép chính xác giữa TKB và PPCT.
    """
    if not name:
        return ""
    s = name.strip().lower()
    
    s = re.sub(r'^t\.?\s*việt', 'tiếng việt', s)
    s = re.sub(r'^t\.?\s*anh', 'tiếng anh', s)
    s = re.sub(r'^hđtn', 'hoạt động trải nghiệm', s)
    s = re.sub(r'^shtt', 'sinh hoạt tập thể', s)
    s = re.sub(r'^gdtch', 'giáo dục thể chất', s)
    s = re.sub(r'^đđ', 'đạo đức', s)
    s = re.sub(r'^tnxh', 'tự nhiên và xã hội', s)
    s = re.sub(r'^ls&đl', 'lich sử và địa lý', s)
    s = re.sub(r'^cn', 'công nghệ', s)
    s = re.sub(r'^tin', 'tin học', s)
    s = re.sub(r'^mth', 'mĩ thuật', s)
    s = re.sub(r'^ân', 'âm nhạc', s)
    s = re.sub(r'^tăng cường toán', 'tăng cường toán', s)
    s = re.sub(r'^tăng cường tv', 'tăng cường tiếng việt', s)
    
    # Bỏ lớp trong ngoặc đơn (ví dụ "(Lớp 1/1)")
    s = re.sub(r'\(.*?\)', '', s).strip()
    return s

def generate_multi_week_schedule(
    tkb_slots: List[TKBSlot], 
    ppct_items: List[BaiHoc],
    start_tuan: int = 1,
    end_tuan: int = 1
) -> Dict[int, List[ScheduleRow]]:
    """
    Thuật toán ghép Lịch báo dạy liên tuần (từ start_tuan đến end_tuan).
    Bỏ qua các bài học của các tuần trước start_tuan và nối tiếp tuần tự qua từng tuần.
    """
    # 1. Nhóm bài học theo môn
    subject_queues: Dict[str, List[BaiHoc]] = defaultdict(list)
    sorted_ppct = sorted(ppct_items, key=lambda x: x.tiet_ppct)
    for item in sorted_ppct:
        key = normalize_subject_name(item.mon)
        subject_queues[key].append(item)

    # 2. Xử lý bỏ qua (skipping) số bài học cho các tuần trước start_tuan
    if start_tuan > 1:
        for _ in range(1, start_tuan):
            for slot in tkb_slots:
                norm_mon = normalize_subject_name(slot.mon)
                matched_key = None
                if norm_mon in subject_queues and subject_queues[norm_mon]:
                    matched_key = norm_mon
                else:
                    for key in subject_queues:
                        if key and (key in norm_mon or norm_mon in key) and subject_queues[key]:
                            matched_key = key
                            break
                if matched_key:
                    subject_queues[matched_key].pop(0)

    # 3. Duyệt ghép tuần tự từ start_tuan đến end_tuan
    result: Dict[int, List[ScheduleRow]] = {}

    for t in range(start_tuan, end_tuan + 1):
        week_schedule: List[ScheduleRow] = []
        for slot in tkb_slots:
            raw_mon = slot.mon
            norm_mon = normalize_subject_name(raw_mon)

            current_lesson: Optional[BaiHoc] = None
            matched_key = None
            if norm_mon in subject_queues and subject_queues[norm_mon]:
                matched_key = norm_mon
            else:
                for key in subject_queues:
                    if key and (key in norm_mon or norm_mon in key) and subject_queues[key]:
                        matched_key = key
                        break

            if matched_key:
                current_lesson = subject_queues[matched_key].pop(0)

            week_schedule.append(ScheduleRow(
                thu=slot.thu,
                buoi=slot.buoi,
                tiet_tkb=slot.tiet_tkb,
                mon=slot.mon,
                tiet_ppct=str(current_lesson.tiet_ppct) if current_lesson else "",
                ten_bai=current_lesson.ten_bai if current_lesson else "(Chưa có PPCT)",
                thiet_bi=current_lesson.thiet_bi if current_lesson else "",
                ghi_chu=current_lesson.ghi_chu if current_lesson else ""
            ))
        result[t] = week_schedule

    return result

def generate_schedule(
    tkb_slots: List[TKBSlot], 
    ppct_items: List[BaiHoc],
    start_tuan: int = 1
) -> List[ScheduleRow]:
    res = generate_multi_week_schedule(tkb_slots, ppct_items, start_tuan=start_tuan, end_tuan=start_tuan)
    return res.get(start_tuan, [])
