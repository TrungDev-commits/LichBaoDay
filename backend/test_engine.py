import sys
import os

# Set UTF-8 encoding cho console Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Thêm đường dẫn backend vào sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from app.core.schemas import TKBSlot, BaiHoc
from app.core.engine import generate_schedule
from app.exporters.docx_export import export_to_docx

def test_auto_schedule():
    print("--- 1. BẮT ĐẦU MOCK DỮ LIỆU TKB VÀ PPCT ---")
    
    # Mock TKB 1 Tuần (Tiếng Việt, Toán, Hoạt động trải nghiệm, Đạo đức)
    tkb_slots = [
        TKBSlot(thu="Hai", buoi="Sáng", tiet_tkb=1, mon="Chào cờ"),
        TKBSlot(thu="Hai", buoi="Sáng", tiet_tkb=2, mon="T.Việt"),
        TKBSlot(thu="Hai", buoi="Sáng", tiet_tkb=3, mon="Toán"),
        TKBSlot(thu="Hai", buoi="Sáng", tiet_tkb=4, mon="Hoạt động trải nghiệm"),
        
        TKBSlot(thu="Ba", buoi="Sáng", tiet_tkb=1, mon="Toán"),
        TKBSlot(thu="Ba", buoi="Sáng", tiet_tkb=2, mon="T.Việt"),
        TKBSlot(thu="Ba", buoi="Sáng", tiet_tkb=3, mon="T.Việt"),
        TKBSlot(thu="Ba", buoi="Sáng", tiet_tkb=4, mon="Đạo đức"),
        
        TKBSlot(thu="Tư", buoi="Sáng", tiet_tkb=1, mon="T.Việt"),
        TKBSlot(thu="Tư", buoi="Sáng", tiet_tkb=2, mon="Toán"),
        TKBSlot(thu="Tư", buoi="Sáng", tiet_tkb=3, mon="Tự nhiên và Xã hội"),
        
        TKBSlot(thu="Năm", buoi="Sáng", tiet_tkb=1, mon="Toán"),
        TKBSlot(thu="Năm", buoi="Sáng", tiet_tkb=2, mon="T.Việt"),
        TKBSlot(thu="Năm", buoi="Sáng", tiet_tkb=3, mon="Tiếng Anh"),
        
        TKBSlot(thu="Sáu", buoi="Sáng", tiet_tkb=1, mon="Toán"),
        TKBSlot(thu="Sáu", buoi="Sáng", tiet_tkb=2, mon="T.Việt"),
        TKBSlot(thu="Sáu", buoi="Sáng", tiet_tkb=3, mon="Sinh hoạt lớp"),
    ]

    # Mock Phân phối chương trình
    ppct_items = [
        # Tiếng Việt
        BaiHoc(mon="Tiếng Việt", tiet_ppct=1, ten_bai="Tập đọc: Cậu học sinh mới", thiet_bi="Tranh minh họa"),
        BaiHoc(mon="Tiếng Việt", tiet_ppct=2, ten_bai="Chính tả: Nghe - viết: Cậu học sinh mới", thiet_bi="Bảng phụ"),
        BaiHoc(mon="Tiếng Việt", tiet_ppct=3, ten_bai="Luyện từ và câu: Mở rộng vốn từ Trường học", thiet_bi="Phiếu học tập"),
        BaiHoc(mon="Tiếng Việt", tiet_ppct=4, ten_bai="Tập làm văn: Kể lại buổi đầu đi học", thiet_bi="Gợi ý dàn ý"),
        BaiHoc(mon="Tiếng Việt", tiet_ppct=5, ten_bai="Tập đọc: Nhớ lại buổi đầu đi học", thiet_bi="Tranh minh họa"),
        BaiHoc(mon="Tiếng Việt", tiet_ppct=6, ten_bai="Luyện nói: Trường học của em", thiet_bi=""),
        
        # Toán
        BaiHoc(mon="Toán", tiet_ppct=1, ten_bai="Đọc, viết, so sánh các số có ba chữ số", thiet_bi="Bộ đồ dùng Toán 3"),
        BaiHoc(mon="Toán", tiet_ppct=2, ten_bai="Cộng các số có ba chữ số (không nhớ)", thiet_bi="Bảng nhóm"),
        BaiHoc(mon="Toán", tiet_ppct=3, ten_bai="Luyện tập cộng các số có ba chữ số", thiet_bi=""),
        BaiHoc(mon="Toán", tiet_ppct=4, ten_bai="Trừ các số có ba chữ số (không nhớ)", thiet_bi="Bộ đồ dùng Toán 3"),
        BaiHoc(mon="Toán", tiet_ppct=5, ten_bai="Luyện tập tổng hợp", thiet_bi=""),

        # Đạo đức
        BaiHoc(mon="Đạo đức", tiet_ppct=1, ten_bai="Bài 1: Kính yêu Bác Hồ (Tiết 1)", thiet_bi="Video Bác Hồ với thiếu nhi"),

        # HĐTN
        BaiHoc(mon="Hoạt động trải nghiệm", tiet_ppct=1, ten_bai="Sinh hoạt dưới cờ: Khai giảng năm học mới", thiet_bi="Cờ, hoa"),

        # TNXH
        BaiHoc(mon="Tự nhiên và Xã hội", tiet_ppct=1, ten_bai="Bài 1: Họ hàng nội, ngoại (Tiết 1)", thiet_bi="Sơ đồ gia đình"),
    ]

    print("--- 2. CHẠY THUẬT TOÁN GHÉP LỊCH ---")
    schedule = generate_schedule(tkb_slots, ppct_items)
    
    for row in schedule:
        print(f"[{row.thu} - {row.buoi} - Tiết {row.tiet_tkb}] {row.mon:<22} | Tiết PPCT: {row.tiet_ppct:<3} | {row.ten_bai}")

    print("\n--- 3. XUẤT THỬ FILE WORD ---")
    output_stream = export_to_docx(schedule, tuan=1)
    
    test_output_path = "backend/test_output_lich_bao_day.docx"
    with open(test_output_path, "wb") as f:
        f.write(output_stream.getvalue())
        
    print(f"✅ ĐÃ XUẤT THÀNH CÔNG FILE TEST WORD TẠI: {test_output_path}")

if __name__ == "__main__":
    test_auto_schedule()
