## Đóng gói Desktop App hiện đại bằng REACT

### 1. Kiến trúc công nghệ (Tech Stack)

- **Backend Engine:** **Python 3.11+ (FastAPI)**
- `fastapi` & `uvicorn`: Xây dựng REST API siêu nhẹ, tốc độ cao, tự sinh Swagger docs để test ngay mà chưa cần frontend.
- `pydantic`: Định nghĩa Data Contract (Schema) chuẩn hóa dữ liệu.
- `openpyxl`: Bóc tách và thao tác với file Excel.
- `python-docx`: Đọc file Word giáo án.
- `docxtpl`: Điền dữ liệu vào template Word `.docx` giữ nguyên 100% định dạng trang in.

- **Frontend (Nếu làm UI):** **Vue 3 (hoặc React) + Tailwind CSS + Vite**
- Giao diện dạng **Single-Page Application (SPA)** không cần đăng nhập.
- Có thể deploy hoàn toàn miễn phí trên Netlify / Vercel.

- **Môi trường Deploy Backend:** **Render.com (Web Service)** hoặc chạy local qua Docker.

---

### 2. Thiết kế luồng dữ liệu & Thư mục dự án

```text
auto-lich-bao-day/
├── backend/
│   ├── app/
│   │   ├── core/
│   │   │   ├── engine.py       # Core logic ghép TKB + Giáo án
│   │   │   └── schemas.py      # Pydantic models (Data Contract)
│   │   ├── parsers/
│   │   │   ├── tkb_parser.py   # Parse Excel/Word TKB -> TKBSlot
│   │   │   └── ppct_parser.py  # Parse Giáo án/PPCT -> BaiHoc
│   │   ├── exporters/
│   │   │   └── docx_export.py  # Render docxtpl -> file Lịch báo dạy .docx
│   │   ├── templates/
│   │   │   └── mau_lich.docx   # File Word mẫu chuẩn của Bộ/Sở GD
│   │   └── main.py             # FastAPI entrypoint
│   └── requirements.txt
└── frontend/ (Tùy chọn)
    ├── src/
    │   ├── App.vue             # Giao diện kéo-thả file + preview table
    │   └── ...

```

---

### 3. Chi tiết triển khai Backend (Core Engine)

#### A. Data Contract (`core/schemas.py`)

Chuẩn hóa dữ liệu trung gian để không bị phụ thuộc vào định dạng file gốc:

```python
from pydantic import BaseModel
from typing import Optional, List

class TKBSlot(BaseModel):
    thu: str           # "Hai", "Ba", "Tư", ...
    buoi: str          # "Sáng", "Chiều"
    tiet_tkb: int      # 1, 2, 3, 4, 5
    mon: str           # "Toán", "Tiếng Việt", "Đạo đức"
    lop: Optional[str] = "3A"

class BaiHoc(BaseModel):
    mon: str           # "Toán"
    tiet_ppct: int     # 15
    ten_bai: str       # "Bảng nhân 6 (Tiết 1)"
    thiet_bi: Optional[str] = "Bộ đồ dùng Toán 3"

class ScheduleRow(BaseModel):
    thu: str
    buoi: str
    tiet_tkb: int
    mon: str
    tiet_ppct: str
    ten_bai: str
    thiet_bi: str

```

#### B. Thuật toán ghép lịch (`core/engine.py`)

Sử dụng hàng đợi FIFO (First In First Out) theo từng môn:

```python
from collections import defaultdict
from typing import List
from .schemas import TKBSlot, BaiHoc, ScheduleRow

def generate_schedule(tkb_slots: List[TKBSlot], ppct_items: List[BaiHoc]) -> List[ScheduleRow]:
    # 1. Tạo hàng đợi bài học theo từng môn
    subject_queues = defaultdict(list)
    for item in sorted(ppct_items, key=lambda x: x.tiet_ppct):
        subject_queues[item.mon].append(item)

    schedule: List[ScheduleRow] = []

    # 2. Quét qua từng tiết trong TKB của tuần
    for slot in tkb_slots:
        current_lesson = None
        if subject_queues[slot.mon]:
            current_lesson = subject_queues[slot.mon].pop(0)

        schedule.append(ScheduleRow(
            thu=slot.thu,
            buoi=slot.buoi,
            tiet_tkb=slot.tiet_tkb,
            mon=slot.mon,
            tiet_ppct=str(current_lesson.tiet_ppct) if current_lesson else "",
            ten_bai=current_lesson.ten_bai if current_lesson else "(Chưa có phân phối)",
            thiet_bi=current_lesson.thiet_bi if current_lesson else ""
        ))

    return schedule

```

#### C. Render Template Word (`exporters/docx_export.py`)

Sử dụng `docxtpl`. Bạn chuẩn bị sẵn file `mau_lich.docx` với bảng chứa thẻ Jinja:
`{% for r in danh_sach %} ... {{ r.mon }} ... {{ r.ten_bai }} ... {% endfor %}`.

```python
from docxtpl import DocxTemplate
from typing import List
from .schemas import ScheduleRow
import io

def export_to_docx(schedule: List[ScheduleRow], tuan: int, template_path: str = "templates/mau_lich.docx") -> io.BytesIO:
    doc = DocxTemplate(template_path)
    context = {
        "tuan": tuan,
        "danh_sach": [row.model_dump() for row in schedule]
    }
    doc.render(context)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

```

#### D. API Endpoint (`main.py`)

FastAPI cho phép test trực tiếp file upload qua Swagger UI (`/docs`) mà **chưa cần viết frontend**:

```python
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from app.core.engine import generate_schedule
from app.exporters.docx_export import export_to_docx
from app.parsers.tkb_parser import parse_tkb_excel
from app.parsers.ppct_parser import parse_ppct_docx

app = FastAPI(title="Lịch Báo Dạy Engine")

@app.post("/api/generate")
async def generate(
    tuan: int = Form(1),
    file_tkb: UploadFile = File(...),
    file_ppct: UploadFile = File(...)
):
    # 1. Parse dữ liệu đầu vào
    tkb_data = parse_tkb_excel(await file_tkb.read())
    ppct_data = parse_ppct_docx(await file_ppct.read())

    # 2. Xử lý ghép bài
    schedule = generate_schedule(tkb_data, ppct_data)

    # 3. Render file Word trả về cho client tải xuống
    file_output = export_to_docx(schedule, tuan=tuan)

    return StreamingResponse(
        file_output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename=Lich_Bao_Day_Tuan_{tuan}.docx"}
    )

```

---

### 4. Thiết kế Giao diện (Nếu triển khai Web UI)

Nếu bạn làm giao diện cho giáo viên dùng, mô hình tối ưu nhất là **Giao diện 3 bước (Wizard Step)**:

1. **Bước 1: Tải lên (Dropzone)**

- Ô 1: Kéo thả file Thời khóa biểu (`.xlsx`).
- Ô 2: Kéo thả file Kế hoạch dạy học / Giáo án (`.docx` hoặc `.xlsx`).
- Chọn Tuần bắt đầu (ví dụ: Tuần 1).

2. **Bước 2: Bảng xem trước (Live Editable Table)**

- Sau khi upload, Backend trả về JSON mảng các tiết học.
- Giao diện hiển thị bảng Lịch báo dạy dạng bảng tương tác.
- Giáo viên có thể click vào sửa nhanh tên bài hoặc đồ dùng dạy học nếu tuần đó có lịch nghỉ lễ/đổi tiết.

3. **Bước 3: Tải về**

- Nút **"Xuất file Word (.docx)"**: Gửi dữ liệu đã kiểm duyệt lên backend để xuất file hoàn chỉnh.

---

### 5. Lộ trình thực hiện từng bước (Step-by-Step)

1. **Giai đoạn 1 (1–2 ngày - Không cần UI):**

- Tạo repo Backend FastAPI.
- Tạo 1 file Word `mau_lich.docx` có bảng biểu chuẩn.
- Viết Schema, Core Engine và Exporter với mock data cứng.
- Dùng giao diện `/docs` của FastAPI để upload mock file và kiểm tra xem file Word tải về có đẹp và đúng ô hay không.

2. **Giai đoạn 2 (Khi có file thực tế):**

- Viết module `tkb_parser.py` và `ppct_parser.py` dựa trên cấu trúc file thật của trường/giáo viên.

3. **Giai đoạn 3 (Nếu muốn đưa vào sử dụng rộng rãi):**

- Viết giao diện Single-Page bằng Vue/React.
- Deploy Backend lên Render (hoặc Docker VPS), Frontend lên Netlify.
