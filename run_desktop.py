import uvicorn
import webview
import threading
import sys
import os
import base64

# Cấu hình đường dẫn PyInstaller (_MEIPASS) khi đóng gói .exe
if getattr(sys, 'frozen', False):
    base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

backend_path = os.path.join(base_dir, "backend")
if backend_path not in sys.path:
    sys.path.insert(0, backend_path)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from app.main import app

class DesktopApi:
    def save_file_dialog(self, filename: str, base64_content: str) -> dict:
        """
        Mở Hộp thoại Save File của Windows để người dùng chọn vị trí lưu file (.docx hoặc .zip).
        """
        try:
            ext = os.path.splitext(filename)[1].lower()
            if ext == '.docx':
                file_types = ('Word Document (*.docx)', 'All files (*.*)')
            elif ext == '.zip':
                file_types = ('ZIP Archive (*.zip)', 'All files (*.*)')
            else:
                file_types = ('All files (*.*)',)

            window = webview.windows[0]
            save_path = window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=filename,
                file_types=file_types
            )

            if not save_path:
                return {"success": False, "cancelled": True}

            # Khử toàn bộ tuple/list bọc ngoài để lấy chuỗi đường dẫn chính xác
            target_path = save_path
            while isinstance(target_path, (list, tuple)):
                if len(target_path) == 0:
                    return {"success": False, "cancelled": True}
                target_path = target_path[0]

            target_path = str(target_path)
            if not target_path or target_path == 'None':
                return {"success": False, "cancelled": True}

            file_bytes = base64.b64decode(base64_content)
            with open(target_path, 'wb') as f:
                f.write(file_bytes)

            return {"success": True, "path": target_path}
        except Exception as e:
            return {"success": False, "error": str(e)}

def start_backend():
    # Khởi chạy FastAPI backend ngầm ở cổng 8000
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="error")

if __name__ == "__main__":
    t = threading.Thread(target=start_backend, daemon=True)
    t.start()

    api = DesktopApi()
    webview.create_window(
        title="Auto Lịch Báo Dạy - Lâm Huệ Trung", 
        url="http://127.0.0.1:8000",
        width=1280, 
        height=850,
        resizable=True,
        js_api=api
    )
    webview.start()
