import io
import os
import zipfile
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from app.core.schemas import ScheduleRow

def export_to_docx(
    schedule: List[ScheduleRow], 
    tuan: int = 1, 
    template_path: str = "app/templates/mau_lich.docx"
) -> io.BytesIO:
    """
    Xuất 1 tuần Lịch báo dạy thành file Word (.docx).
    """
    if os.path.exists(template_path):
        try:
            from docxtpl import DocxTemplate
            doc = DocxTemplate(template_path)
            context = {
                "tuan": tuan,
                "danh_sach": [row.model_dump() for row in schedule]
            }
            doc.render(context)
            stream = io.BytesIO()
            doc.save(stream)
            stream.seek(0)
            return stream
        except Exception as e:
            print(f"Lỗi docxtpl: {e}")

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    _append_week_table(doc, schedule, tuan)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

def export_multi_week_to_docx(multi_schedule: Dict[int, List[ScheduleRow]]) -> io.BytesIO:
    """
    Xuất nhiều tuần vào 1 file Word duy nhất, phân trang đẹp mắt giữa các tuần.
    """
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    sorted_weeks = sorted(multi_schedule.keys())
    for idx, tuan_num in enumerate(sorted_weeks):
        if idx > 0:
            doc.add_page_break()
        _append_week_table(doc, multi_schedule[tuan_num], tuan_num)

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

def export_weeks_to_zip(multi_schedule: Dict[int, List[ScheduleRow]]) -> io.BytesIO:
    """
    Xuất nhiều tuần thành nén ZIP chứa các file Word riêng lẻ cho từng tuần.
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for tuan_num, schedule in multi_schedule.items():
            docx_stream = export_to_docx(schedule, tuan=tuan_num)
            zip_file.writestr(f"Lich_Bao_Day_Tuan_{tuan_num}.docx", docx_stream.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

def _append_week_table(doc: Document, schedule: List[ScheduleRow], tuan: int):
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"LỊCH BÁO DẠY - TUẦN {tuan}")
    run.font.name = 'Times New Roman'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph()

    headers = ["Thứ", "Buổi", "Tiết TKB", "Môn học", "Tiết PPCT", "Tên bài dạy / Nội dung", "Thiết bị", "Ghi chú"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    for idx, header_text in enumerate(headers):
        hdr_cells[idx].text = header_text
        for p in hdr_cells[idx].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.bold = True
                r.font.size = Pt(11)

    for cell in hdr_cells:
        tcPr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'EBF2F7')
        tcPr.append(shd)

    for row_data in schedule:
        row_cells = table.add_row().cells
        data_values = [
            row_data.thu,
            row_data.buoi,
            str(row_data.tiet_tkb),
            row_data.mon,
            row_data.tiet_ppct,
            row_data.ten_bai,
            row_data.thiet_bi,
            row_data.ghi_chu
        ]
        for idx, val in enumerate(data_values):
            row_cells[idx].text = str(val)
            for p in row_cells[idx].paragraphs:
                if idx in [0, 1, 2, 4]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
