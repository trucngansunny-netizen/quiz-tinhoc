# core.py
import os
import json
from docx import Document
from pptx import Presentation
import zipfile

# =============================
# 📦 HÀM ĐỌC TIÊU CHÍ
# =============================
def load_criteria(subject, grade, folder):
    """
    Đọc tiêu chí chấm điểm dựa theo môn và khối.
    """
    subject = subject.lower()
    filename = f"{subject}{grade}.json"
    filepath = os.path.join(folder, filename)
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception as e:
        print(f"Lỗi đọc file tiêu chí {filepath}: {e}")
        return None


# =============================
# 🧩 HÀM CHẤM WORD
# =============================
def grade_word(filepath, criteria):
    """
    Chấm bài Word theo tiêu chí trong file JSON.
    """
    try:
        doc = Document(filepath)
        text = " ".join(p.text for p in doc.paragraphs).lower()
    except Exception as e:
        return None, [f"Lỗi khi đọc file Word: {e}"]

    total = 0
    notes = []
    for c in criteria.get("tieu_chi", []):
        mo_ta = c["mo_ta"]
        diem = c["diem"]
        # Kiểm tra đơn giản: nếu từ khóa trong tiêu chí xuất hiện trong văn bản
        if any(k.lower() in text for k in mo_ta.split()):
            total += diem
            notes.append(f"✅ {mo_ta} (+{diem}đ)")
        else:
            notes.append(f"⚠️ {mo_ta} (chưa đạt)")
    total = min(round(total, 1), 10)
    return total, notes


# =============================
# 🎞️ HÀM CHẤM POWERPOINT
# =============================
def grade_ppt(filepath, criteria):
    """
    Chấm bài PowerPoint theo tiêu chí trong file JSON.
    """
    try:
        prs = Presentation(filepath)
        num_slides = len(prs.slides)
    except Exception as e:
        return None, [f"Lỗi khi đọc file PowerPoint: {e}"]

    total = 0
    notes = []
    text_content = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_content += shape.text.lower() + " "

    for c in criteria.get("tieu_chi", []):
        mo_ta = c["mo_ta"]
        diem = c["diem"]
        # Ví dụ: kiểm tra theo từ khóa trong tiêu chí
        if any(k.lower() in text_content for k in mo_ta.split()) or "trang trình chiếu" in mo_ta.lower() and num_slides >= 3:
            total += diem
            notes.append(f"✅ {mo_ta} (+{diem}đ)")
        else:
            notes.append(f"⚠️ {mo_ta} (chưa đạt)")
    total = min(round(total, 1), 10)
    return total, notes


# =============================
# 🐱‍💻 HÀM CHẤM SCRATCH
# =============================
def grade_scratch(filepath, criteria):
    """
    Chấm file Scratch (.sb3) dựa vào nội dung JSON bên trong.
    """
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            if 'project.json' not in z.namelist():
                return None, ["File .sb3 không hợp lệ (thiếu project.json)."]
            with z.open('project.json') as f:
                project_data = json.load(f)
    except Exception as e:
        return None, [f"Lỗi khi đọc file Scratch: {e}"]

    # Nội dung chính để kiểm tra
    scripts_text = json.dumps(project_data).lower()
    total = 0
    notes = []

    for c in criteria.get("tieu_chi", []):
        mo_ta = c["mo_ta"]
        diem = c["diem"]
        if any(k.lower() in scripts_text for k in mo_ta.split()):
            total += diem
            notes.append(f"✅ {mo_ta} (+{diem}đ)")
        else:
            notes.append(f"⚠️ {mo_ta} (chưa đạt)")
    total = min(round(total, 1), 10)
    return total, notes


# =============================
# 🧾 HÀM ĐẢM BẢO FILE EXCEL TỒN TẠI
# =============================
from openpyxl import Workbook

def ensure_workbook_exists(path):
    if not os.path.exists(path):
        wb = Workbook()
        wb.save(path)


# =============================
# 🪶 HÀM XỬ LÝ TÊN FILE HỌC SINH
# =============================
def pretty_name_from_filename(filename):
    """
    Trích tên học sinh từ tên file. 
    Ví dụ: 'tranminhduc_5a1.docx' → 'Trần Minh Đức'
    """
    name = os.path.splitext(filename)[0]
    name = name.replace("_", " ").replace("-", " ")
    parts = name.split()
    return " ".join(p.capitalize() for p in parts if not p.lower().startswith("lop"))
