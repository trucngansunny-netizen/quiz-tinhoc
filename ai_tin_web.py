import os
import json
import shutil
import tempfile
import datetime
from flask import Flask, render_template, request, send_from_directory, flash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# ---- Import từ file chấm ----
from cham_tieuchi import (
    pretty_name_from_filename,
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch
)

# ---------------- CONFIG ----------------
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")
VISIT_FILE = os.path.join(RESULTS_DIR, "visit_count.txt")
ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

# Các lớp trong trường
CLASSES = [
    "3A1", "3A2", "3A3", "3A4",
    "4A1", "4A2", "4A3", "4A4", "4A5",
    "5A1", "5A2", "5A3", "5A4", "5A5"
]

AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word", "PowerPoint", "Scratch"],
    5: ["Word", "Scratch"]
}

# Tạo thư mục nếu chưa có
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

# ---------------- Flask app ----------------
app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "ai_tin_secret_key_replace_if_needed"

# ---------------- Utilities ----------------
def allowed_file(filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    return ext in ALLOWED_EXT


def increase_visit():
    """Tăng lượt truy cập và lưu vào file"""
    try:
        if not os.path.exists(VISIT_FILE):
            with open(VISIT_FILE, "w", encoding="utf-8") as f:
                f.write("0")
        with open(VISIT_FILE, "r+", encoding="utf-8") as f:
            c = int(f.read().strip() or "0")
            c += 1
            f.seek(0)
            f.write(str(c))
            f.truncate()
        return c
    except Exception:
        return 0


def read_visit():
    try:
        if not os.path.exists(VISIT_FILE):
            return 0
        with open(VISIT_FILE, "r", encoding="utf-8") as f:
            return int(f.read().strip() or "0")
    except:
        return 0


def ensure_tonghop():
    """Tạo file tổng hợp nếu chưa có"""
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = CLASSES[0]
        ws.append([
            "Họ và tên", "Lớp", "Khối", "Môn học",
            "Tên tệp", "Điểm", "Ngày chấm", "Nhận xét", "Chi tiết"
        ])
        for cls in CLASSES[1:]:
            w = wb.create_sheet(title=cls)
            w.append([
                "Họ và tên", "Lớp", "Khối", "Môn học",
                "Tên tệp", "Điểm", "Ngày chấm", "Nhận xét", "Chi tiết"
            ])
        wb.save(TONGHOP_FILE)


def append_result(class_name, grade, student_name, subject, filename, total, remark, notes):
    """Ghi kết quả vào Excel"""
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append([
            "Họ và tên", "Lớp", "Khối", "Môn học",
            "Tên tệp", "Điểm", "Ngày chấm", "Nhận xét", "Chi tiết"
        ])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    ws.append([
        student_name, class_name, grade, subject,
        filename, total, now, remark, "; ".join(notes)
    ])
    wb.save(TONGHOP_FILE)


# ---------------- Routes ----------------
@app.route("/static/<path:filename>")
def custom_static(filename):
    return send_from_directory(STATIC_DIR, filename)


@app.route("/", methods=["GET", "POST"])
def index():
    visit_count = increase_visit()
    message = None
    result = None

    if request.method == "POST":
        grade = request.form.get("grade")
        cls = request.form.get("class")
        subject = request.form.get("subject")
        uploaded = request.files.get("file")

        if not grade or not cls or not subject:
            message = "⚠️ Vui lòng chọn đầy đủ Khối, Lớp và Môn học."
        elif uploaded is None or uploaded.filename == "":
            message = "⚠️ Vui lòng chọn tệp để chấm."
        else:
            filename = secure_filename(uploaded.filename)
            if not allowed_file(filename):
                message = "⚠️ Chỉ hỗ trợ file: .docx, .pptx, .sb3, .zip"
            else:
                tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                file_path = os.path.join(tmp_dir, filename)
                uploaded.save(file_path)

                # Nạp tiêu chí
                try:
                    criteria = load_criteria(subject, int(grade), CRITERIA_DIR)
                except Exception as e:
                    message = f"Lỗi đọc tiêu chí: {e}"
                    criteria = None

                if not criteria:
                    message = f"⚠️ Chưa có tiêu chí cho {subject} khối {grade}."
                else:
                    try:
                        if subject.lower().startswith("word"):
                            total, notes = grade_word(file_path, criteria)
                        elif subject.lower().startswith("power"):
                            total, notes = grade_ppt(file_path, criteria)
                        elif subject.lower().startswith("scratch"):
                            total, notes = grade_scratch(file_path, criteria)
                        else:
                            total, notes = None, ["Môn học không hợp lệ."]
                    except Exception as e:
                        total, notes = None, [f"Lỗi khi chấm: {e}"]

                    if total is not None:
                        # Nhận xét tự động
                        if total >= 9.5:
                            remark = "Hoàn thành xuất sắc"
                        elif total >= 8.0:
                            remark = "Hoàn thành tốt"
                        elif total >= 6.5:
                            remark = "Đạt yêu cầu"
                        else:
                            remark = "Cần cố gắng thêm"

                        append_result(
                            cls, grade, pretty_name_from_filename(filename),
                            subject, filename, total, remark, notes
                        )

                        result = {
                            "student": pretty_name_from_filename(filename),
                            "class": cls,
                            "grade": grade,
                            "subject": subject,
                            "file": filename,
                            "total": total,
                            "remark": remark,
                            "notes": notes
                        }
                    else:
                        message = "⚠️ Không thể chấm bài. Kiểm tra lại file."
                shutil.rmtree(tmp_dir, ignore_errors=True)

    return render_template(
        "index.html",
        classes=CLASSES,
        avail_by_grade=AVAILABLE_BY_GRADE,
        visit_count=visit_count,
        result=result,
        message=message
    )


# ---------------- Run local ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
