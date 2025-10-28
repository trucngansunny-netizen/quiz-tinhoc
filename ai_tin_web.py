# ai_tin_web.py  — Flask WSGI app (gunicorn ai_tin_web:app)
import os
import json
import shutil
import tempfile
import datetime
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# Import hàm chấm của cô (giữ nguyên file cham_tieuchi.py mà cô đã có)
# cham_tieuchi.py phải nằm cùng thư mục với ai_tin_web.py
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
VISIT_FILE = os.path.join(RESULTS_DIR, "visits.txt")
ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

# Danh sách 14 lớp (cô yêu cầu)
CLASSES = [
    "3A1","3A2","3A3","3A4",
    "4A1","4A2","4A3","4A4","4A5",
    "5A1","5A2","5A3","5A4","5A5"
]

# Môn hợp lệ theo khối
AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word", "PowerPoint", "Scratch"],
    5: ["Word", "Scratch"]
}

# Tạo các thư mục cần thiết
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

# ---------------- Flask app ----------------
app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "ai_tin_secret_key_replace_if_needed"  # chỉ để flash message

# ---------------- Utilities ----------------
def allowed_file(filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    return ext in ALLOWED_EXT

def increase_visit():
    # Tăng lượt truy cập (tạo file nếu chưa có)
    try:
        if not os.path.exists(RESULTS_DIR):
            os.makedirs(RESULTS_DIR, exist_ok=True)
        if not os.path.exists(VISIT_FILE):
            with open(VISIT_FILE, "w", encoding="utf-8") as f:
                f.write("0")
        with open(VISIT_FILE, "r+", encoding="utf-8") as f:
            try:
                c = int(f.read().strip() or "0")
            except:
                c = 0
            c += 1
            f.seek(0); f.write(str(c)); f.truncate()
        return c
    except Exception:
        return None

def read_visit():
    try:
        if not os.path.exists(VISIT_FILE):
            return 0
        with open(VISIT_FILE, "r", encoding="utf-8") as f:
            return int(f.read().strip() or "0")
    except:
        return 0

def ensure_tonghop():
    # Tạo file tonghop.xlsx nếu chưa có, tạo 14 sheet
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        # remove default sheet
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
            ws = wb.active
            ws.title = CLASSES[0]  # đổi tên sheet mặc định thành sheet đầu
            ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
            # tạo các sheet còn lại
            for cls in CLASSES[1:]:
                ws2 = wb.create_sheet(title=cls)
                ws2.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        else:
            # nếu có cấu trúc khác
            for cls in CLASSES:
                if cls not in wb.sheetnames:
                    ws = wb.create_sheet(title=cls)
                    ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        wb.save(TONGHOP_FILE)

def append_result(class_name, grade, student_name, subject, filename, total, remark, crit_summary):
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student_name, class_name, grade, subject, filename, float(total), now, remark, crit_summary])
    wb.save(TONGHOP_FILE)

# ---------------- Routes ----------------

@app.route("/static/<path:filename>")
def custom_static(filename):
    # Serve static files if needed
    return send_from_directory(STATIC_DIR, filename)

@app.route("/", methods=["GET", "POST"])
def index():
    # tăng lượt truy cập khi open trang
    visit_count = increase_visit() or read_visit()

    message = None
    result = None
    criteria_list = None
    selected_grade = None
    selected_class = None
    selected_subject = None

    if request.method == "POST":
        # lấy thông tin form
        selected_grade = request.form.get("grade")
        selected_class = request.form.get("class")
        selected_subject = request.form.get("subject")  # Word / PowerPoint / Scratch
        uploaded = request.files.get("file")

        # kiểm tra hợp lệ
        if not selected_grade or not selected_class or not selected_subject:
            message = "Vui lòng chọn Khối, Lớp và Môn trước khi nộp."
        elif uploaded is None or uploaded.filename == "":
            message = "Vui lòng chọn tệp để nộp."
        else:
            try:
                grade_num = int(selected_grade)
            except:
                message = "Khối không hợp lệ."
                grade_num = None

            if grade_num and selected_subject not in AVAILABLE_BY_GRADE.get(grade_num, []):
                message = f"Khối {grade_num} không học môn {selected_subject}."
            else:
                # lưu file tạm
                filename = secure_filename(uploaded.filename)
                if not allowed_file(filename):
                    message = "Định dạng tệp không được hỗ trợ. Hỗ trợ: .docx .pptx .sb3 .zip"
                else:
                    tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                    tmp_path = os.path.join(tmp_dir, filename)
                    uploaded.save(tmp_path)

                    # nạp tiêu chí từ thư mục criteria — using load_criteria(subject, grade, criteria_folder) signature from cham_tieuchi
                    # Note: cham_tieuchi.load_criteria signature in earlier messages: load_criteria(subject, grade, criteria_folder="criteria")
                    try:
                        criteria = load_criteria(selected_subject.lower(), int(selected_grade), CRITERIA_DIR)
                    except TypeError:
                        # some earlier versions of load_criteria might accept 2 args (subject, grade)
                        criteria = load_criteria(selected_subject.lower(), int(selected_grade))

                    if not criteria:
                        message = f"Chưa có tiêu chí cho {selected_subject} khối {selected_grade}."
                        shutil.rmtree(tmp_dir, ignore_errors=True)
                    else:
                        # chấm theo môn
                        total = None
                        notes = []
                        try:
                            if selected_subject == "Word":
                                total, notes = grade_word(tmp_path, criteria)
                            elif selected_subject == "PowerPoint":
                                total, notes = grade_ppt(tmp_path, criteria)
                            elif selected_subject == "Scratch":
                                total, notes = grade_scratch(tmp_path, criteria)
                        except Exception as e:
                            message = f"Lỗi khi chấm: {e}"
                            notes = [str(e)]
                            total = None

                        if total is None:
                            message = message or "Lỗi khi chấm bài (không nhận diện được nội dung)."
                            shutil.rmtree(tmp_dir, ignore_errors=True)
                        else:
                            # convert notes list to summary string (✅/❌)
                            crit_summary = "; ".join(notes)
                            # remark auto
                            try:
                                total_float = float(total)
                            except:
                                total_float = float(round(total, 2))
                            if total_float >= 9.5:
                                remark = "Hoàn thành xuất sắc"
                            elif total_float >= 8.0:
                                remark = "Hoàn thành tốt"
                            elif total_float >= 6.5:
                                remark = "Đạt yêu cầu"
                            else:
                                remark = "Cần cố gắng thêm"

                            # Lưu vào excel tổng
                            append_result(selected_class, selected_grade, pretty_name_from_filename(filename),
                                          selected_subject, filename, total_float, remark, crit_summary)

                            # Dọn tmp
                            shutil.rmtree(tmp_dir, ignore_errors=True)

                            # Kết quả hiển thị cho học sinh (theo yêu cầu: bây giờ hiển thị KẾT QUẢ luôn)
                            result = {
                                "student": pretty_name_from_filename(filename),
                                "class": selected_class,
                                "grade": selected_grade,
                                "subject": selected_subject,
                                "file": filename,
                                "total": total_float,
                                "remark": remark,
                                "notes": notes
                            }

    # chuẩn dữ liệu lên template
    return render_template("index.html",
                           classes=CLASSES,
                           avail_by_grade=AVAILABLE_BY_GRADE,
                           visit_count=read_visit(),
                           result=result,
                           message=message)

# ---------------- Run (for local debug only) ----------------
if __name__ == "__main__":
    # chạy debug local (không ảnh hưởng render/gunicorn)
    app.run(host="0.0.0.0", port=5000, debug=False)
