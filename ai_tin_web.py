# ai_tin_web.py
# Flask web app — dùng gunicorn ai_tin_web:app
import os
import datetime
import shutil
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# import functions from cham_tieuchi.py (file must be in same folder)
from cham_tieuchi import (
    pretty_name_from_filename,
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch
)

# ---------------- Config ----------------
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
TONGHOP_XLSX = os.path.join(RESULTS_DIR, "tonghop.xlsx")
VISIT_FILE = os.path.join(RESULTS_DIR, "visit_count.txt")
ALLOWED_EXT = {"docx", "pptx", "sb3", "zip"}

CLASSES = [
    "3A1","3A2","3A3","3A4",
    "4A1","4A2","4A3","4A4","4A5",
    "5A1","5A2","5A3","5A4","5A5"
]

AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word", "PowerPoint", "Scratch"],
    5: ["Word", "Scratch"]
}

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "replace-with-secure-key-if-needed"

# ---------------- Utilities ----------------
def allowed_file(filename):
    if "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXT

def read_visit_count():
    try:
        if not os.path.exists(VISIT_FILE):
            with open(VISIT_FILE, "w", encoding="utf-8") as f:
                f.write("0")
            return 0
        with open(VISIT_FILE, "r", encoding="utf-8") as f:
            return int(f.read().strip() or "0")
    except Exception:
        return 0

def increase_visit_count():
    try:
        count = read_visit_count()
        count += 1
        with open(VISIT_FILE, "w", encoding="utf-8") as f:
            f.write(str(count))
        return count
    except Exception:
        return None

def ensure_tonghop_exists():
    if not os.path.exists(TONGHOP_XLSX):
        wb = Workbook()
        # create sheet for each class with header
        first = True
        for cls in CLASSES:
            if first:
                ws = wb.active
                ws.title = cls
                first = False
            else:
                ws = wb.create_sheet(title=cls)
            ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        wb.save(TONGHOP_XLSX)

def append_to_tonghop(class_name, grade, student_name, subject, filename, total_score, remark, crit_summary):
    ensure_tonghop_exists()
    try:
        wb = load_workbook(TONGHOP_XLSX)
        if class_name not in wb.sheetnames:
            ws = wb.create_sheet(title=class_name)
            ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        ws = wb[class_name]
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append([student_name, class_name, grade, subject, filename, float(total_score), now, remark, crit_summary])
        wb.save(TONGHOP_XLSX)
        return True
    except Exception as e:
        print("Error writing tonghop:", e)
        return False

# ---------------- Routes ----------------
@app.route("/static/<path:filename>")
def custom_static(filename):
    return send_from_directory(STATIC_DIR, filename)

@app.route("/", methods=["GET", "POST"])
def index():
    # increase visit count on page load (GET); if POST and successful also show)
    if request.method == "GET":
        increase_visit_count()

    message = None
    result = None
    visit_count = read_visit_count()

    if request.method == "POST":
        grade = request.form.get("grade", "").strip()
        class_name = request.form.get("class", "").strip()
        subject = request.form.get("subject", "").strip()
        fileobj = request.files.get("file")

        # validate
        if not grade or not class_name or not subject:
            message = "Vui lòng chọn Khối, Lớp và Môn trước khi nộp."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)
        try:
            grade_num = int(grade)
        except:
            message = "Khối không hợp lệ."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        # check subject allowed
        if subject not in AVAILABLE_BY_GRADE.get(grade_num, []):
            message = f"Khối {grade_num} không học môn {subject}."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        if not fileobj or fileobj.filename == "":
            message = "Vui lòng chọn tệp để nộp."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        filename = secure_filename(fileobj.filename)
        if not allowed_file(filename):
            message = "Định dạng tệp không được hỗ trợ. Hỗ trợ: .docx .pptx .sb3 .zip"
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        # save temporary file
        tmpdir = tempfile.mkdtemp(prefix="ai_tin_")
        tmp_path = os.path.join(tmpdir, filename)
        try:
            fileobj.save(tmp_path)
        except Exception as e:
            shutil.rmtree(tmpdir, ignore_errors=True)
            message = f"Lỗi khi lưu tệp: {e}"
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        # load criteria (subject_lower, grade_num)
        try:
            crit = load_criteria(subject.lower(), grade_num, CRITERIA_DIR)
        except TypeError:
            # backward compat
            crit = load_criteria(subject.lower(), grade_num)
        if crit is None:
            shutil.rmtree(tmpdir, ignore_errors=True)
            message = f"Chưa có tiêu chí cho {subject} khối {grade}."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        # call correct grader
        total = None
        notes = []
        try:
            if subject == "Word":
                total, notes = grade_word(tmp_path, crit)
            elif subject == "PowerPoint":
                total, notes = grade_ppt(tmp_path, crit)
            elif subject == "Scratch":
                total, notes = grade_scratch(tmp_path, crit)
        except Exception as e:
            total = None
            notes = [f"Lỗi khi chấm: {e}"]

        if total is None:
            shutil.rmtree(tmpdir, ignore_errors=True)
            message = message or "Lỗi khi chấm bài."
            return render_template("index.html", classes=CLASSES, avail_by_grade=AVAILABLE_BY_GRADE, visit_count=visit_count, message=message)

        # determine remark
        try:
            ts = float(total)
        except:
            ts = 0.0
        if ts >= 9.5:
            remark = "Hoàn thành xuất sắc"
        elif ts >= 8.0:
            remark = "Hoàn thành tốt"
        elif ts >= 6.5:
            remark = "Đạt yêu cầu"
        else:
            remark = "Cần cố gắng thêm"

        # append to excel
        student_name = pretty_name_from_filename(filename)
        crit_summary = "; ".join(notes)
        appended = append_to_tonghop(class_name, grade, student_name, subject, filename, total, remark, crit_summary)

        # cleanup tmp
        shutil.rmtree(tmpdir, ignore_errors=True)

        # increase visit on successful POST as well
        increase_visit_count()
        visit_count = read_visit_count()

        result = {
            "student": student_name,
            "class": class_name,
            "grade": grade,
            "subject": subject,
            "file": filename,
            "total": total,
            "remark": remark,
            "notes": notes
        }

    # render template
    return render_template("index.html",
                           classes=CLASSES,
                           avail_by_grade=AVAILABLE_BY_GRADE,
                           visit_count=read_visit_count(),
                           result=result,
                           message=message)

if __name__ == "__main__":
    # for local debug only (gunicorn will import app and run)
    app.run(host="0.0.0.0", port=5000, debug=False)
