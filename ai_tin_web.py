# ai_tin_web.py
import os
import json
import shutil
import tempfile
import datetime
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

from cham_tieuchi import (
    pretty_name_from_filename,
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch
)

# Config
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")
VISIT_FILE = os.path.join(RESULTS_DIR, "visits.txt")
ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

CLASSES = ["3A1","3A2","3A3","3A4",
           "4A1","4A2","4A3","4A4","4A5",
           "5A1","5A2","5A3","5A4","5A5"]

AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word","PowerPoint","Scratch"],
    5: ["Word","Scratch"]
}

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "replace_with_a_random_secret"

# Utilities for visits and excel
def read_visit():
    try:
        if not os.path.exists(VISIT_FILE): return 0
        return int(open(VISIT_FILE,"r",encoding="utf-8").read().strip() or "0")
    except Exception:
        return 0

def increase_visit():
    c = read_visit() + 1
    with open(VISIT_FILE, "w", encoding="utf-8") as f:
        f.write(str(c))
    return c

def ensure_tonghop():
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        # remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        for cls in CLASSES:
            ws = wb.create_sheet(title=cls)
            ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm"])
        wb.save(TONGHOP_FILE)

def append_result(class_name, grade, student_name, subject, filename, total):
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student_name, class_name, grade, subject, filename, float(total), now])
    wb.save(TONGHOP_FILE)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".",1)[1].lower() in ALLOWED_EXT

@app.route("/static/<path:filename>")
def custom_static(filename):
    return send_from_directory(STATIC_DIR, filename)

@app.route("/", methods=["GET","POST"])
def index():
    visit_count = increase_visit()
    message = None
    result = None

    if request.method == "POST":
        selected_grade = request.form.get("grade")
        selected_class = request.form.get("class")
        selected_subject = request.form.get("subject")
        uploaded = request.files.get("file")

        if not (selected_grade and selected_class and selected_subject):
            message = "Vui lòng chọn Khối, Lớp, Môn."
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
                filename = secure_filename(uploaded.filename)
                if not allowed_file(filename):
                    message = "Định dạng tệp không được hỗ trợ. Chọn .docx, .pptx, .sb3, .zip"
                else:
                    tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                    tmp_path = os.path.join(tmp_dir, filename)
                    uploaded.save(tmp_path)

                    # load criteria with explicit folder
                    criteria = load_criteria(selected_subject.lower(), int(selected_grade), CRITERIA_DIR)
                    if not criteria:
                        message = f"Chưa có tiêu chí cho {selected_subject} khối {selected_grade}."
                        shutil.rmtree(tmp_dir, ignore_errors=True)
                    else:
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

                        if total is not None:
                            # save to excel
                            append_result(selected_class, selected_grade, pretty_name_from_filename(filename),
                                          selected_subject, filename, total)
                            result = {
                                "student": pretty_name_from_filename(filename),
                                "class": selected_class,
                                "grade": selected_grade,
                                "subject": selected_subject,
                                "file": filename,
                                "total": total,
                                # keep notes internal but we won't show "nhận xét" in main display if you don't want
                                "notes": notes
                            }
                        else:
                            message = message or "Lỗi khi chấm bài."
                        shutil.rmtree(tmp_dir, ignore_errors=True)

    # pass avail_by_grade with string keys for template JS safety
    avail_by_grade = {str(k): v for k, v in AVAILABLE_BY_GRADE.items()}
    return render_template("index.html",
                           classes=CLASSES,
                           avail_by_grade=avail_by_grade,
                           visit_count=read_visit(),
                           result=result,
                           message=message)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)


# === Added by ChatGPT ===
import openpyxl
from openpyxl import Workbook
from datetime import datetime

def save_to_excel(class_name, student_name, results):
    file_path = "results/ket_qua_tong_hop.xlsx"
    import os
    if not os.path.exists("results"):
        os.makedirs("results")
    if not os.path.exists(file_path):
        wb = Workbook()
        wb.remove(wb.active)
        wb.save(file_path)
    wb = openpyxl.load_workbook(file_path)
    if class_name not in wb.sheetnames:
        sheet = wb.create_sheet(title=class_name)
        sheet.append([
            "Thời gian","Họ tên","Lớp",
            "Tổng điểm","Điểm A","Điểm B","Điểm C","Điểm D",
            "Lỗi A","Lỗi B","Lỗi C","Lỗi D"
        ])
    else:
        sheet = wb[class_name]

    sheet.append([
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        student_name,
        class_name,
        results.get("total",0),
        results["A"]["score"],
        results["B"]["score"],
        results["C"]["score"],
        results["D"]["score"],
        results["A"]["errors"],
        results["B"]["errors"],
        results["C"]["errors"],
        results["D"]["errors"],
    ])
    wb.save(file_path)

def count_view():
    import os
    if not os.path.exists("stats"):
        os.makedirs("stats")
    path="stats/views.txt"
    if not os.path.exists(path):
        with open(path,"w") as f: f.write("0")
    with open(path) as f: v=int(f.read().strip())
    v+=1
    with open(path,"w") as f: f.write(str(v))
    return v
# === End Added ===

