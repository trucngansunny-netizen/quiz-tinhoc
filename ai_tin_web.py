# ai_tin_web.py
# Flask app — dùng Gunicorn (Procfile: web: gunicorn ai_tin_web:app)

import os
import shutil
import datetime
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# import hàm chấm từ cham_tieuchi.py
from cham_tieuchi import (
    pretty_name_from_filename,
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch,
)

# ------------- cấu hình -------------
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")
VISIT_FILE = os.path.join(RESULTS_DIR, "visits.txt")

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

# tạo thư mục nếu chưa có
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = "ai_tin_secret_key_replace_if_needed"


# ------------- helpers -------------
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT


def ensure_tonghop():
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = CLASSES[0]
        headers = ["Họ tên", "Lớp", "Khối", "Môn học", "Tên tệp", "Điểm tổng", "Ngày chấm", "Nhận xét", "Tiêu chí chi tiết"]
        ws.append(headers)
        for cls in CLASSES[1:]:
            w = wb.create_sheet(title=cls)
            w.append(headers)
        wb.save(TONGHOP_FILE)


def append_result(class_name, grade, student_name, subject, filename, total, remark, crit_summary):
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student_name, class_name, grade, subject, filename, float(total), now, remark, crit_summary])
    wb.save(TONGHOP_FILE)


def increase_visit():
    try:
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
    except Exception:
        return 0


def compute_stats():
    """Trích dữ liệu từ TONGHOP_FILE -> trả stats cơ bản."""
    if not os.path.exists(TONGHOP_FILE):
        return {"total_sub": 0, "avg": 0.0, "num_ge5": 0}
    try:
        wb = load_workbook(TONGHOP_FILE, data_only=True)
        totals = []
        for sh in wb.sheetnames:
            ws = wb[sh]
            for row in ws.iter_rows(min_row=2, values_only=True):
                try:
                    val = row[5]  # Điểm tổng cột thứ 6
                    if val is not None:
                        totals.append(float(val))
                except Exception:
                    continue
        cnt = len(totals)
        avg = round(sum(totals)/cnt, 2) if cnt > 0 else 0.0
        ge5 = sum(1 for v in totals if v >= 5.0)
        return {"total_sub": cnt, "avg": avg, "num_ge5": ge5}
    except Exception:
        return {"total_sub": 0, "avg": 0.0, "num_ge5": 0}


# ------------- routes -------------
@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(STATIC_DIR, filename)


@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    result = None
    # tăng lượt truy cập (ghi vào file)
    _ = increase_visit()
    visit = read_visit()
    stats = compute_stats()

    if request.method == "POST":
        grade = request.form.get("grade", "").strip()
        cls = request.form.get("class", "").strip()
        subject = request.form.get("subject", "").strip()  # Word / PowerPoint / Scratch
        uploaded = request.files.get("file")

        if not grade or not cls or not subject:
            message = "Vui lòng chọn Khối, Lớp và Môn."
        else:
            try:
                gnum = int(grade)
            except:
                gnum = None
                message = "Khối không hợp lệ."

            if gnum and subject not in AVAILABLE_BY_GRADE.get(gnum, []):
                message = f"Khối {gnum} không học môn {subject}."
            else:
                if not uploaded or uploaded.filename == "":
                    message = "Vui lòng chọn tệp để nộp."
                else:
                    fname = secure_filename(uploaded.filename)
                    if not allowed_file(fname):
                        message = "Định dạng tệp không được hỗ trợ (hỗ trợ: .docx .pptx .sb3 .zip)."
                    else:
                        tmp_dir = None
                        try:
                            tmp_dir = os.path.join(RESULTS_DIR, "tmp")
                            os.makedirs(tmp_dir, exist_ok=True)
                            tmp_path = os.path.join(tmp_dir, fname)
                            uploaded.save(tmp_path)

                            # load criteria
                            crit = load_criteria(subject.lower(), int(grade), CRITERIA_DIR)
                            if not crit:
                                message = f"Chưa có tiêu chí cho {subject} khối {grade}."
                                shutil.rmtree(tmp_dir, ignore_errors=True)
                            else:
                                total = None
                                notes = []
                                # gọi hàm chấm tương ứng
                                if subject == "Word":
                                    total, notes = grade_word(tmp_path, crit)
                                elif subject == "PowerPoint":
                                    total, notes = grade_ppt(tmp_path, crit)
                                elif subject == "Scratch":
                                    total, notes = grade_scratch(tmp_path, crit)

                                if total is None:
                                    message = "Lỗi khi chấm bài (không đọc được nội dung)."
                                else:
                                    try:
                                        tval = float(total)
                                    except:
                                        tval = 0.0

                                    if tval >= 9.5:
                                        remark = "Hoàn thành xuất sắc"
                                    elif tval >= 8.0:
                                        remark = "Hoàn thành tốt"
                                    elif tval >= 6.5:
                                        remark = "Đạt yêu cầu"
                                    else:
                                        remark = "Cần cố gắng thêm"

                                    crit_summary = "; ".join(notes)
                                    append_result(cls, grade, pretty_name_from_filename(fname),
                                                  subject, fname, tval, remark, crit_summary)

                                    result = {
                                        "student": pretty_name_from_filename(fname),
                                        "class": cls,
                                        "grade": grade,
                                        "subject": subject,
                                        "file": fname,
                                        "total": tval,
                                        "remark": remark,
                                        "notes": notes
                                    }
                                    stats = compute_stats()
                                    message = None
                        except Exception as e:
                            message = f"Lỗi xử lý tệp: {e}"
                        finally:
                            try:
                                if tmp_dir and os.path.exists(tmp_dir):
                                    shutil.rmtree(tmp_dir, ignore_errors=True)
                            except Exception:
                                pass

    return render_template("index.html",
                           classes=CLASSES,
                           avail_by_grade=AVAILABLE_BY_GRADE,
                           visit_count=read_visit(),
                           result=result,
                           message=message,
                           stats=stats)


# ------------- run (local/debug) -------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
