import os, json, tempfile, shutil, datetime
from flask import Flask, render_template, request, send_from_directory, flash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook
from cham_tieuchi import pretty_name_from_filename, load_criteria, grade_word, grade_ppt, grade_scratch

# ---------------- CONFIG ----------------
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
VISIT_FILE = os.path.join(RESULTS_DIR, "visits.txt")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")

ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

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
app.secret_key = "ai_tin_secret_key"

# ---------------- UTILS ----------------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

def increase_visit():
    try:
        if not os.path.exists(VISIT_FILE):
            with open(VISIT_FILE, "w", encoding="utf-8") as f:
                f.write("0")
        with open(VISIT_FILE, "r+", encoding="utf-8") as f:
            count = int(f.read().strip() or "0") + 1
            f.seek(0)
            f.write(str(count))
            f.truncate()
        return count
    except:
        return 0

def read_visit():
    try:
        with open(VISIT_FILE, "r", encoding="utf-8") as f:
            return int(f.read().strip() or "0")
    except:
        return 0

def ensure_tonghop():
    """Tạo file tổng hợp kết quả nếu chưa có"""
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = CLASSES[0]
        ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        for cls in CLASSES[1:]:
            ws2 = wb.create_sheet(title=cls)
            ws2.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
        wb.save(TONGHOP_FILE)

def append_result(class_name, grade, student_name, subject, filename, total, remark, detail):
    """Ghi kết quả vào Excel tổng hợp"""
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm","Nhận xét","Tiêu chí chi tiết"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student_name, class_name, grade, subject, filename, total, now, remark, detail])
    wb.save(TONGHOP_FILE)

# ---------------- ROUTES ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    visit_count = increase_visit() or read_visit()
    message = result = None

    if request.method == "POST":
        grade = request.form.get("grade")
        cls = request.form.get("class")
        subject = request.form.get("subject")
        uploaded = request.files.get("file")

        if not all([grade, cls, subject, uploaded]):
            message = "⚠️ Vui lòng chọn đầy đủ thông tin và tệp."
        elif not allowed_file(uploaded.filename):
            message = "⚠️ Định dạng không hợp lệ (.docx, .pptx, .sb3, .zip)."
        else:
            filename = secure_filename(uploaded.filename)
            tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
            tmp_path = os.path.join(tmp_dir, filename)
            uploaded.save(tmp_path)

            # Nạp tiêu chí
            try:
                criteria = load_criteria(subject.lower(), int(grade), CRITERIA_DIR)
            except:
                criteria = None

            if not criteria:
                message = f"❌ Chưa có tiêu chí cho {subject} khối {grade}."
                shutil.rmtree(tmp_dir, ignore_errors=True)
            else:
                try:
                    if subject == "Word":
                        total, notes = grade_word(tmp_path, criteria)
                    elif subject == "PowerPoint":
                        total, notes = grade_ppt(tmp_path, criteria)
                    elif subject == "Scratch":
                        total, notes = grade_scratch(tmp_path, criteria)
                    else:
                        total, notes = None, ["❌ Môn không hợp lệ."]
                except Exception as e:
                    total, notes = None, [f"Lỗi khi chấm: {e}"]

                if total is None:
                    message = "❌ Không thể chấm bài (file bị lỗi hoặc tiêu chí sai)."
                else:
                    detail = "; ".join(notes)
                    # Xếp loại
                    if total >= 9.5:
                        remark = "Hoàn thành xuất sắc"
                    elif total >= 8:
                        remark = "Hoàn thành tốt"
                    elif total >= 6.5:
                        remark = "Đạt yêu cầu"
                    else:
                        remark = "Cần cố gắng thêm"

                    append_result(cls, grade, pretty_name_from_filename(filename), subject, filename, total, remark, detail)
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

            shutil.rmtree(tmp_dir, ignore_errors=True)

    return render_template("index.html",
                           classes=CLASSES,
                           avail_by_grade=AVAILABLE_BY_GRADE,
                           visit_count=read_visit(),
                           result=result,
                           message=message)

# ---------------- STATIC FILE ----------------
@app.route("/static/<path:filename>")
def custom_static(filename):
    return send_from_directory(STATIC_DIR, filename)

# ---------------- MAIN ----------------
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

