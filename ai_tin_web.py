# ai_tin_web.py
import os
import json
import tempfile
import shutil
import datetime
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# import functions chấm bạn đã viết
from cham_tieuchi import (
    pretty_name_from_filename,
    grade_word,
    grade_ppt,
    grade_scratch
)

# --------- Cấu hình ---------
APP_ROOT = os.path.dirname(__file__)
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
STATIC_DIR = os.path.join(APP_ROOT, "static")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
COUNTER_FILE = os.path.join(APP_ROOT, "counter.txt")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")

ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

# Danh sách 14 lớp (theo yêu cầu bạn)
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

# map hiển thị -> tiền tố file criteria
PREFIX = {"Word": "word", "PowerPoint": "ppt", "Scratch": "scratch"}

# Tạo folder cần thiết
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")


# --------- Counter lượt truy cập ----------
def get_counter():
    try:
        if not os.path.exists(COUNTER_FILE):
            with open(COUNTER_FILE, "w", encoding="utf-8") as f:
                f.write("0")
            return 0
        with open(COUNTER_FILE, "r", encoding="utf-8") as f:
            return int(f.read().strip() or "0")
    except:
        return 0

def increase_counter():
    c = get_counter() + 1
    with open(COUNTER_FILE, "w", encoding="utf-8") as f:
        f.write(str(c))
    return c


# --------- Load tiêu chí (hỗ trợ nhiều biến thể tên file) ----------
def load_criteria_local(software, grade):
    """
    software: 'Word' / 'PowerPoint' / 'Scratch'
    grade: int 3/4/5
    tìm file: ppt3.json, ppt_khoi3.json, ppt3_khoi.json, ...
    """
    pref = PREFIX.get(software)
    if not pref:
        return None
    candidates = [
        f"{pref}{grade}.json",
        f"{pref}_khoi{grade}.json",
        f"{pref}_khoi_{grade}.json",
        f"{pref}{grade}.JSON",
        f"{pref}_khoi{grade}.JSON"
    ]
    for fn in candidates:
        path = os.path.join(CRITERIA_DIR, fn)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and "tieu_chi" in data:
                    return data
                # nếu file chưa đúng cấu trúc, cố lấy danh sách tieu_chi nếu có
                return {"tieu_chi": data.get("tieu_chi") if isinstance(data, dict) else []}
            except Exception:
                return None
    return None


# --------- Tạo file Excel tonghop với 14 sheet (mỗi sheet 1 lớp) ----------
def ensure_tonghop_excel():
    if not os.path.exists(RESULTS_DIR):
        os.makedirs(RESULTS_DIR, exist_ok=True)
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        # tạo sheet cho từng lớp
        for i, cls in enumerate(CLASSES):
            if i == 0 and "Sheet" in wb.sheetnames:
                ws = wb["Sheet"]
                ws.title = cls
            else:
                wb.create_sheet(title=cls)
            ws = wb[cls]
            # header
            ws.append(["Thời gian", "Họ tên học sinh", "Khối", "Lớp", "Phần mềm", "Điểm", "Nhận xét", "Tên tệp gốc"])
        # nếu còn sheet mặc định khác thì xóa (an toàn)
        for s in list(wb.sheetnames):
            if s not in CLASSES:
                try:
                    del wb[s]
                except:
                    pass
        wb.save(TONGHOP_FILE)
    return TONGHOP_FILE

def append_to_tonghop(class_name, student_name, grade, software, score, notes, original_filename):
    ensure_tonghop_excel()
    wb = load_workbook(TONGHOP_FILE)
    # nếu sheet chưa tồn tại (không khả thi), tạo
    if class_name not in wb.sheetnames:
        wb.create_sheet(title=class_name)
        ws = wb[class_name]
        ws.append(["Thời gian", "Họ tên học sinh", "Khối", "Lớp", "Phần mềm", "Điểm", "Nhận xét", "Tên tệp gốc"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([now, student_name, grade, class_name, software, score, "; ".join(notes), original_filename])
    wb.save(TONGHOP_FILE)


# --------- Utilities ----------
def allowed_file(filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    return ext in ALLOWED_EXT

@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(STATIC_DIR, filename)


# --------- Trang chính (GET tăng counter) ----------
@app.route("/", methods=["GET", "POST"])
def home():
    # tăng lượt truy cập khi GET
    if request.method == "GET":
        counter = increase_counter()
    else:
        counter = get_counter()

    message = None
    criteria = None
    result = None

    selected_class = request.form.get("lop") if request.method == "POST" else None
    selected_software = request.form.get("software") if request.method == "POST" else None
    selected_grade = int(selected_class[0]) if selected_class else None

    # nếu POST và đã chọn lớp + phần mềm, load tiêu chí
    if request.method == "POST" and selected_class and selected_software:
        # kiểm tra môn có thuộc khối không
        if selected_grade and selected_software not in AVAILABLE_BY_GRADE.get(selected_grade, []):
            message = f"Khối {selected_grade} không học phần mềm {selected_software}."
        else:
            criteria = load_criteria_local(selected_software, selected_grade)
            if not criteria:
                message = f"Chưa có tiêu chí cho {selected_software} khối {selected_grade}."

        # xử lý file nếu upload
        uploaded = request.files.get("file")
        if uploaded and uploaded.filename:
            filename = secure_filename(uploaded.filename)
            if not allowed_file(filename):
                message = "Định dạng file không hợp lệ (hỗ trợ: .pptx .docx .sb3 .zip)."
            else:
                tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                tmp_path = os.path.join(tmp_dir, filename)
                uploaded.save(tmp_path)
                # gọi hàm chấm
                score = None
                notes = []
                try:
                    if selected_software == "Word":
                        score, notes = grade_word(tmp_path, criteria if criteria else {"tieu_chi":[]})
                    elif selected_software == "PowerPoint":
                        score, notes = grade_ppt(tmp_path, criteria if criteria else {"tieu_chi":[]})
                    elif selected_software == "Scratch":
                        score, notes = grade_scratch(tmp_path, criteria if criteria else {"tieu_chi":[]})

                except Exception as e:
                    message = f"Lỗi khi chấm: {e}"
                    notes = [str(e)]
                    score = None

                student_name = pretty_name_from_filename(filename)

                # lưu vào tonghop.xlsx nếu có điểm
                if score is not None and selected_class:
                    append_to_tonghop(selected_class, student_name, selected_grade, selected_software, score, notes, filename)
                    result = {
                        "student": student_name,
                        "score": score,
                        "notes": notes,
                        "file": filename
                    }

                # cleanup
                try:
                    shutil.rmtree(tmp_dir)
                except:
                    pass

    # background chọn theo grade để template hiển thị
    bg = "bg_default.jpg"
    if selected_grade:
        if selected_grade == 3:
            bg = "bg_3.jpg"
        elif selected_grade == 4:
            bg = "bg_4.jpg"
        elif selected_grade == 5:
            bg = "bg_5.jpg"

    softwares = ["PowerPoint", "Word", "Scratch"]
    criteria_list = criteria.get("tieu_chi", []) if criteria else None

    return render_template(
        "index.html",
        classes=CLASSES,
        softwares=softwares,
        selected_class=selected_class,
        selected_software=selected_software,
        criteria=criteria_list,
        result=result,
        message=message,
        background=bg,
        counter=counter
    )

# --------- chạy trực tiếp cho debug local (không dùng khi deploy với gunicorn) ----------
if __name__ == "__main__":
    # chạy local thử nghiệm
    app.run(host="0.0.0.0", port=5000)
