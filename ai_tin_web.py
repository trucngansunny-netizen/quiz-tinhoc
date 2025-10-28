# ai_tin_web.py
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os, json, datetime, tempfile, shutil
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# dùng các hàm chấm của bạn (giữ nguyên file cham_tieuchi.py)
from cham_tieuchi import (
    pretty_name_from_filename,
    # load_criteria,  # không gọi hàm load_criteria trong cham_tieuchi để tránh mismatch tên file
    grade_word,
    grade_ppt,
    grade_scratch
)

# ---------- Cấu hình ----------
APP_ROOT = os.path.dirname(__file__)
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
STATIC_DIR = os.path.join(APP_ROOT, "static")
COUNTER_FILE = os.path.join(APP_ROOT, "counter.txt")
ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

# tạo thư mục nếu chưa có
os.makedirs(RESULTS_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# danh sách lớp mặc định (theo yêu cầu bạn)
CLASSES = [
    "3A1","3A2","3A3","3A4",   # 4 lớp khối 3
    "4A1","4A2","4A3","4A4","4A5",  # 5 lớp khối 4
    "5A1","5A2","5A3","5A4","5A5"   # 5 lớp khối 5
]

# môn hợp lệ theo khối
AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word", "PowerPoint", "Scratch"],
    5: ["Word", "Scratch"]
}

# map software key -> filename prefix used in criteria files
PREFIX = {"Word": "word", "PowerPoint": "ppt", "Scratch": "scratch"}

app = Flask(__name__, static_folder="static", template_folder="templates")


# ---------- Counter lượt truy cập ----------
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


# ---------- Tải tiêu chí từ folder criteria ----------
def load_criteria_local(software, grade):
    """
    software: 'Word' / 'PowerPoint' / 'Scratch'
    grade: int (3/4/5)
    Try filename variants:
      - prefix + grade + .json (e.g. ppt3.json)
      - prefix + '_khoi' + grade + .json (e.g. ppt_khoi3.json)
    Returns dict or None
    """
    pref = PREFIX.get(software)
    if not pref:
        return None
    candidates = [
        f"{pref}{grade}.json",
        f"{pref}_khoi{grade}.json",
        f"{pref}_khoi_{grade}.json",
        f"{pref}{grade}.JSON"
    ]
    for fn in candidates:
        path = os.path.join(CRITERIA_DIR, fn)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # expect data to have "tieu_chi"
                if isinstance(data, dict) and "tieu_chi" in data:
                    return data
                else:
                    # if file exists but structure different, still return dict wrapper
                    return {"tieu_chi": data.get("tieu_chi") if isinstance(data, dict) else []}
            except Exception as e:
                return None
    return None


# ---------- Ghi kết quả vào file Excel theo lớp ----------
def ensure_results_workbook(class_name):
    """
    Ensure that results/ketqua_<class_name>.xlsx exists with header.
    Return path.
    """
    os.makedirs(RESULTS_DIR, exist_ok=True)
    fname = f"ketqua_{class_name}.xlsx"
    path = os.path.join(RESULTS_DIR, fname)
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "KẾT QUẢ"
        ws.append(["Ngày giờ", "Tên học sinh", "Lớp", "Phần mềm", "Điểm", "Nhận xét", "Tên tệp gốc"])
        wb.save(path)
    return path

def append_result(class_name, student_name, software, score, notes, original_filename):
    path = ensure_results_workbook(class_name)
    wb = load_workbook(path)
    ws = wb.active
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([now, student_name, class_name, software, score, "; ".join(notes), original_filename])
    wb.save(path)


# ---------- Helper: validate extension ----------
def allowed_file(filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    return ext in ALLOWED_EXT

# ---------- Route: tệp static (nếu cần) ----------
@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(STATIC_DIR, filename)

# ---------- Trang chính ----------
@app.route("/", methods=["GET", "POST"])
def home():
    # tăng counter mỗi lần GET (lượt truy cập)
    if request.method == "GET":
        counter = increase_counter()
    else:
        counter = get_counter()

    result = None
    result_notes = []
    result_score = None
    selected_class = None
    selected_grade = None
    selected_software = None
    criteria = None
    message = None

    if request.method == "POST":
        # form data: lop, software, file
        selected_class = request.form.get("lop")
        selected_software = request.form.get("software")
        if not selected_class:
            message = "Vui lòng chọn lớp (ví dụ 4A1)."
        else:
            selected_grade = int(selected_class[0])

        # check software availability for grade
        if selected_grade and selected_software:
            avail = AVAILABLE_BY_GRADE.get(selected_grade, [])
            if selected_software not in avail:
                message = f"Khối {selected_grade} không học phần mềm {selected_software}."
            else:
                # load criteria
                criteria = load_criteria_local(selected_software, selected_grade)
                if criteria is None:
                    message = f"Chưa có tiêu chí cho {selected_software} khối {selected_grade}."
        # handle file upload
        uploaded = request.files.get("file")
        if uploaded and uploaded.filename:
            filename = secure_filename(uploaded.filename)
            if not allowed_file(filename):
                message = f"Định dạng file không hợp lệ. Hỗ trợ: .pptx .docx .sb3 .zip"
            else:
                # save to temp file
                tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                tmp_path = os.path.join(tmp_dir, filename)
                uploaded.save(tmp_path)

                # if criteria still None try load (in case teacher didn't select software before upload)
                if criteria is None and selected_grade and selected_software:
                    criteria = load_criteria_local(selected_software, selected_grade)

                # perform grading by calling your grading functions
                score = None
                notes = []
                try:
                    if selected_software == "Word":
                        score, notes = grade_word(tmp_path, criteria if criteria else {"tieu_chi":[]})
                    elif selected_software == "PowerPoint":
                        score, notes = grade_ppt(tmp_path, criteria if criteria else {"tieu_chi":[]})
                    elif selected_software == "Scratch":
                        score, notes = grade_scratch(tmp_path, criteria if criteria else {"tieu_chi":[]})
                    else:
                        message = "Phần mềm không hợp lệ."
                except Exception as e:
                    message = f"Lỗi khi chấm: {e}"
                    score = None
                    notes = [str(e)]

                # extract student name from filename using your function
                student_name = pretty_name_from_filename(filename)

                # append result if score computed
                if score is not None and selected_class:
                    append_result(selected_class, student_name, selected_software, score, notes, filename)
                    result = {
                        "student": student_name,
                        "score": score,
                        "notes": notes,
                        "file": filename
                    }
                    result_score = score
                    result_notes = notes

                # cleanup temp
                try:
                    shutil.rmtree(tmp_dir)
                except:
                    pass

        else:
            # no file uploaded: maybe teacher only requested to view criteria
            if not uploaded and criteria and not message:
                message = "Đã tải tiêu chí. Vui lòng tải file học sinh để chấm."

    # background to use in template (static path)
    bg = "bg_default.jpg"
    if selected_grade:
        if selected_grade == 3:
            bg = "bg_3.jpg"
        elif selected_grade == 4:
            bg = "bg_4.jpg"
        elif selected_grade == 5:
            bg = "bg_5.jpg"

    # list of softwares available to select box (we show all but front-end can hide invalid ones)
    softwares = ["PowerPoint", "Word", "Scratch"]

    # prepare criteria list for rendering (if loaded)
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

if __name__ == "__main__":
    # note: run via gunicorn ai_tin_web:app in production
    app.run(host="0.0.0.0", port=5000)
