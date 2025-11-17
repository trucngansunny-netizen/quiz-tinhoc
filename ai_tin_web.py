# ai_tin_web.py
import os
import json
import shutil
import tempfile
import datetime
from flask import Flask, render_template, request, send_from_directory, send_file
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

# Import module chấm tiêu chí (giữ nguyên file bạn có: cham_tieuchi.py)
from cham_tieuchi import (
    pretty_name_from_filename,
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch
)

# === CẤU HÌNH ===
APP_ROOT = os.path.dirname(__file__)
STATIC_DIR = os.path.join(APP_ROOT, "static")
CRITERIA_DIR = os.path.join(APP_ROOT, "criteria")
RESULTS_DIR = os.path.join(APP_ROOT, "results")
TONGHOP_FILE = os.path.join(RESULTS_DIR, "tonghop.xlsx")
DETAILS_FILE = os.path.join(RESULTS_DIR, "details.xlsx")
VISIT_FILE = os.path.join(RESULTS_DIR, "visits.txt")

# Mật khẩu giáo viên (bạn có thể đổi nếu muốn)
TEACHER_PASSWORD = "giaovien123"

ALLOWED_EXT = {"pptx", "docx", "sb3", "zip"}

# Danh sách 14 lớp (đảm bảo đúng 14 lớp)
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

# Tạo thư mục cần thiết nếu chưa có
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CRITERIA_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET", "change_this_secret_in_prod")

# === LƯỢT TRUY CẬP ===
def read_visit():
    try:
        if not os.path.exists(VISIT_FILE):
            return 0
        return int(open(VISIT_FILE, "r", encoding="utf-8").read().strip() or "0")
    except Exception:
        return 0

def increase_visit():
    # tăng +1 và ghi lại (khi GET trang)
    c = read_visit() + 1
    with open(VISIT_FILE, "w", encoding="utf-8") as f:
        f.write(str(c))
    return c

# === TỔNG HỢP EXCEL (1 sheet / 1 lớp) ===
def ensure_tonghop():
    if not os.path.exists(TONGHOP_FILE):
        wb = Workbook()
        # loại sheet mặc định "Sheet" nếu có
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        for cls in CLASSES:
            ws = wb.create_sheet(title=cls)
            ws.append(["Họ và tên", "Lớp", "Khối", "Môn học", "Tên tệp", "Điểm tổng", "Ngày chấm"])
        wb.save(TONGHOP_FILE)

def append_result(class_name, grade, student_name, subject, filename, total):
    ensure_tonghop()
    wb = load_workbook(TONGHOP_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ và tên", "Lớp", "Khối", "Môn học", "Tên tệp", "Điểm tổng", "Ngày chấm"])
    ws = wb[class_name]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student_name, class_name, grade, subject, filename, float(total), now])
    wb.save(TONGHOP_FILE)

# === EXCEL CHI TIẾT (ghi JSON notes vào ô) ===
def ensure_details():
    if not os.path.exists(DETAILS_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "ChiTiet"
        ws.append(["Thời gian", "Họ tên", "Lớp", "Khối", "Môn", "Tên tệp", "Tổng điểm", "Chi tiết (JSON)"])
        wb.save(DETAILS_FILE)

def save_detail_excel(student_name, class_name, grade, subject, filename, total, notes):
    ensure_details()
    wb = load_workbook(DETAILS_FILE)
    ws = wb["ChiTiet"]
    ws.append([
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        student_name,
        class_name,
        grade,
        subject,
        filename,
        float(total),
        json.dumps(notes, ensure_ascii=False)
    ])
    wb.save(DETAILS_FILE)

# === HỖ TRỢ KIỂM TRA FILE ===
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

# Static route (nếu cần)
@app.route("/static/<path:filename>")
def custom_static(filename):
    return send_from_directory(STATIC_DIR, filename)

# Helper: map tên môn hiển thị -> prefix file tiêu chí mà cham_tieuchi tìm (ppt/word/scratch)
SUBJECT_TO_PREFIX = {
    "Word": "word",
    "PowerPoint": "ppt",
    "Scratch": "scratch"
}

# === TRANG CHÍNH (Nộp & Chấm) ===
@app.route("/", methods=["GET", "POST"])
def index():
    # Tăng lượt truy cập chỉ khi GET
    if request.method == "GET":
        increase_visit()

    message = None
    result = None

    if request.method == "POST":
        selected_grade = request.form.get("grade")
        selected_class = request.form.get("class")
        selected_subject = request.form.get("subject")
        uploaded = request.files.get("file")

        if not (selected_grade and selected_class and selected_subject):
            message = "Vui lòng chọn Khối, Lớp và Môn."
        elif not uploaded or uploaded.filename == "":
            message = "Vui lòng chọn tệp để nộp."
        else:
            # kiểm tra khối hợp lệ
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
                    message = "Chỉ hỗ trợ .docx, .pptx, .sb3, .zip"
                else:
                    tmp_dir = tempfile.mkdtemp(prefix="ai_tin_")
                    tmp_path = os.path.join(tmp_dir, filename)
                    uploaded.save(tmp_path)

                    # Map subject -> prefix (word/ppt/scratch) để cham_tieuchi.load_criteria tìm file đúng
                    subj_prefix = SUBJECT_TO_PREFIX.get(selected_subject, selected_subject.lower())

                    # LOAD TIÊU CHÍ
                    criteria = load_criteria(subj_prefix, int(selected_grade), CRITERIA_DIR)
                    if not criteria:
                        message = f"Chưa có tiêu chí cho {selected_subject} khối {selected_grade}."
                        shutil.rmtree(tmp_dir, ignore_errors=True)
                    else:
                        notes = []
                        total = None
                        try:
                            # gọi hàm chấm tương ứng - các hàm trong cham_tieuchi.py (do bạn cung cấp) trả về (total, notes)
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

                        # Nếu grader của bạn chỉ trả về tổng (và notes rỗng), xây notes mặc định từ tiêu chí
                        if total is not None and (not notes):
                            # dùng tiêu chí để tạo danh sách ✅/❌ tạm thời dựa trên tổng (không chuẩn xác)
                            # Tuy nhiên thường cham_tieuchi.py của cô đã trả notes chi tiết — nếu không, ta hiển thị chỉ tiêu.
                            notes = []
                            for it in criteria.get("tieu_chi", []):
                                mo_ta = it.get("mo_ta", "")
                                diem = it.get("diem", 0)
                                notes.append(f"- {mo_ta} ({diem}đ)")

                        if total is not None:
                            student_name = pretty_name_from_filename(filename)

                            # LƯU TỔNG HỢP CHO GIÁO VIÊN
                            append_result(
                                selected_class,
                                selected_grade,
                                student_name,
                                selected_subject,
                                filename,
                                total
                            )

                            # LƯU CHI TIẾT (ghi notes JSON)
                            save_detail_excel(
                                student_name,
                                selected_class,
                                selected_grade,
                                selected_subject,
                                filename,
                                total,
                                notes
                            )

                            # TRẢ KẾT QUẢ CHO HỌC SINH
                            # result keys phải khớp template index.html: result.student, result.class, result.subject, result.total, result.remark, result.notes
                            result = {
                                "student": student_name,
                                "class": selected_class,
                                "grade": selected_grade,
                                "subject": selected_subject,
                                "file": filename,
                                "total": total,
                                "remark": "Đã nộp thành công",
                                "notes": notes
                            }
                        else:
                            message = message or "Lỗi khi chấm bài."
                        shutil.rmtree(tmp_dir, ignore_errors=True)

    # pass avail_by_grade with string keys for template JS safety
    avail_by_grade = {str(k): v for k, v in AVAILABLE_BY_GRADE.items()}

    return render_template(
        "index.html",
        classes=CLASSES,
        avail_by_grade=avail_by_grade,
        visit_count=read_visit(),
        result=result,
        message=message
    )

# === ROUTE BẢO MẬT CHO GIÁO VIÊN TẢI FILE TỔNG HỢP ===
@app.route("/download-tonghop", methods=["GET", "POST"])
def download_tonghop():
    if request.method == "GET":
        return """
        <html><body>
        <h3>Download file tổng hợp (giáo viên)</h3>
        <form method="post">
          Mật khẩu: <input name="pw" type="password"/>
          <button type="submit">Tải về</button>
        </form>
        <p><a href="/">Quay về trang nộp bài</a></p>
        </body></html>
        """

    pw = request.form.get("pw", "")
    if pw != TEACHER_PASSWORD:
        return "<h3>Mật khẩu sai.</h3><p><a href='/download-tonghop'>Thử lại</a></p>", 401

    if not os.path.exists(TONGHOP_FILE):
        return "<h3>Chưa có file tổng hợp nào.</h3><p><a href='/'>Quay về</a></p>", 404

    return send_file(TONGHOP_FILE, as_attachment=True, download_name=os.path.basename(TONGHOP_FILE))

# === Route để xem file chi tiết (giáo viên) nếu cần ===
@app.route("/download-details", methods=["GET", "POST"])
def download_details():
    if request.method == "GET":
        return """
        <html><body>
        <h3>Download file chi tiết (giáo viên)</h3>
        <form method="post">
          Mật khẩu: <input name="pw" type="password"/>
          <button type="submit">Tải về</button>
        </form>
        <p><a href="/">Quay về trang nộp bài</a></p>
        </body></html>
        """

    pw = request.form.get("pw", "")
    if pw != TEACHER_PASSWORD:
        return "<h3>Mật khẩu sai.</h3><p><a href='/download-details'>Thử lại</a></p>", 401

    if not os.path.exists(DETAILS_FILE):
        return "<h3>Chưa có file chi tiết nào.</h3><p><a href='/'>Quay về</a></p>", 404

    return send_file(DETAILS_FILE, as_attachment=True, download_name=os.path.basename(DETAILS_FILE))

# === RUN ===
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
