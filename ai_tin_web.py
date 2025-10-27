# ai_tin_web.py
import streamlit as st
import os
import tempfile
from pathlib import Path
from openpyxl import load_workbook

# import xử lý/chấm từ core.py (bạn đã có)
from core import (
    load_criteria,
    grade_word,
    grade_ppt,
    grade_scratch,
    pretty_name_from_filename,
    ensure_workbook_exists,
)

# ========== Cấu hình ========== #
EXCEL_FILE = "ketqua_tonghop.xlsx"
CRITERIA_FOLDER = "criteria"

# Danh sách lớp (bạn có thể điều chỉnh)
CLASSES = ["3A1","3A2","3A3","3A4","4A1","4A2","4A3","4A4","4A5","5A1","5A2","5A3","5A4","5A5"]

# Ảnh nền & logo (file nằm cùng folder với ai_tin_web.py)
BASE_DIR = os.path.dirname(__file__)
BG_FILES = {
    "default": os.path.join(BASE_DIR, "bg_default.jpg"),
    3: os.path.join(BASE_DIR, "bg_3.jpg"),
    4: os.path.join(BASE_DIR, "bg_4.jpg"),
    5: os.path.join(BASE_DIR, "bg_5.jpg"),
}
LOGO_PATH = os.path.join(BASE_DIR, "logo_tranquoctoan.png")

st.set_page_config(page_title="AI-TIN Web", page_icon="🧠", layout="centered")

# ========== Helper ==========
def safe_grade_from_class(class_name):
    # Lấy số đầu tiên làm khối (an toàn hơn)
    digits = "".join(ch for ch in class_name if ch.isdigit())
    return int(digits[0]) if digits else 0

def save_uploaded(tmpfile):
    # lưu file upload tạm (trả về đường dẫn)
    suffix = os.path.splitext(tmpfile.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(tmpfile.getbuffer())
    tmp.close()
    return tmp.name

def append_to_sheet(class_name, student_name, subject, score, notes):
    ensure_workbook_exists(EXCEL_FILE)
    wb = load_workbook(EXCEL_FILE)
    # tạo sheet theo lớp nếu chưa có
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ tên học sinh","Môn","Điểm","Nhận xét"])
    else:
        ws = wb[class_name]
    ws.append([student_name, subject, score, "; ".join(notes)])
    wb.save(EXCEL_FILE)

def criteria_file_exists(subj_code, grade):
    # subj_code: 'word' / 'ppt' / 'scratch'
    fname = f"{subj_code}{grade}.json"
    return os.path.exists(os.path.join(BASE_DIR, CRITERIA_FOLDER, fname))

# ========== CSS + Background ==========
def set_background_for_grade(grade):
    bg = BG_FILES.get(grade, BG_FILES["default"])
    # nếu file không tồn tại thì dùng default
    if not os.path.exists(bg):
        bg = BG_FILES["default"]
    # CSS: set background + overlay mờ tối
    css = f"""
    <style>
    .stApp {{
        background: linear-gradient(rgba(0,0,0,0.35), rgba(0,0,0,0.35)), url("file://{bg}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        color: #111;
        font-family: 'Segoe UI', sans-serif;
    }}
    /* logo góc trái */
    .logo-top-left {{
        position: fixed;
        top: 12px;
        left: 14px;
        z-index: 9999;
    }}
    .card {{
        background: rgba(255,255,255,0.92);
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 8px 30px rgba(0,0,0,0.25);
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)
    # hiển thị logo góc trên trái (nếu có)
    if os.path.exists(LOGO_PATH):
        st.markdown(
            f"""<div class="logo-top-left"><img src="file://{LOGO_PATH}" width="64" style="border-radius:50%; opacity:0.95;"></div>""",
            unsafe_allow_html=True
        )

# ========== Giao diện ==========
set_background_for_grade("default")  # mặc định trước khi chọn lớp

st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("<h1 style='margin:0;'>🧠 AI-TIN Web — Trợ lý chấm bài</h1>", unsafe_allow_html=True)
st.markdown("<p style='margin-top:6px; color:#333;'>Chọn lớp và tải file để chấm tự động</p>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# Sidebar: chọn lớp
st.sidebar.header("🎓 Chọn lớp học")
selected_class = st.sidebar.selectbox("Chọn lớp", CLASSES)
grade = safe_grade_from_class(selected_class)

# Thay nền theo khối đã chọn
set_background_for_grade(grade)

# Hiển thị tiêu đề môn theo khối
st.sidebar.subheader(f"📚 Tiêu chí chấm cho khối {grade}")

# Map hiển thị label -> subject code
SUBJ_LABELS = {
    "Word": "word",
    "Bài thuyết trình PowerPoint": "ppt",
    "Lập trình Scratch": "scratch"
}

# Kiểm tra môn nào có file criteria cho khối hiện tại
available = {}
for label, code in SUBJ_LABELS.items():
    exists = criteria_file_exists(code, grade)
    available[label] = exists

# Sidebar selectbox để chọn xem tiêu chí cho môn nào (chỉ hiện labels)
available_labels = [lbl for lbl, ok in available.items() if ok]
if not available_labels:
    # nếu không có tiêu chí nào (ví dụ khối 3 có chỉ ppt), hiển thị thông báo
    st.sidebar.info(f"⚠️ Khối {grade} hiện chưa có tiêu chí hiển thị (hoặc không học các phần mềm).")
    subj_choice_label = st.sidebar.selectbox("Xem tiêu chí cho", list(SUBJ_LABELS.keys()))
else:
    subj_choice_label = st.sidebar.selectbox("Xem tiêu chí cho", available_labels)

# Nếu môn được chọn không có tiêu chí => show message
subj_code = SUBJ_LABELS.get(subj_choice_label, None)
if subj_code is None:
    st.sidebar.error("Lỗi chọn môn.")
else:
    # nếu file tiêu chí tồn tại => load và hiển thị
    crit = None
    if criteria_file_exists(subj_code, grade):
        crit = load_criteria(subj_code, grade, CRITERIA_FOLDER)
        if crit is None:
            st.sidebar.error(f"⚠️ Không thể đọc file tiêu chí cho {subj_choice_label} khối {grade}.")
        else:
            st.sidebar.markdown("### 🔍 Tiêu chí")
            for it in crit.get("tieu_chi", []):
                mo = it.get("mo_ta", "")
                diem = it.get("diem", 0)
                st.sidebar.markdown(f"- {mo} (`{diem}` đ)")
    else:
        # rõ ràng hiển thị nếu khối không học phần mềm này
        not_learn_msg = ""
        if subj_code == "word" and grade == 3:
            not_learn_msg = f"Khối {grade} không học phần mềm Word."
        elif subj_code == "ppt" and grade == 5:
            not_learn_msg = f"Khối {grade} không học phần mềm PowerPoint."
        elif subj_code == "scratch" and grade == 3:
            not_learn_msg = f"Khối {grade} không học phần mềm Scratch."
        else:
            not_learn_msg = f"Không có tiêu chí cho {subj_choice_label} khối {grade}."
        st.sidebar.warning(not_learn_msg)

# ========== Main: Upload & Chấm ==========
st.markdown("<div class='card' style='max-width:900px; margin: 10px auto;'>", unsafe_allow_html=True)
st.markdown(f"**Lớp đang chọn:** {selected_class} (Khối {grade})")

st.write("Chọn loại bài để nộp và chấm (hệ thống sẽ chỉ hiện những loại phù hợp với khối).")

cols = st.columns(3)

uploaded_word = None
uploaded_ppt = None
uploaded_sb3 = None

with cols[0]:
    if criteria_file_exists("word", grade):
        uploaded_word = st.file_uploader("📄 File Word (.docx)", type=["docx"], key="word")
    else:
        st.info("Word: Không áp dụng cho khối này." if grade in (3,) else "Word: Không có tiêu chí.")

with cols[1]:
    if criteria_file_exists("ppt", grade):
        uploaded_ppt = st.file_uploader("🎞️ File PowerPoint (.pptx)", type=["pptx"], key="ppt")
    else:
        st.info("PowerPoint: Không áp dụng cho khối này." if grade in (5,) else "PowerPoint: Không có tiêu chí.")

with cols[2]:
    if criteria_file_exists("scratch", grade):
        uploaded_sb3 = st.file_uploader("🐱‍💻 File Scratch (.sb3)", type=["sb3"], key="sb3")
    else:
        st.info("Scratch: Không áp dụng cho khối này." if grade in (3,) else "Scratch: Không có tiêu chí.")

st.markdown("</div>", unsafe_allow_html=True)

# Xử lý từng upload
if uploaded_word is not None:
    tmpf = save_uploaded(uploaded_word)
    # load criteria bằng core.load_criteria (trong core.py, filename = f"{subject}{grade}.json")
    criteria = load_criteria("word", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error("Không tìm thấy tiêu chí Word cho khối này.")
    else:
        score, notes = grade_word(tmpf, criteria)
        if score is None:
            st.error("Lỗi khi chấm Word: " + (notes[0] if notes else ""))
        else:
            hocsinh = pretty_name_from_filename(uploaded_word.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(selected_class, hocsinh, "Word", score, notes)

if uploaded_ppt is not None:
    tmpf = save_uploaded(uploaded_ppt)
    criteria = load_criteria("ppt", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error("Không tìm thấy tiêu chí PowerPoint cho khối này.")
    else:
        score, notes = grade_ppt(tmpf, criteria)
        if score is None:
            st.error("Lỗi khi chấm PowerPoint: " + (notes[0] if notes else ""))
        else:
            hocsinh = pretty_name_from_filename(uploaded_ppt.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(selected_class, hocsinh, "PowerPoint", score, notes)

if uploaded_sb3 is not None:
    tmpf = save_uploaded(uploaded_sb3)
    criteria = load_criteria("scratch", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error("Không tìm thấy tiêu chí Scratch cho khối này.")
    else:
        score, notes = grade_scratch(tmpf, criteria)
        if score is None:
            st.error("Lỗi khi chấm Scratch: " + (notes[0] if notes else ""))
        else:
            hocsinh = pretty_name_from_filename(uploaded_sb3.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(selected_class, hocsinh, "Scratch", score, notes)

# Hiển thị đường dẫn file Excel
st.info(f"Kết quả được lưu vào file: `{os.path.abspath(EXCEL_FILE)}` (mỗi sheet là 1 lớp).")
