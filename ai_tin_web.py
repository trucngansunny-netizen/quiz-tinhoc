# ai_tin_web.py
import streamlit as st
import os
from core import load_criteria, grade_word, grade_ppt, grade_scratch, pretty_name_from_filename, ensure_workbook_exists
from openpyxl import load_workbook
import tempfile

EXCEL_FILE = "ketqua_tonghop.xlsx"
CRITERIA_FOLDER = "criteria"

CLASSES = ["3A1","3A2","3A3","3A4","4A1","4A2","4A3","4A4","4A5","5A1","5A2","5A3","5A4","5A5"]

st.set_page_config(page_title="AI-TIN Web", page_icon="🧠", layout="centered")

# --- Logo ---
logo_path = "assets/logo_tranquoctoan.png"
if os.path.exists(logo_path):
    st.markdown(
        f"""
        <div style='position:fixed; top:10px; left:15px; z-index:100;'>
            <img src='{logo_path}' width='60' style='border-radius:50%; opacity:0.95;'>
        </div>
        """,
        unsafe_allow_html=True
    )

# --- Màu nền + hình nền theo khối ---
GRADE_COLORS = {3: "#3498db", 4: "#e67e22", 5: "#9b59b6"}
GRADE_BACKGROUNDS = {
    3: "https://i.imgur.com/Wu2kZyv.jpg",
    4: "https://i.imgur.com/NXh3NGB.jpg",
    5: "https://i.imgur.com/mtvu7dR.jpg"
}

# --- Sidebar chọn lớp ---
st.sidebar.header("🎓 Chọn lớp học")
current_class = st.sidebar.selectbox("Chọn lớp", CLASSES)
grade = int(current_class[0])

# --- Thiết lập màu theo khối ---
bg_color = GRADE_COLORS.get(grade, "#2ecc71")
bg_image = GRADE_BACKGROUNDS.get(grade, "")
text_color = "#222"

if grade == 3:
    text_color = "#1B4F72"
elif grade == 4:
    text_color = "#7E5109"
elif grade == 5:
    text_color = "#512E5F"

# --- CSS giao diện tổng thể ---
st.markdown(
    f"""
    <style>
    .stApp {{
        background-image: url("{bg_image}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        color: {text_color};
        font-family: 'Segoe UI', sans-serif;
    }}

    .title-container {{
        text-align: center;
        padding: 1rem 0;
        border-radius: 16px;
        background: rgba(255,255,255,0.8);
        margin-bottom: 1rem;
        box-shadow: 0 0 20px rgba(255,255,255,0.4);
    }}

    /* 💫 Tiêu đề 3 dòng */
    .title-move {{
        font-size: 2.2rem;
        font-weight: 800;
        color: {bg_color};
        animation: slide 3s ease-in-out infinite alternate;
    }}
    @keyframes slide {{
        0% {{ transform: translateX(-6px); }}
        100% {{ transform: translateX(6px); }}
    }}

    .title-glow {{
        font-size: 1.5rem;
        font-weight: 600;
        color: #333;
        animation: glow 2.5s ease-in-out infinite alternate;
    }}
    @keyframes glow {{
        0% {{ opacity: 0.6; text-shadow: 0 0 5px {bg_color}; }}
        100% {{ opacity: 1; text-shadow: 0 0 15px {bg_color}; }}
    }}

    .title-rainbow {{
        font-size: 1.3rem;
        font-weight: 500;
        background: linear-gradient(90deg, red, orange, yellow, green, cyan, blue, violet);
        -webkit-background-clip: text;
        color: transparent;
        animation: rainbow 4s linear infinite;
    }}
    @keyframes rainbow {{
        0% {{ filter: hue-rotate(0deg); }}
        100% {{ filter: hue-rotate(360deg); }}
    }}

    /* 🌟 Nút upload và khung kết quả */
    .stFileUploader label {{
        border: 2px solid {bg_color};
        border-radius: 12px;
        padding: 0.4rem 1rem;
        background: rgba(255,255,255,0.8);
        box-shadow: 0 0 10px rgba(0,0,0,0.15);
        transition: all 0.3s ease-in-out;
    }}
    .stFileUploader label:hover {{
        box-shadow: 0 0 15px {bg_color};
        transform: scale(1.02);
    }}
    .stAlert {{
        border-radius: 10px !important;
        border: 1.5px solid {bg_color} !important;
        background: rgba(255,255,255,0.85) !important;
    }}
    </style>

    <div class="title-container">
        <div class="title-move">🧠 AI-TIN Web</div>
        <div class="title-glow">Trợ lý chấm bài tự động</div>
        <div class="title-rainbow">Trường Tiểu học Trần Quốc Toản</div>
    </div>
    """,
    unsafe_allow_html=True
)

# --- Môn học hợp lệ ---
AVAILABLE_BY_GRADE = {
    3: ["PowerPoint"],
    4: ["Word", "PowerPoint", "Scratch"],
    5: ["Word", "Scratch"]
}

available_subjects = AVAILABLE_BY_GRADE.get(grade, [])

st.sidebar.subheader(f"📚 Tiêu chí chấm cho khối {grade}")
if available_subjects:
    subj_show = st.sidebar.selectbox("Xem tiêu chí cho", available_subjects)
    critfile = {"Word": "word", "PowerPoint": "ppt", "Scratch": "scratch"}[subj_show]
    criteria = load_criteria(critfile, grade, CRITERIA_FOLDER)
    if criteria is None:
        st.sidebar.warning(f"⚠️ Chưa có tiêu chí cho {subj_show} khối {grade}.")
    else:
        for it in criteria.get("tieu_chi", []):
            st.sidebar.write(f"- {it.get('mo_ta')} ({it.get('diem')}đ)")
else:
    st.sidebar.error(f"❌ Khối {grade} hiện chưa có môn nào để chấm.")

# --- Nộp bài ---
st.write(f"**Lớp đang chọn:** {current_class} (Khối {grade})")
st.write("Chọn loại bài để nộp và chấm:")

col1, col2, col3 = st.columns(3)
with col1:
    uploaded_word = st.file_uploader("📄 File Word (.docx)", type=["docx"], key="word") if "Word" in available_subjects else None
with col2:
    uploaded_ppt = st.file_uploader("🎞️ File PowerPoint (.pptx)", type=["pptx"], key="ppt") if "PowerPoint" in available_subjects else None
with col3:
    uploaded_sb3 = st.file_uploader("🐱 File Scratch (.sb3)", type=["sb3"], key="sb3") if "Scratch" in available_subjects else None

def save_uploaded(tmpfile):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(tmpfile.name)[1])
    tmp.write(tmpfile.getbuffer())
    tmp.close()
    return tmp.name

def append_to_sheet(class_name, hocsinh, subject, score, notes):
    ensure_workbook_exists(EXCEL_FILE)
    wb = load_workbook(EXCEL_FILE)
    if class_name not in wb.sheetnames:
        ws = wb.create_sheet(title=class_name)
        ws.append(["Họ tên học sinh","Môn","Điểm","Nhận xét"])
    else:
        ws = wb[class_name]
    ws.append([hocsinh, subject, score, "; ".join(notes)])
    wb.save(EXCEL_FILE)

# --- Chấm ---
if uploaded_word:
    tmpf = save_uploaded(uploaded_word)
    criteria = load_criteria("word", grade, CRITERIA_FOLDER)
    if criteria:
        score, notes = grade_word(tmpf, criteria)
        if score is not None:
            hocsinh = pretty_name_from_filename(uploaded_word.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "Word", score, notes)

if uploaded_ppt:
    tmpf = save_uploaded(uploaded_ppt)
    criteria = load_criteria("ppt", grade, CRITERIA_FOLDER)
    if criteria:
        score, notes = grade_ppt(tmpf, criteria)
        if score is not None:
            hocsinh = pretty_name_from_filename(uploaded_ppt.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "PowerPoint", score, notes)

if uploaded_sb3:
    tmpf = save_uploaded(uploaded_sb3)
    criteria = load_criteria("scratch", grade, CRITERIA_FOLDER)
    if criteria:
        score, notes = grade_scratch(tmpf, criteria)
        if score is not None:
            hocsinh = pretty_name_from_filename(uploaded_sb3.name)
            st.success(f"💯 Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "Scratch", score, notes)

st.info(f"Kết quả được lưu vào file: `{os.path.abspath(EXCEL_FILE)}` (mỗi sheet là 1 lớp).")
