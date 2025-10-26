import streamlit as st
import os
from core import load_criteria, grade_word, grade_ppt, grade_scratch, pretty_name_from_filename, ensure_workbook_exists
from openpyxl import load_workbook
import tempfile
import pandas as pd

EXCEL_FILE = "ketqua_tonghop.xlsx"
CRITERIA_FOLDER = "criteria"

CLASSES = ["3A1","3A2","3A3","3A4","4A1","4A2","4A3","4A4","4A5","5A1","5A2","5A3","5A4","5A5"]

st.set_page_config(page_title="AI-TIN Web", page_icon="🧠", layout="centered")

# --- Hiệu ứng tiêu đề ---
st.markdown("""
<style>
@keyframes rainbow {
  0% {color: #e74c3c;}
  25% {color: #f1c40f;}
  50% {color: #2ecc71;}
  75% {color: #3498db;}
  100% {color: #9b59b6;}
}

@keyframes moveTitle {
  0% {transform: translateX(0);}
  50% {transform: translateX(10px);}
  100% {transform: translateX(0);}
}

@keyframes glow {
  0% {text-shadow: 0 0 5px #fff;}
  50% {text-shadow: 0 0 20px #ffeaa7;}
  100% {text-shadow: 0 0 5px #fff;}
}

.title-container {
  text-align: center;
  margin-top: -10px;
  margin-bottom: 15px;
}

.title-main {
  font-size: 2rem;
  font-weight: 700;
  animation: moveTitle 4s ease-in-out infinite, glow 6s ease-in-out infinite;
}

.title-sub {
  font-size: 1rem;
  margin-top: -4px;
  color: #555;
  font-style: italic;
  animation: rainbow 10s linear infinite;
}
</style>

<div class="title-container">
  <div class="title-main">🧠 AI-TIN Web — Trợ lý chấm bài tự động (Tiểu học)</div>
  <div class="title-sub">💫 Giáo viên đăng nhập: <b>giaovien123</b></div>
</div>
""", unsafe_allow_html=True)

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

# --- Màu nền theo khối ---
GRADE_COLORS = {3: "#3498db", 4: "#e67e22", 5: "#9b59b6"}
GRADE_BACKGROUNDS = {
    3: "https://i.imgur.com/Wu2kZyv.jpg",
    4: "https://i.imgur.com/NXh3NGB.jpg",
    5: "https://i.imgur.com/mtvu7dR.jpg"
}

# --- Chọn lớp ---
st.sidebar.header("🎓 Chọn lớp học")
current_class = st.sidebar.selectbox("Chọn lớp", CLASSES)
grade = int(current_class[0])

bg_color = GRADE_COLORS.get(grade, "#2ecc71")
bg_image = GRADE_BACKGROUNDS.get(grade, "")

# --- Thay nền ---
st.markdown(
    f"""
    <style>
    .stApp {{
        background-image: url("{bg_image}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}
    .st-emotion-cache-18ni7ap {{
        background: rgba(255,255,255,0.85);
        border-radius: 12px;
        padding: 1rem;
    }}
    </style>
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
st.sidebar.markdown("---")
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
st.write("📥 Chọn loại bài để nộp và chấm:")

col1, col2, col3 = st.columns(3)
with col1:
    uploaded_word = st.file_uploader("Tải lên file Word (.docx)", type=["docx"], key="word") if "Word" in available_subjects else None
with col2:
    uploaded_ppt = st.file_uploader("Tải lên file PowerPoint (.pptx)", type=["pptx"], key="ppt") if "PowerPoint" in available_subjects else None
with col3:
    uploaded_sb3 = st.file_uploader("Tải lên file Scratch (.sb3)", type=["sb3"], key="sb3") if "Scratch" in available_subjects else None

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

# --- Chấm bài ---
if uploaded_word:
    tmpf = save_uploaded(uploaded_word)
    criteria = load_criteria("word", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error(f"Khối {grade} không có tiêu chí chấm Word.")
    else:
        score, notes = grade_word(tmpf, criteria)
        if score is None:
            st.error(notes[0])
        else:
            hocsinh = pretty_name_from_filename(uploaded_word.name)
            st.success(f"✅ Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "Word", score, notes)

if uploaded_ppt:
    tmpf = save_uploaded(uploaded_ppt)
    criteria = load_criteria("ppt", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error(f"Khối {grade} không có tiêu chí chấm PowerPoint.")
    else:
        score, notes = grade_ppt(tmpf, criteria)
        if score is None:
            st.error(notes[0])
        else:
            hocsinh = pretty_name_from_filename(uploaded_ppt.name)
            st.success(f"✅ Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "PowerPoint", score, notes)

if uploaded_sb3:
    tmpf = save_uploaded(uploaded_sb3)
    criteria = load_criteria("scratch", grade, CRITERIA_FOLDER)
    if criteria is None:
        st.error(f"Khối {grade} không có tiêu chí chấm Scratch.")
    else:
        score, notes = grade_scratch(tmpf, criteria)
        if score is None:
            st.error(notes[0])
        else:
            hocsinh = pretty_name_from_filename(uploaded_sb3.name)
            st.success(f"✅ Điểm: {score}/10")
            for n in notes: st.write("• " + n)
            append_to_sheet(current_class, hocsinh, "Scratch", score, notes)

# --- Hiển thị kết quả tổng hợp ---
st.markdown("---")
st.subheader("📊 Xem bảng điểm tổng hợp")

if os.path.exists(EXCEL_FILE):
    wb = load_workbook(EXCEL_FILE)
    sheets = wb.sheetnames
    selected_sheet = st.selectbox("Chọn lớp để xem", sheets)
    if selected_sheet:
        df = pd.DataFrame(wb[selected_sheet].values)
        df.columns = df.iloc[0]
        df = df[1:]
        st.dataframe(df, use_container_width=True)
        st.download_button("⬇️ Tải file Excel", data=open(EXCEL_FILE, "rb"), file_name="ketqua_tonghop.xlsx")
else:
    st.info("📁 Chưa có dữ liệu chấm nào được lưu.")

st.info(f"🗂️ Kết quả được lưu vào: `{os.path.abspath(EXCEL_FILE)}` (mỗi sheet là 1 lớp).")
