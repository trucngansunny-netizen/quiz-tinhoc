import streamlit as st
import pandas as pd
import os, datetime
from openpyxl import Workbook, load_workbook
from cham_tieuchi import grade_word, grade_ppt, grade_scratch, load_criteria, pretty_name_from_filename

# ---------------------------
# Cấu hình trang web
# ---------------------------
st.set_page_config(page_title="Hệ thống chấm Tin học Tiểu học", layout="wide")

# Thư mục chính
RESULTS_DIR = "results"
CRITERIA_DIR = "criteria"
EXCEL_PATH = os.path.join(RESULTS_DIR, "tonghop.xlsx")
COUNTER_PATH = "results/visits.txt"

os.makedirs(RESULTS_DIR, exist_ok=True)

# ---------------------------
# Quản lý lượt truy cập
# ---------------------------
def update_visit_count():
    count = 0
    if os.path.exists(COUNTER_PATH):
        with open(COUNTER_PATH, "r", encoding="utf-8") as f:
            try:
                count = int(f.read().strip())
            except:
                count = 0
    count += 1
    with open(COUNTER_PATH, "w", encoding="utf-8") as f:
        f.write(str(count))
    return count

visit_count = update_visit_count()

# ---------------------------
# Chuẩn bị file Excel tổng
# ---------------------------
SHEETS = ["3A1","3A2","3A3","3A4","4A1","4A2","4A3","4A4","4A5",
          "5A1","5A2","5A3","5A4","5A5"]

if not os.path.exists(EXCEL_PATH):
    wb = Workbook()
    for i, sheet_name in enumerate(SHEETS):
        ws = wb.create_sheet(sheet_name, i)
        ws.append(["Họ tên", "Lớp", "Khối", "Tên tệp", "Điểm tổng", "Ngày chấm", "Nhận xét"])
    if "Sheet" in wb.sheetnames:
        std = wb["Sheet"]; wb.remove(std)
    wb.save(EXCEL_PATH)

# ---------------------------
# Giao diện nền
# ---------------------------
grade_bg = {"3": "bg_3.jpg", "4": "bg_4.jpg", "5": "bg_5.jpg"}
default_bg = os.path.join("static", "bg_default.jpg")

def set_background(grade=None):
    bg_file = grade_bg.get(str(grade), "bg_default.jpg")
    bg_path = os.path.join("static", bg_file)
    if os.path.exists(bg_path):
        with open(bg_path, "rb") as f:
            img_bytes = f.read()
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/jpg;base64,{img_bytes.hex()}");
                background-size: cover;
                background-position: center;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

# ---------------------------
# Giao diện chọn khối/lớp
# ---------------------------
st.markdown("<h1 style='text-align:center;color:#0055aa;'>HỆ THỐNG CHẤM TIN HỌC TIỂU HỌC</h1>", unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    grade = st.selectbox("Chọn khối:", ["3", "4", "5"])
with col2:
    lop = st.selectbox("Chọn lớp:", [s for s in SHEETS if s.startswith(grade)])

set_background(grade)

st.markdown("---")
uploaded_file = st.file_uploader("Chọn tệp bài làm của em", type=["docx", "pptx", "sb3"])

# ---------------------------
# Xử lý chấm điểm
# ---------------------------
if uploaded_file:
    file_bytes = uploaded_file.read()
    tmp_path = os.path.join("results", uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(file_bytes)

    name = pretty_name_from_filename(uploaded_file.name)
    ext = os.path.splitext(uploaded_file.name)[1].lower()

    # Xác định môn theo đuôi file
    if ext == ".docx":
        subject = "word"
    elif ext == ".pptx":
        subject = "ppt"
    elif ext == ".sb3":
        subject = "scratch"
    else:
        st.error("Định dạng file không hợp lệ.")
        st.stop()

    criteria = load_criteria(subject, int(grade), CRITERIA_DIR)
    if not criteria:
        st.error(f"Không tìm thấy tiêu chí cho {subject} khối {grade}.")
        st.stop()

    # Chấm
    if subject == "word":
        total, notes = grade_word(tmp_path, criteria)
    elif subject == "ppt":
        total, notes = grade_ppt(tmp_path, criteria)
    else:
        total, notes = grade_scratch(tmp_path, criteria)

    if total is None:
        st.error("❌ Lỗi khi chấm file, vui lòng thử lại.")
        st.stop()

    # Cộng điểm chính xác theo tiêu chí JSON
    total = round(total, 2)
    nhan_xet = "Hoàn thành tốt" if total >= 8 else "Cần cố gắng hơn"

    # Ghi kết quả vào Excel
    wb = load_workbook(EXCEL_PATH)
    if lop not in wb.sheetnames:
        ws = wb.create_sheet(lop)
        ws.append(["Họ tên", "Lớp", "Khối", "Tên tệp", "Điểm tổng", "Ngày chấm", "Nhận xét"])
    ws = wb[lop]
    ws.append([name, lop, grade, uploaded_file.name, total,
               datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), nhan_xet])
    wb.save(EXCEL_PATH)

    # Thông báo nộp thành công
    st.success("✅ Bài làm đã nộp thành công!")

st.markdown(f"<p style='text-align:right;color:#444;'>Lượt truy cập: {visit_count}</p>", unsafe_allow_html=True)
