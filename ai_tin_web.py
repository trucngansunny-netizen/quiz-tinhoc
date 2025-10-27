from flask import Flask, render_template, request
import json
import os

# --- Khởi tạo ứng dụng Flask ---
app = Flask(__name__, static_folder="static", template_folder="templates")

# --- Đường dẫn thư mục tiêu chí ---
CRITERIA_DIR = os.path.join(os.path.dirname(__file__), "criteria")


# --- Hàm xác định hình nền theo khối ---
def get_background_for_grade(grade):
    return {
        "3": "bg_3.jpg",
        "4": "bg_4.jpg",
        "5": "bg_5.jpg"
    }.get(grade, "bg_default.jpg")


# --- Hàm tải file tiêu chí ---
def load_criteria(software, grade):
    # Trường hợp không học phần mềm
    if software == "word" and grade == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Word", "diem": ""}]}
    elif software == "powerpoint" and grade == "5":
        return {"tieu_chi": [{"mo_ta": "Khối 5 không học phần mềm PowerPoint", "diem": ""}]}
    elif software == "scratch" and grade == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Scratch", "diem": ""}]}

    # Các trường hợp khác thì đọc file JSON
    filename = f"{software}{grade}.json"
    file_path = os.path.join(CRITERIA_DIR, filename)

    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return None


# --- Trang chủ ---
@app.route("/", methods=["GET", "POST"])
def home():
    selected_class = None
    selected_software = None
    criteria = None
    message = None
    background = "bg_default.jpg"

    if request.method == "POST":
        selected_class = request.form.get("lop", "")
        selected_software = request.form.get("software", "")

        if selected_class:
            grade = selected_class[0]  # lấy số đầu trong "5A" -> "5"
            background = get_background_for_grade(grade)

            if selected_software:
                criteria = load_criteria(selected_software, grade)
                if not criteria:
                    message = f"⚠️ Khối {grade} hiện chưa có tiêu chí hiển thị (hoặc không học phần mềm này)."
            else:
                message = "⚠️ Vui lòng chọn phần mềm để xem tiêu chí."
        else:
            message = "⚠️ Vui lòng chọn lớp học."

    return render_template(
        "index.html",
        selected_class=selected_class,
        selected_software=selected_software,
        criteria=criteria,
        message=message,
        background=background
    )


# --- Chạy app ---
if __name__ == "__main__":
    # Render hoặc môi trường deploy sẽ cung cấp biến PORT
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
