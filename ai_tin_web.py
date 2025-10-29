from flask import Flask, render_template, request, url_for
import json
import os

# --- Khởi tạo ứng dụng Flask ---
app = Flask(__name__, static_folder="static", template_folder="templates")

# --- Đường dẫn thư mục tiêu chí ---
CRITERIA_DIR = os.path.join(os.path.dirname(__file__), "criteria")


# --- Xác định nền theo khối ---
def get_background_for_grade(grade):
    return {
        "3": "bg_3.jpg",
        "4": "bg_4.jpg",
        "5": "bg_5.jpg"
    }.get(grade, "bg_default.jpg")


# --- Hàm tải tiêu chí từ file JSON ---
def load_criteria(software, grade, criteria_folder="criteria"):
    """Đọc tiêu chí chấm theo phần mềm và khối"""
    # Trường hợp không học phần mềm
    if software == "word" and str(grade) == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Word", "diem": ""}]}
    elif software == "powerpoint" and str(grade) == "5":
        return {"tieu_chi": [{"mo_ta": "Khối 5 không học phần mềm PowerPoint", "diem": ""}]}
    elif software == "scratch" and str(grade) == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Scratch", "diem": ""}]}

    # Đọc file JSON
    filename = f"{software}{grade}.json"
    file_path = os.path.join(criteria_folder, filename)

    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            return {"tieu_chi": [{"mo_ta": f"Lỗi đọc file {filename}: {e}", "diem": ""}]}
    else:
        return None


# --- Trang chính ---
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
            grade = selected_class[0]  # Lấy số khối từ lớp (ví dụ: 4A2 -> 4)
            background = get_background_for_grade(grade)

            if selected_software:
                # 🔧 Dòng này đã sửa đúng cú pháp để liên kết tiêu chí
                criteria = load_criteria(selected_software, grade, criteria_folder=CRITERIA_DIR)

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


if __name__ == "__main__":
    app.run(debug=True)
