from flask import Flask, render_template, request, url_for
import json
import os

app = Flask(__name__)

# --- Đường dẫn thư mục tiêu chí ---
CRITERIA_DIR = os.path.join(os.path.dirname(__file__), "criteria")

# --- Xác định nền theo khối ---
def get_background_for_grade(grade):
    return {
        "3": "bg_3.jpg",
        "4": "bg_4.jpg",
        "5": "bg_5.jpg"
    }.get(grade, "bg_default.jpg")


# --- Tải file tiêu chí ---
def load_criteria(software, grade):
    filename = None

    # Map file theo khối và phần mềm
    if software == "word" and grade == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Word", "diem": ""}]}
    elif software == "powerpoint" and grade == "5":
        return {"tieu_chi": [{"mo_ta": "Khối 5 không học phần mềm PowerPoint", "diem": ""}]}
    elif software == "scratch" and grade == "3":
        return {"tieu_chi": [{"mo_ta": "Khối 3 không học phần mềm Scratch", "diem": ""}]}
    else:
        filename = f"{software}{grade}.json"

    # 🔧 Chỉ sửa đúng dòng dưới đây, thêm CRITERIA_DIR vào đường dẫn
    file_path = os.path.join(CRITERIA_DIR, filename)

    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return None


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
            grade = selected_class[0]  # Lấy số khối từ lớp
            background = get_background_for_grade(grade)

            if selected_software:
                # 🔧 Chỉ sửa đúng dòng này
                criteria = load_criteria(selected_software, grade)
                if not criteria:
                    message = f"⚠️ Khối {grade} hiện chưa có tiêu chí hiển thị (hoặc không học phần mềm này)."
            else:
                message = "⚠️ Vui lòng chọn phần mềm để xem tiêu chí."
        else:
            message = "⚠️ Vui lòng chọn lớp học."

       # --- trước khi trả về template: tạo dict avail_by_grade để template dùng (tránh Undefined) ---
    # subjects và tên file tương ứng trong thư mục criteria
    subj_file_map = {
        "word":   lambda g: f"word{g}.json",
        "powerpoint": lambda g: f"ppt{g}.json",
        "scratch": lambda g: f"scratch{g}.json",
    }

    avail_by_grade = {}
    for g in ("3", "4", "5"):
        avail = {}
        for subj, fn_builder in subj_file_map.items():
            fname = fn_builder(g)
            path = os.path.join(CRITERIA_DIR, fname)
            avail[subj] = os.path.exists(path)
        avail_by_grade[g] = avail

    return render_template(
        "index.html",
        selected_class=selected_class,
        selected_software=selected_software,
        criteria=criteria,
        message=message,
        background=background,
        avail_by_grade=avail_by_grade  # <-- thêm biến này cho template (đã có tojson)
    )

if __name__ == "__main__":
    app.run(debug=True)

