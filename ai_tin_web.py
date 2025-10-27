from flask import Flask, render_template, request
import json
import os

app = Flask(__name__)

# --- Đường dẫn tới thư mục chứa các file tiêu chí ---
CRITERIA_PATH = os.path.join(os.path.dirname(__file__), 'criteria')

# --- Hàm đọc file JSON tiêu chí ---
def load_criteria(file_name):
    try:
        file_path = os.path.join(CRITERIA_PATH, file_name)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return None
    except Exception as e:
        print(f"Lỗi đọc file {file_name}: {e}")
        return None


# --- Hàm xác định khối từ tên lớp ---
def get_grade_from_class(class_name):
    if not class_name:
        return None
    for c in class_name:
        if c.isdigit():
            return int(c)
    return None


# --- Hàm chọn hình nền theo khối ---
def get_background_for_grade(grade):
    if grade == 3:
        return "bg_3.jpg"
    elif grade == 4:
        return "bg_4.jpg"
    elif grade == 5:
        return "bg_5.jpg"
    else:
        return "bg_default.jpg"


# --- Hàm chọn file tiêu chí phù hợp ---
def get_criteria_file(grade, software):
    mapping = {
        3: {"PowerPoint": "ppt3.json"},
        4: {"PowerPoint": "ppt4.json", "Word": "word4.json", "Scratch": "scratch4.json"},
        5: {"PowerPoint": "ppt5.json", "Word": "word5.json", "Scratch": "scratch5.json"}
    }

    # Kiểm tra nếu khối có học phần mềm này không
    if grade not in mapping or software not in mapping[grade]:
        return None
    return mapping[grade][software]


@app.route('/', methods=['GET', 'POST'])
def index():
    selected_class = None
    selected_software = None
    grade = None
    criteria_data = None
    message = None

    if request.method == 'POST':
        selected_class = request.form.get('class_name')
        selected_software = request.form.get('software')

        grade = get_grade_from_class(selected_class)
        if grade:
            file_name = get_criteria_file(grade, selected_software)
            if file_name:
                criteria_data = load_criteria(file_name)
                if not criteria_data:
                    message = f"Không thể đọc tiêu chí cho {selected_software} khối {grade}."
            else:
                message = f"Khối {grade} không học phần mềm {selected_software}."
        else:
            message = "Không xác định được khối học từ tên lớp."

    background = get_background_for_grade(grade) if grade else "bg_default.jpg"

    return render_template(
        'index.html',
        selected_class=selected_class,
        selected_software=selected_software,
        criteria=criteria_data,
        message=message,
        background=background
    )


if __name__ == '__main__':
    app.run(debug=True)
