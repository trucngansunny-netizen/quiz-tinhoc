# ai_tin_web.py
import os
from flask import Flask, render_template, request
from datetime import datetime
from cham_tieuchi import load_criteria, grade_word, grade_ppt, grade_scratch

app = Flask(__name__)

COUNTER_FILE = "counter.txt"

def increase_counter():
    if not os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE,"w",encoding="utf-8") as f: f.write("0")
    with open(COUNTER_FILE,"r+",encoding="utf-8") as f:
        num = int(f.read().strip() or 0)
        num += 1
        f.seek(0)
        f.write(str(num))
    return num

def save_student_file(file, subject, lop):
    today = datetime.now().strftime("%Y-%m-%d")
    folder = os.path.join("uploads", lop, subject, today)
    os.makedirs(folder, exist_ok=True)
    save_path = os.path.join(folder, file.filename)
    file.save(save_path)
    return save_path

def cham_file(path, subject, grade):
    criteria = load_criteria(subject, grade)
    if not criteria:
        return None, ["Không tìm thấy tiêu chí chấm"]

    if subject.lower() == "word":
        return grade_word(path, criteria)
    elif subject.lower() == "ppt":
        return grade_ppt(path, criteria)
    elif subject.lower() == "scratch":
        return grade_scratch(path, criteria)
    return None, ["Môn không hợp lệ"]

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/grade", methods=["POST"])
def grade():
    file = request.files["file"]
    subject = request.form.get("subject")
    lop = request.form.get("lop")
    grade = request.form.get("grade")

    # lưu file
    saved_path = save_student_file(file, subject, lop)

    # tăng bộ đếm
    count = increase_counter()

    # chấm bài
    score, errors = cham_file(saved_path, subject, grade)

    return render_template(
        "index.html",
        score=score,
        errors=errors,
        count=count,
        filename=file.filename
    )

if __name__ == "__main__":
    app.run(debug=True)
