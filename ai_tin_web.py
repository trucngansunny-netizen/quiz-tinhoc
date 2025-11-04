import os, json, shutil, tempfile, datetime
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook
from cham_tieuchi import pretty_name_from_filename, load_criteria, grade_word, grade_ppt, grade_scratch

APP_ROOT=os.path.dirname(__file__)
STATIC_DIR=os.path.join(APP_ROOT,"static")
CRITERIA_DIR=os.path.join(APP_ROOT,"criteria")
RESULTS_DIR=os.path.join(APP_ROOT,"results")
TONGHOP_FILE=os.path.join(RESULTS_DIR,"tonghop.xlsx")
VISIT_FILE=os.path.join(RESULTS_DIR,"visits.txt")
ALLOWED_EXT={"pptx","docx","sb3","zip"}
CLASSES=["3A1","3A2","3A3","3A4","4A1","4A2","4A3","4A4","4A5","5A1","5A2","5A3","5A4","5A5"]
AVAILABLE_BY_GRADE={3:["PowerPoint"],4:["Word","PowerPoint","Scratch"],5:["Word","Scratch"]}
os.makedirs(STATIC_DIR,exist_ok=True); os.makedirs(CRITERIA_DIR,exist_ok=True); os.makedirs(RESULTS_DIR,exist_ok=True)

app=Flask(__name__,static_folder="static",template_folder="templates")

def allowed_file(fn): return fn.rsplit(".",1)[-1].lower() in ALLOWED_EXT
def read_visit():
    if not os.path.exists(VISIT_FILE): return 0
    try: return int(open(VISIT_FILE,"r",encoding="utf-8").read().strip() or "0")
    except: return 0
def increase_visit():
    c=read_visit()+1
    with open(VISIT_FILE,"w",encoding="utf-8") as f: f.write(str(c))
    return c
def ensure_tonghop():
    if not os.path.exists(TONGHOP_FILE):
        wb=Workbook(); wb.remove(wb.active)
        for cls in CLASSES:
            ws=wb.create_sheet(cls)
            ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm"])
        wb.save(TONGHOP_FILE)
def append_result(cls,grade,student,subject,filename,total):
    ensure_tonghop()
    wb=load_workbook(TONGHOP_FILE)
    if cls not in wb.sheetnames:
        ws=wb.create_sheet(cls)
        ws.append(["Họ và tên","Lớp","Khối","Môn học","Tên tệp","Điểm tổng","Ngày chấm"])
    ws=wb[cls]
    now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([student,cls,grade,subject,filename,total,now])
    wb.save(TONGHOP_FILE)

@app.route("/",methods=["GET","POST"])
def index():
    visit=increase_visit()
    msg=None; result=None
    if request.method=="POST":
        grade=request.form.get("grade"); cls=request.form.get("class"); subject=request.form.get("subject")
        file=request.files.get("file")
        if not (grade and cls and subject and file and file.filename):
            msg="Vui lòng chọn đầy đủ thông tin và tệp!"
        else:
            fn=secure_filename(file.filename)
            if not allowed_file(fn): msg="Định dạng tệp không được hỗ trợ!"
            else:
                tmp=tempfile.mkdtemp(prefix="ai_tin_"); path=os.path.join(tmp,fn); file.save(path)
                try: criteria=load_criteria(subject.lower(),int(grade),CRITERIA_DIR)
                except: criteria=None
                if not criteria: msg=f"Chưa có tiêu chí cho {subject} khối {grade}."
                else:
                    try:
                        if subject=="Word": total,notes=grade_word(path,criteria)
                        elif subject=="PowerPoint": total,notes=grade_ppt(path,criteria)
                        elif subject=="Scratch": total,notes=grade_scratch(path,criteria)
                        else: total=None; notes=["Môn học không hợp lệ."]
                    except Exception as e: total=None; notes=[str(e)]
                    if total is not None:
                        append_result(cls,grade,pretty_name_from_filename(fn),subject,fn,total)
                        result={"student":pretty_name_from_filename(fn),"class":cls,"grade":grade,
                                "subject":subject,"file":fn,"total":total,"notes":notes}
                    else: msg="Lỗi khi chấm bài."
                shutil.rmtree(tmp,ignore_errors=True)
    return render_template("index.html",classes=CLASSES,avail_by_grade=AVAILABLE_BY_GRADE,visit_count=read_visit(),
                           result=result,message=msg)

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5000,debug=False)
