# core.py – Mô-đun trung tâm cho AI-TIN Web
import os
import json
from openpyxl import Workbook


# ==================================================
# 🧩 1️⃣ HÀM TẢI TIÊU CHÍ CHẤM
# ==================================================
def load_criteria(critfile, grade, folder="criteria"):
    """
    Tải tiêu chí chấm cho một phần mềm (Word/PPT/Scratch) và khối lớp.
    Ưu tiên đọc file JSON từ thư mục criteria. 
    Nếu không có file, trả về bộ tiêu chí mẫu để demo.
    """
    path = os.path.join(folder, f"{critfile}{grade}.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print("Lỗi khi đọc tiêu chí:", e)
            return None

    # Trường hợp chưa có file JSON thì trả về tiêu chí mẫu
    return {
        "tieu_chi": [
            {"mo_ta": "Hoàn thành đúng yêu cầu bài", "diem": 5},
            {"mo_ta": "Trình bày đẹp, rõ ràng", "diem": 3},
            {"mo_ta": "Có yếu tố sáng tạo", "diem": 2},
        ]
    }


# ==================================================
# 🧩 2️⃣ HÀM CHẤM BÀI WORD
# ==================================================
def grade_word(file_path, criteria):
    """
    Giả lập chấm bài Word. Trả về điểm và nhận xét.
    Trong phiên bản thực tế, có thể dùng python-docx để kiểm tra nội dung.
    """
    try:
        total = sum(item["diem"] for item in criteria["tieu_chi"])
        notes = [f"{item['mo_ta']} (+{item['diem']}đ)" for item in criteria["tieu_chi"]]
        return min(total, 10), notes
    except Exception as e:
        return None, [f"Lỗi khi chấm Word: {e}"]


# ==================================================
# 🧩 3️⃣ HÀM CHẤM BÀI POWERPOINT
# ==================================================
def grade_ppt(file_path, criteria):
    """
    Giả lập chấm bài PowerPoint.
    Có thể dùng python-pptx để đọc nội dung slide trong tương lai.
    """
    try:
        total = sum(item["diem"] for item in criteria["tieu_chi"]) - 1  # ví dụ điểm thấp hơn 1
        notes = [f"{item['mo_ta']} (+{item['diem']}đ)" for item in criteria["tieu_chi"]]
        return min(total, 10), notes
    except Exception as e:
        return None, [f"Lỗi khi chấm PowerPoint: {e}"]


# ==================================================
# 🧩 4️⃣ HÀM CHẤM BÀI SCRATCH
# ==================================================
def grade_scratch(file_path, criteria):
    """
    Giả lập chấm bài Scratch (.sb3)
    Có thể dùng json để đọc project.sb3 trong tương lai.
    """
    try:
        total = sum(item["diem"] for item in criteria["tieu_chi"]) - 2  # ví dụ điểm thấp hơn 2
        notes = [f"{item['mo_ta']} (+{item['diem']}đ)" for item in criteria["tieu_chi"]]
        return max(min(total, 10), 0), notes
    except Exception as e:
        return None, [f"Lỗi khi chấm Scratch: {e}"]


# ==================================================
# 🧩 5️⃣ HÀM XỬ LÝ TÊN HỌC SINH
# ==================================================
def pretty_name_from_filename(filename):
    """
    Chuyển tên file thành tên dễ đọc để hiển thị.
    Ví dụ: 'le_thi_bich_3a1.docx' -> 'Le Thi Bich 3A1'
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    name = name.replace("_", " ").replace("-", " ").title()
    return name


# ==================================================
# 🧩 6️⃣ HÀM TẠO FILE EXCEL NẾU CHƯA CÓ
# ==================================================
def ensure_workbook_exists(path="ketqua_tonghop.xlsx"):
    """
    Kiểm tra nếu file Excel tổng hợp chưa tồn tại thì tạo mới.
    """
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "TỔNG HỢP"
        ws.append(["Họ tên học sinh", "Môn", "Điểm", "Nhận xét"])
        wb.save(path)
    return path
