# ---------------- PPT grading ----------------
import os
import json
import zipfile
import tempfile
import shutil
import re
import unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- Utilities ----------------
def pretty_name_from_filename(filename):
    """Lấy tên gợi ý từ tên file (ví dụ: NguyenVanA_Lop4A1.pptx -> Nguyen Van A Lop 4 A1)"""
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)


def load_criteria(subject, grade, criteria_folder="criteria"):
    """Tải file tiêu chí chấm điểm theo phần mềm và khối"""
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = subject.lower()
    if s in ("powerpoint", "ppt", "pptx"):
        pref = "ppt"
    elif s in ("word", "docx", "doc"):
        pref = "word"
    elif s in ("scratch", "sb3"):
        pref = "scratch"
    else:
        pref = s

    # Tìm file tiêu chí
    possible_names = [
        f"{pref}{grade}.json",
        f"{pref}_khoi{grade}.json",
        f"{pref}-khoi{grade}.json"
    ]
    for name in possible_names:
        path = os.path.join(criteria_folder, name)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and "tieu_chi" in data:
                    return data
                elif isinstance(data, list):
                    return {"tieu_chi": data}
            except Exception:
                pass
    return None

def _shape_has_picture(shape):
    """Nhận diện toàn diện tất cả các loại ảnh trong PowerPoint."""
    try:
        # 1️⃣ Ảnh trực tiếp (Picture shape)
        if getattr(shape, "shape_type", None) == 13:  # MSO_SHAPE_TYPE.PICTURE
            return True

        # 2️⃣ Ảnh nhóm (Group shape)
        if hasattr(shape, "shapes"):
            for s in shape.shapes:
                if _shape_has_picture(s):
                    return True

        # 3️⃣ Ảnh trong fill (background image hoặc pattern fill)
        if hasattr(shape, "fill") and hasattr(shape.fill, "type"):
            try:
                fill = shape.fill
                # Kiểm tra fill là kiểu picture hoặc XML chứa tag blip
                if getattr(fill, "type", None) == 6:
                    return True
                xml_fill = getattr(fill, "_xFill", None)
                if xml_fill is not None and "blip" in str(xml_fill).lower():
                    return True
            except Exception:
                pass

        # 4️⃣ Kiểm tra XML trực tiếp: mọi tag chứa ảnh (p:pic, a:blip, a:blipFill, blipFill, r:link)
        xml_str = shape._element.xml.lower()
        if any(tag in xml_str for tag in [
            "p:pic", "a:blip", "a:blipfill", "blipfill", "r:link", "r:embed", "a14:imgprops", "a:stretch"
        ]):
            return True

    except Exception:
        pass
    return False


def grade_ppt(file_path, criteria):
    from pptx import Presentation
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    for it in items:
        try:
            it["diem"] = float(it.get("diem", 0))
        except:
            it["diem"] = 0.0

    num_slides = len(slides)
    total = 0.0
    notes = []

    # --- Kiểm tra hình ảnh (rộng nhất có thể) ---
    has_picture_any = False
    for slide in slides:
        try:
            xml = slide._element.xml.lower()
            if any(tag in xml for tag in ["p:pic", "a:blip", "a:blipfill", "r:embed", "r:link", "blipfill"]):
                has_picture_any = True
                break
            for shape in slide.shapes:
                if _shape_has_picture(shape):
                    has_picture_any = True
                    break
            if has_picture_any:
                break
        except Exception:
            continue

    # --- Kiểm tra hiệu ứng chuyển slide ---
    has_transition_any = any("transition" in slide._element.xml for slide in slides)

    # --- Gom text (cả có dấu và không dấu) ---
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                ppt_text += " " + (shape.text or "")
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    # --- Chấm theo tiêu chí ---
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "has_image":
            ok = has_picture_any
        elif key == "has_transition":
            ok = has_transition_any
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(s.has_text_frame and s.text.strip() for s in first.shapes)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in ppt_text_noaccent)

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total, 2), notes


