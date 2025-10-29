import os
import json
import zipfile
import tempfile
import shutil
import re
import unicodedata

from docx import Document
from pptx import Presentation

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
    """Chuẩn hóa chuỗi: bỏ dấu tiếng Việt, đưa về chữ thường."""
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    patterns = [
        f"{subject_prefix}{grade}.json",
        f"{subject_prefix}_khoi{grade}.json",
        f"{subject_prefix}_khoi_{grade}.json",
        f"{subject_prefix}-khoi{grade}.json",
        f"{subject_prefix}{grade}.JSON",
        f"{subject_prefix}_khoi{grade}.JSON",
        f"{subject_prefix}_khoi_{grade}.JSON",
        f"{subject_prefix}-khoi{grade}.JSON",
    ]
    for p in patterns:
        full = os.path.join(criteria_folder, p)
        if os.path.exists(full):
            return full
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            lower = fn.lower()
            if fn.lower().startswith(subject_prefix.lower()) and str(grade) in lower and fn.lower().endswith(".json"):
                return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
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

    path = find_criteria_file(pref, grade, criteria_folder)
    if not path:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and "tieu_chi" in data:
            for it in data["tieu_chi"]:
                if "diem" in it:
                    try:
                        it["diem"] = float(it["diem"])
                    except:
                        it["diem"] = 0.0
                else:
                    it["diem"] = 0.0
            return data
        elif isinstance(data, list):
            return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join([p.text for p in doc.paragraphs]))

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in normalize_text_no_diacritics("\n".join(p._element.xml for p in doc.paragraphs))
        elif key == "format_text":
            bold_text = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
            ok = bold_text
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes

# ---------------- PPT grading ----------------
def grade_ppt(file_path, criteria):
    import unicodedata
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    # không chia 10 nữa, lấy đúng điểm gốc
    items = criteria.get("tieu_chi", [])
    for it in items:
        try:
            it["diem"] = float(it.get("diem", 0))
        except:
            it["diem"] = 0.0

    total_awarded = 0.0
    notes = []

    num_slides = len(slides)

    # ======== HÀM HỖ TRỢ ========
    def no_accent_vn(text):
        """Bỏ dấu tiếng Việt để so sánh không phân biệt có dấu."""
        return ''.join(
            c for c in unicodedata.normalize('NFD', text)
            if unicodedata.category(c) != 'Mn'
        )

    def shape_has_picture(shape):
        """Nhận biết bất kỳ hình ảnh nào, kể cả group, SmartArt, background."""
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True
            # nếu là nhóm
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for s in shape.shapes:
                    if shape_has_picture(s):
                        return True
            # kiểm tra background có ảnh
            if hasattr(shape, "fill") and hasattr(shape.fill, "type"):
                fill = shape.fill
                if getattr(fill, "type", None) == 6:  # picture fill
                    return True
            # kiểm tra XML nếu chứa tag <p:pic>
            if "p:pic" in shape._element.xml or "a:blip" in shape._element.xml:
                return True
        except Exception:
            pass
        return False

    # ======== KIỂM TRA CÓ HÌNH ========
  # Kiểm tra có hình ảnh không (mở rộng tất cả kiểu)
def shape_has_picture(shape):
    try:
        # Kiểm tra shape là ảnh
        if shape.shape_type == 13:
            return True
        # Ảnh trong fill (dùng làm nền hoặc khung)
        if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
            return True
        # Kiểm tra trong GroupShape
        if hasattr(shape, "shapes"):
            return any(shape_has_picture(s) for s in shape.shapes)
    except Exception:
        return False
    return False

has_picture_any = any(
    shape_has_picture(shape)
    for slide in slides
    for shape in slide.shapes
)


    has_transition_any = any("transition" in slide._element.xml for slide in slides)

    # Gom tất cả text (cả có dấu và không dấu)
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                ppt_text += " " + (shape.text or "")
    ppt_text_noaccent = no_accent_vn(ppt_text.lower())

    # ======== DUYỆT TIÊU CHÍ ========
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        ori = float(it.get("diem", 0) or 0)
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val

        elif key == "has_image":
            ok = has_picture_any

        elif key == "has_transition":
            ok = has_transition_any

        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(
                    shape.has_text_frame and shape.text.strip()
                    for shape in first.shapes
                )

        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(
                no_accent_vn(t.strip().lower()) in ppt_text_noaccent
                for t in terms if t.strip()
            )

        elif key == "any":
            ok = True

        else:
            words = re.findall(r"\w+", desc.lower())
            ok = any(
                no_accent_vn(w) in ppt_text_noaccent
                for w in words if len(w) > 1
            )

        if ok:
            total_awarded += ori
            notes.append(f"✅ {desc} (+{ori})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes


# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = None
    try:
        tempdir = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tempdir)
        proj_path = os.path.join(tempdir, "project.json")
        if not os.path.exists(proj_path):
            return None, ["Không tìm thấy project.json trong sb3"]
        with open(proj_path, "r", encoding="utf-8") as f:
            proj = json.load(f)
        targets = proj.get("targets", [])
        flags = {
            "has_loop": False,
            "has_condition": False,
            "has_interaction": False,
            "has_variable": False,
            "multiple_sprites_or_animation": False
        }
        sprite_count = sum(1 for t in targets if not t.get("isStage", False))
        if sprite_count >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            if "variables" in t and len(t.get("variables", [])) > 0:
                flags["has_variable"] = True
            blocks = t.get("blocks", {})
            for block_id, block in blocks.items():
                opcode = block.get("opcode", "").lower()
                if any(k in opcode for k in ["control_repeat","control_forever","control_repeat_until"]):
                    flags["has_loop"] = True
                if any(k in opcode for k in ["control_if","control_if_else"]):
                    flags["has_condition"] = True
                if any(k in opcode for k in ["sensing_keypressed","sensing_touchingobject","event_whenthisspriteclicked",
                                             "event_whenflagclicked","event_whenbroadcastreceived","sensing_mousedown"]):
                    flags["has_interaction"] = True
                if any(k in opcode for k in ["data_setvariableto","data_changevariableby","data_hidevariable","data_showvariable"]):
                    flags["has_variable"] = True
                if "event_broadcast" in opcode:
                    flags["has_interaction"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi khi phân tích file Scratch: {e}"]
    finally:
        if tempdir and os.path.exists(tempdir):
            try:
                shutil.rmtree(tempdir)
            except Exception:
                pass

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        diem = float(it.get("diem", 0) or 0)
        ok = False
        if key == "has_loop":
            ok = flags.get("has_loop", False)
        elif key == "has_condition":
            ok = flags.get("has_condition", False)
        elif key == "has_interaction":
            ok = flags.get("has_interaction", False)
        elif key == "has_variable":
            ok = flags.get("has_variable", False)
        elif key == "multiple_sprites_or_animation":
            ok = flags.get("multiple_sprites_or_animation", False)
        elif key == "any":
            ok = True
        else:
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "phát sóng", "điều kiện", "nối"])
        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    total_awarded = round(total_awarded, 2)
    return total_awarded, notes



