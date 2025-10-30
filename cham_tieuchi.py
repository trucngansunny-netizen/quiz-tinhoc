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
    """Tìm file tiêu chí chấm tương ứng"""
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
    # fallback quét thư mục
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            lower = fn.lower()
            if subject_prefix.lower() in lower and str(grade) in lower and lower.endswith(".json"):
                return os.path.join(criteria_folder, fn)
    return None


def load_criteria(subject, grade, criteria_folder="criteria"):
    """Đọc tiêu chí từ file JSON"""
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
                try:
                    it["diem"] = float(it.get("diem", 0))
                except:
                    it["diem"] = 0.0
            return data
        elif isinstance(data, list):
            return {"tieu_chi": data}
    except Exception:
        return None
    return None


# ---------------- WORD grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join([p.text for p in doc.paragraphs]))
    total_awarded = 0.0
    notes = []

    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in "\n".join(p._element.xml for p in doc.paragraphs)
        elif key == "format_text":
            bold_text = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
            ok = bold_text
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            ok = any(w for w in normalize_text_no_diacritics(desc).split() if w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes


# ---------------- POWERPOINT grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)
    total_awarded = 0.0
    notes = []

    def shape_has_picture(shape):
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                return any(shape_has_picture(s) for s in shape.shapes)
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            xml_str = shape._element.xml.lower()
            if any(tag in xml_str for tag in ["p:pic", "a:blip", "a:blipfill", "svgblip", "blipfill", "r:link", "http"]):
                return True
        except Exception:
            pass
        return False

    has_picture_any = any(shape_has_picture(shape) for slide in slides for shape in slide.shapes)
    has_transition_any = any("transition" in slide._element.xml for slide in slides)
    ppt_text = " ".join(shape.text for slide in slides for shape in slide.shapes if getattr(shape, "has_text_frame", False))
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    for it in criteria.get("tieu_chi", []):
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
            ok = any(shape.has_text_frame and shape.text.strip() for shape in slides[0].shapes) if slides else False
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms if t.strip())
        else:
            ok = any(w in ppt_text_noaccent for w in normalize_text_no_diacritics(desc).split() if len(w) > 1)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes


# ---------------- SCRATCH grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = None
    try:
        tempdir = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(file_path, "r") as z:
            z.extractall(tempdir)
        with open(os.path.join(tempdir, "project.json"), "r", encoding="utf-8") as f:
            proj = json.load(f)
        targets = proj.get("targets", [])
        flags = {
            "has_loop": False, "has_condition": False,
            "has_interaction": False, "has_variable": False,
            "multiple_sprites_or_animation": False
        }
        sprite_count = sum(1 for t in targets if not t.get("isStage", False))
        if sprite_count >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            blocks = t.get("blocks", {})
            for block in blocks.values():
                op = block.get("opcode", "").lower()
                if "repeat" in op or "forever" in op:
                    flags["has_loop"] = True
                if "if" in op:
                    flags["has_condition"] = True
                if any(k in op for k in ["sensing_keypressed", "event_whenflagclicked", "event_whenthisspriteclicked"]):
                    flags["has_interaction"] = True
                if "data_" in op:
                    flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi phân tích Scratch: {e}"]
    finally:
        if tempdir and os.path.exists(tempdir):
            shutil.rmtree(tempdir, ignore_errors=True)


def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err

    total_awarded = 0.0
    notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = flags.get(key, False) if key in flags else False
        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total_awarded, 2), notes
