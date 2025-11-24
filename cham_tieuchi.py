# cham_tieuchi.py
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

# ===================== HÀM TIỆN ÍCH =====================
def normalize_text_no_diacritics(s):
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
    """
    subject: 'word'/'powerpoint'/'scratch' etc
    grade: int or str
    returns dict with key 'tieu_chi' (list of {mo_ta, diem, key?})
    """
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
        # normalize diem to float
        if isinstance(data, dict) and "tieu_chi" in data:
            for it in data["tieu_chi"]:
                it["diem"] = float(it.get("diem", 0) or 0)
            return data
        elif isinstance(data, list):
            return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ===================== WORD =====================
def grade_word(file_path, criteria, filename=None, grade=None):
    """
    Returns: total_awarded (float), notes (list of "✅ ..."/"❌ ...")
    """
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
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        # Heuristics for common keys / descriptions
        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes)
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t).strip() in text for t in terms)
        else:
            # fallback: look for keywords in description
            desc_norm = normalize_text_no_diacritics(desc)
            # special: "doi mau", "doi co", "doi phong" -> search for 'mau' 'co' 'phong'
            if "doi mau" in desc_norm or "mau" in desc_norm and "phong" in desc_norm:
                ok = ("mau" in desc_norm and ("mau" in text or "color" in text))
            else:
                words = re.findall(r"\w+", desc_norm)
                ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} ({diem})")
        else:
            notes.append(f"❌ {desc} (0/{diem})")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes

# ===================== POWERPOINT =====================
def grade_ppt(file_path, criteria, filename=None, grade=None):
    """
    Returns: total_awarded (float), notes list
    """
    try:
        prs = Presentation(file_path)
        slides = list(prs.slides)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []
    num_slides = len(slides)

    # helper: detect image anywhere on slides
    def slide_has_picture(slide):
        for shape in slide.shapes:
            try:
                if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                    return True
                xml_str = getattr(shape, "_element", None)
                if xml_str is not None:
                    xml_s = shape._element.xml.lower()
                    if any(tag in xml_s for tag in ["p:pic", "a:blip", ".jpg", ".png", ".gif", ".jpeg", "blip"]):
                        return True
            except Exception:
                continue
        return False

    has_image = any(slide_has_picture(s) for s in slides)
    has_transition = any("transition" in s._element.xml.lower() for s in slides if getattr(s, "_element", None))
    ppt_text = " ".join(shape.text for slide in slides for shape in slide.shapes if getattr(shape, "has_text_frame", False))
    ppt_text_norm = normalize_text_no_diacritics(ppt_text)

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        # Heuristics:
        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val
        elif "3 trang" in desc.lower() or re.search(r"\b3\b.*trang", desc.lower()):
            ok = num_slides >= 3
        elif "trang 1" in desc.lower() or "trang 2" in desc.lower() or "trang 3" in desc.lower():
            # check specific slide has some text
            if "trang 1" in desc.lower() and num_slides >= 1:
                first = slides[0]
                ok = any(getattr(shape, "has_text_frame", False) and shape.text.strip() for shape in first.shapes)
            elif "trang 2" in desc.lower() and num_slides >= 2:
                second = slides[1]
                ok = any(getattr(shape, "has_text_frame", False) and shape.text.strip() for shape in second.shapes)
            elif "trang 3" in desc.lower() and num_slides >= 3:
                third = slides[2]
                ok = any(getattr(shape, "has_text_frame", False) and shape.text.strip() for shape in third.shapes)
        elif key == "has_image" or "chen hinh" in desc.lower() or "chèn hình" in desc.lower():
            ok = has_image
        elif key == "has_transition" or "chuyển trang" in desc.lower() or "hieu ung" in desc.lower():
            ok = has_transition
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_norm for t in terms if t.strip())
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(getattr(shape, "has_text_frame", False) and shape.text.strip() for shape in first.shapes)
        elif key == "any":
            ok = True
        else:
            # fallback: match some words from description
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text_norm for w in words if len(w) > 1)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} ({diem})")
        else:
            notes.append(f"❌ {desc} (0/{diem})")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes

# ===================== SCRATCH =====================
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
            for _, block in blocks.items():
                opcode = (block.get("opcode") or "").lower()
                if any(k in opcode for k in ["control_repeat","control_forever","control_repeat_until"]):
                    flags["has_loop"] = True
                if any(k in opcode for k in ["control_if","control_if_else"]):
                    flags["has_condition"] = True
                if any(k in opcode for k in ["sensing_keypressed","sensing_touchingobject",
                                             "event_whenthisspriteclicked","event_whenflagclicked",
                                             "event_whenbroadcastreceived","sensing_mousedown"]):
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
            shutil.rmtree(tempdir, ignore_errors=True)

def grade_scratch(file_path, criteria, filename=None, grade=None):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False
        if key:
            ok = bool(flags.get(key, False))
        else:
            # fallback: if desc contains keywords
            desc_norm = normalize_text_no_diacritics(desc)
            if "lenh doi" in desc_norm or "đợi" in desc_norm or "doi" in desc_norm:
                ok = flags.get("has_loop", False) or flags.get("has_condition", False)
            elif "hiển thị" in desc_norm or "hien thi" in desc_norm:
                ok = flags.get("has_interaction", False)
            elif "bien" in desc_norm or "variable" in desc_norm:
                ok = flags.get("has_variable", False)
            elif "sprite" in desc_norm or "animation" in desc_norm:
                ok = flags.get("multiple_sprites_or_animation", False)
            else:
                ok = any(flags.values())

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} ({diem})")
        else:
            notes.append(f"❌ {desc} (0/{diem})")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes
