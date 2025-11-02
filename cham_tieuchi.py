import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
    """Chuẩn hóa chuỗi: bỏ dấu tiếng Việt, viết thường."""
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
            if fn.startswith(subject_prefix.lower()) and str(grade) in lower and fn.endswith(".json"):
                return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """Đọc file tiêu chí tương ứng với môn và khối."""
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

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in normalize_text_no_diacritics("\n".join(p._element.xml for p in doc.paragraphs))
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total_awarded, 2), notes

# ---------------- PowerPoint grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)
    items = criteria.get("tieu_chi", [])
    total_awarded, notes = 0.0, []

    def shape_has_picture(shape):
        try:
            # Ảnh chèn trực tiếp
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True
            # Ảnh nền
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            # Ảnh trong nhóm
            if hasattr(shape, "shapes"):
                return any(shape_has_picture(s) for s in shape.shapes)
            # Ảnh trong SmartArt, biểu đồ, bảng
            xml = shape._element.xml.lower()
            if any(tag in xml for tag in ["p:pic", "a:blip", "a:blipfill", "blipfill"]):
                return True
            return False
        except Exception:
            return False

    has_picture_any = any(shape_has_picture(s) for sl in slides for s in sl.shapes)
    has_transition_any = any("transition" in sl._element.xml for sl in slides)
    ppt_text = " ".join(s.text for sl in slides for s in sl.shapes if getattr(s, "has_text_frame", False))
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val
        elif key == "has_image":
            ok = has_picture_any
        elif key == "has_transition":
            ok = has_transition_any
        elif key == "title_first":
            ok = bool(slides and any(s.has_text_frame and s.text.strip() for s in slides[0].shapes))
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text_noaccent for w in words if len(w) > 1)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes

# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tempdir)
        pj = os.path.join(tempdir, "project.json")
        if not os.path.exists(pj):
            return None, ["Không tìm thấy project.json"]
        with open(pj, "r", encoding="utf-8") as f:
            proj = json.load(f)
        flags = {"has_loop": False, "has_condition": False, "has_interaction": False,
                 "has_variable": False, "multiple_sprites_or_animation": False}
        targets = proj.get("targets", [])
        sprite_count = sum(1 for t in targets if not t.get("isStage", False))
        if sprite_count >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            blocks = t.get("blocks", {})
            if "variables" in t and len(t["variables"]) > 0:
                flags["has_variable"] = True
            for b in blocks.values():
                op = b.get("opcode", "").lower()
                if "control_repeat" in op or "forever" in op: flags["has_loop"] = True
                if "control_if" in op: flags["has_condition"] = True
                if "sensing_" in op or "event_" in op: flags["has_interaction"] = True
                if "data_" in op: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi Scratch: {e}"]
    finally:
        shutil.rmtree(tempdir, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    total_awarded = 0.0
    notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = flags.get(key, False) if key in flags else any(k in desc.lower() for k in ["vòng", "lặp", "biến", "phát sóng"])
        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total_awarded, 2), notes
