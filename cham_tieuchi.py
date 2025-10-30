import os, json, zipfile, tempfile, shutil, re, unicodedata
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
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = re.sub(r'[_-]', ' ', base)
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p.capitalize() for p in s.split() if p.strip()]
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
    for fn in os.listdir(criteria_folder):
        if fn.lower().startswith(subject_prefix) and str(grade) in fn.lower():
            return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = subject.lower()
    if "power" in s: pref = "ppt"
    elif "word" in s: pref = "word"
    elif "scratch" in s: pref = "scratch"
    else: pref = s

    path = find_criteria_file(pref, grade, criteria_folder)
    if not path:
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        data = {"tieu_chi": data}
    for it in data["tieu_chi"]:
        it["diem"] = float(it.get("diem", 0) or 0)
    return data

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        desc, key, diem = it.get("mo_ta", ""), (it.get("key") or "").lower(), float(it.get("diem", 0))
        ok = False
        if key == "has_image":
            ok = bool(doc.inline_shapes) or "a:blip" in "\n".join(p._element.xml for p in doc.paragraphs)
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key.startswith("contains:"):
            ok = any(normalize_text_no_diacritics(t) in text for t in key.split(":",1)[1].split("|"))
        else:
            ok = any(w for w in re.findall(r"\w+", normalize_text_no_diacritics(desc)) if w in text)
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes

# ---------------- PPT grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    num_slides = len(slides)

    def shape_has_picture(shape):
        """Nhận diện mọi loại hình ảnh (ảnh web, nền, SmartArt, group, Bing, Chart)."""
        try:
            xml = shape._element.xml.lower()
            if any(tag in xml for tag in ["p:pic", "a:blip", "a:blipfill", "blipfill"]):
                return True
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                return any(shape_has_picture(s) for s in shape.shapes)
        except Exception:
            pass
        return False

    has_picture = any(shape_has_picture(s) for sl in slides for s in sl.shapes)
    has_transition = any("transition" in sl._element.xml for sl in slides)

    ppt_text = ""
    for sl in slides:
        for sh in sl.shapes:
            if getattr(sh, "has_text_frame", False):
                ppt_text += " " + (sh.text or "")
    text_noaccent = normalize_text_no_diacritics(ppt_text)

    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        desc, key, diem = it.get("mo_ta", ""), (it.get("key") or "").lower(), float(it.get("diem", 0))
        ok = False
        if key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "has_image":
            ok = has_picture
        elif key == "has_transition":
            ok = has_transition
        elif key == "title_first":
            ok = any(sh.has_text_frame and sh.text.strip() for sh in slides[0].shapes)
        elif key.startswith("contains:"):
            ok = any(normalize_text_no_diacritics(t.strip()) in text_noaccent for t in key.split(":",1)[1].split("|"))
        else:
            ok = any(normalize_text_no_diacritics(w) in text_noaccent for w in re.findall(r"\w+", desc))
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes

# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tmp = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path) as z: z.extractall(tmp)
        pj = json.load(open(os.path.join(tmp, "project.json"), encoding="utf-8"))
        flags = {k:False for k in ["has_loop","has_condition","has_interaction","has_variable","multiple_sprites_or_animation"]}
        targets = pj.get("targets", [])
        if sum(1 for t in targets if not t.get("isStage")) >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            bl = json.dumps(t.get("blocks", {})).lower()
            if any(k in bl for k in ["repeat","forever"]): flags["has_loop"]=True
            if any(k in bl for k in ["if","else"]): flags["has_condition"]=True
            if any(k in bl for k in ["clicked","touch","keypressed","broadcast"]): flags["has_interaction"]=True
            if "data_setvariableto" in bl or "variable" in bl: flags["has_variable"]=True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if not flags: return None, err
    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        k = (it.get("key") or "").lower()
        desc, diem = it.get("mo_ta", ""), float(it.get("diem", 0))
        ok = flags.get(k, False)
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes
