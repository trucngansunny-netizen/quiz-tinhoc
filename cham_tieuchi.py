import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ========== UTILITIES ==========
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
    parts = [p.capitalize() for p in s.split() if p.strip()]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    for fn in os.listdir(criteria_folder):
        low = fn.lower()
        if low.startswith(subject_prefix) and str(grade) in low and low.endswith(".json"):
            return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
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
    return None

# ========== WORD GRADING ==========
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc Word: {e}"]
    text = normalize_text_no_diacritics("\n".join([p.text for p in doc.paragraphs]))
    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False
        if key == "has_image":
            ok = bool(doc.inline_shapes) or "a:blip" in "\n".join(p._element.xml for p in doc.paragraphs)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in text for w in words if len(w) > 1)
        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total, 2), notes

# ========== PPT GRADING ==========
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc PowerPoint: {e}"]

    # -------- NHẬN DẠNG HÌNH 100% (BẰNG CÁCH ĐỌC TRỰC TIẾP FILE .PPTX) --------
    has_picture_any = False
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            # tất cả phần mở rộng hình ảnh phổ biến
            image_exts = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff',
                          '.svg', '.webp', '.ico', '.emf', '.wmf', '.jfif')
            for name in z.namelist():
                if name.lower().startswith('ppt/media/') and name.lower().endswith(image_exts):
                    has_picture_any = True
                    break
    except Exception:
        has_picture_any = False

    # -------- THÔNG TIN KHÁC --------
    has_transition_any = any("transition" in slide._element.xml for slide in slides)
    num_slides = len(slides)
    ppt_text = " ".join(
        shape.text for s in slides for shape in s.shapes if getattr(shape, "has_text_frame", False)
    ).lower()
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "has_image":
            ok = has_picture_any
        elif key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "has_transition":
            ok = has_transition_any
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in ppt_text_noaccent for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text_noaccent for w in words if len(w) > 1)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes

# ========== SCRATCH ==========
def analyze_sb3_basic(file_path):
    tmp = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path, 'r') as z: z.extractall(tmp)
        with open(os.path.join(tmp, "project.json"), encoding="utf-8") as f:
            proj = json.load(f)
        targets = proj.get("targets", [])
        flags = {k: False for k in ["has_loop","has_condition","has_interaction","has_variable","multiple_sprites_or_animation"]}
        if sum(1 for t in targets if not t.get("isStage")) >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            blocks = t.get("blocks", {})
            for b in blocks.values():
                op = b.get("opcode", "")
                if "repeat" in op: flags["has_loop"] = True
                if "if" in op: flags["has_condition"] = True
                if "event_when" in op or "sensing_" in op: flags["has_interaction"] = True
                if "data_" in op: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi Scratch: {e}"]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        desc, key = it.get("mo_ta",""), (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = flags.get(key, False)
        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total, 2), notes
