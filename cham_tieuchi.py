import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- UTILITIES ----------------
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
    for fn in os.listdir(criteria_folder):
        if fn.lower().startswith(subject_prefix.lower()) and str(grade) in fn.lower() and fn.lower().endswith(".json"):
            return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = subject.lower()
    if s in ("powerpoint", "ppt", "pptx"):
        pref = "ppt"
    elif s in ("word", "doc", "docx"):
        pref = "word"
    elif s in ("scratch", "sb3"):
        pref = "scratch"
    else:
        pref = s
    path = find_criteria_file(pref, grade, criteria_folder)
    if not path: return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and "tieu_chi" in data:
            for it in data["tieu_chi"]:
                try: it["diem"] = float(it.get("diem", 0) or 0)
                except: it["diem"] = 0.0
            return data
        elif isinstance(data, list):
            return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ---------------- WORD GRADING ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total = 0.0; notes = []

    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in normalize_text_no_diacritics(
                "\n".join(p._element.xml for p in doc.paragraphs)
            )
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")

    return round(total, 2), notes

# ---------------- POWERPOINT GRADING ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    for it in items:
        try: it["diem"] = float(it.get("diem", 0) or 0)
        except: it["diem"] = 0.0

    slides = prs.slides
    num_slides = len(slides)

    # --- Loại trừ slide trống / chỉ chữ ---
    def slide_is_text_only(slide):
        txt = ""
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                txt += shape.text or ""
        # Nếu toàn bộ slide chỉ chứa chữ, số hoặc khoảng trắng => text-only
        return txt.strip() != "" and not re.search(r"[^\w\s]", txt)

    # --- Nhận dạng ảnh toàn diện ---
    def slide_has_picture(slide):
        try:
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    return True
                if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                    return True
                if hasattr(shape, "shapes"):
                    for s in shape.shapes:
                        if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            return True
                xml = shape._element.xml.lower()
                if any(tag in xml for tag in [
                    "a:blip", "p:pic", "a:blipfill", "blipfill", "href=\"http", "image", ".jpg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp"
                ]):
                    return True
            # Kiểm tra rels chứa file ảnh
            for rel in slide.part.rels.values():
                t = getattr(rel.target_ref, "lower", lambda: "")()
                if any(ext in t for ext in [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp", ".ico", ".emf", ".wmf"]):
                    return True
            # Kiểm tra XML nền slide
            bg_xml = slide._element.xml.lower()
            if "blipfill" in bg_xml or "a:blip" in bg_xml:
                return True
        except Exception:
            pass
        return False

    # --- Phân tích tất cả slide ---
    has_picture = any(slide_has_picture(s) and not slide_is_text_only(s) for s in slides)
    has_transition = any("transition" in s._element.xml for s in slides)

    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                ppt_text += " " + (shape.text or "")
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    total = 0.0; notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val
        elif key == "has_image":
            ok = has_picture
        elif key == "has_transition":
            ok = has_transition
        elif key == "title_first" and slides:
            ok = any(getattr(s, "has_text_frame", False) and s.text.strip() for s in slides[0].shapes)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text_noaccent for w in words if len(w) > 1)

        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")

    return round(total, 2), notes

# ---------------- SCRATCH GRADING ----------------
def analyze_sb3_basic(file_path):
    tempdir = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tempdir)
        proj_path = os.path.join(tempdir, "project.json")
        if not os.path.exists(proj_path):
            return None, ["Không tìm thấy project.json"]
        with open(proj_path, "r", encoding="utf-8") as f:
            proj = json.load(f)
        targets = proj.get("targets", [])
        flags = {"has_loop": False, "has_condition": False, "has_interaction": False, "has_variable": False, "multiple_sprites_or_animation": False}
        sprite_count = sum(1 for t in targets if not t.get("isStage", False))
        if sprite_count >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            blocks = t.get("blocks", {})
            for b in blocks.values():
                op = b.get("opcode", "").lower()
                if "control_repeat" in op or "forever" in op:
                    flags["has_loop"] = True
                if "if" in op:
                    flags["has_condition"] = True
                if "when" in op or "sensing" in op:
                    flags["has_interaction"] = True
                if "data_" in op:
                    flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        shutil.rmtree(tempdir, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    total = 0.0; notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = flags.get(key, False) if key in flags else False
        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes
