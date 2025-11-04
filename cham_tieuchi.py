import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from openpyxl import Workbook, load_workbook

# ----------------- HỖ TRỢ CHUNG -----------------
def normalize_text_no_diacritics(s):
    """Chuẩn hóa: bỏ dấu, chữ thường."""
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    return " ".join(p.capitalize() for p in s.split() if p.strip())

def find_criteria_file(subject_prefix, grade, folder="criteria"):
    patterns = [
        f"{subject_prefix}{grade}.json",
        f"{subject_prefix}_khoi{grade}.json",
        f"{subject_prefix}-khoi{grade}.json",
    ]
    for p in patterns:
        path = os.path.join(folder, p)
        if os.path.exists(path):
            return path
    for fn in os.listdir(folder):
        if fn.lower().startswith(subject_prefix.lower()) and str(grade) in fn.lower() and fn.lower().endswith(".json"):
            return os.path.join(folder, fn)
    return None

def load_criteria(subject, grade, folder="criteria"):
    s = subject.lower()
    pref = "ppt" if "power" in s else "word" if "word" in s else "scratch"
    path = find_criteria_file(pref, grade, folder)
    if not path:
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "tieu_chi" in data:
        for t in data["tieu_chi"]:
            try: t["diem"] = float(t.get("diem", 0))
            except: t["diem"] = 0.0
        return data
    elif isinstance(data, list):
        return {"tieu_chi": data}
    return None

# ----------------- CHẤM WORD -----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total, notes = 0.0, []

    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in "\n".join(p._element.xml for p in doc.paragraphs).lower()
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in text for t in terms)
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in text for w in words if len(w) > 1)
        if ok: total += diem
    return round(total, 2), []

# ----------------- CHẤM POWERPOINT -----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)

    # --- Quét XML toàn bộ để nhận dạng hình ảnh (tốc độ cao, bao phủ 100%) ---
    xml_all = "\n".join(slide._element.xml.lower() for slide in slides)
    has_picture = any(tag in xml_all for tag in [
        "p:pic", "a:blip", ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff",
        ".svg", ".webp", ".ico", ".emf", ".wmf", "blipfill", "pic:blip"
    ])
    has_transition = "transition" in xml_all

    # --- Gom text một lần duy nhất ---
    ppt_text = " ".join(
        shape.text.strip()
        for slide in slides for shape in slide.shapes
        if getattr(shape, "has_text_frame", False) and shape.text.strip()
    )
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    total = 0.0
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "has_image":
            ok = has_picture
        elif key == "has_transition":
            ok = has_transition
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
            ok = any(w in ppt_text_noaccent for w in words if len(w) > 1)

        if ok: total += diem

    return round(total, 2), []

# ----------------- CHẤM SCRATCH -----------------
def analyze_sb3_basic(file_path):
    tempdir = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tempdir)
        path = os.path.join(tempdir, "project.json")
        if not os.path.exists(path):
            return None, ["Không tìm thấy project.json"]
        with open(path, "r", encoding="utf-8") as f:
            proj = json.load(f)
        targets = proj.get("targets", [])
        flags = {k: False for k in [
            "has_loop","has_condition","has_interaction","has_variable","multiple_sprites_or_animation"
        ]}
        if sum(1 for t in targets if not t.get("isStage")) >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            blocks = t.get("blocks", {})
            for b in blocks.values():
                op = b.get("opcode", "").lower()
                if "repeat" in op: flags["has_loop"] = True
                if "if" in op: flags["has_condition"] = True
                if any(k in op for k in ["keypressed","touchingobject","whenthisspriteclicked","whenflagclicked","broadcast"]):
                    flags["has_interaction"] = True
                if "variable" in op: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        shutil.rmtree(tempdir, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if not flags: return None, err
    total = 0.0
    for it in criteria.get("tieu_chi", []):
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        if flags.get(key, False) or key == "any":
            total += diem
    return round(total, 2), []

# ----------------- GHI FILE EXCEL -----------------
def save_to_excel(student, lop, subject, score, output="ket_qua.xlsx"):
    try:
        if os.path.exists(output):
            wb = load_workbook(output)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["Họ tên", "Lớp", "Môn", "Điểm"])
        ws.append([student, lop, subject, score])
        wb.save(output)
    except Exception:
        pass
