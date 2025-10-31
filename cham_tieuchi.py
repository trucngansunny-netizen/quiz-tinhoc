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
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s


def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
    s = re.sub(r"(\d+)", r" \1 ", s)
    parts = [p.capitalize() for p in s.split() if p.strip()]
    return " ".join(parts)


def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    for fn in os.listdir(criteria_folder):
        if fn.lower().startswith(subject_prefix.lower()) and str(grade) in fn.lower() and fn.lower().endswith(".json"):
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
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return {"tieu_chi": data}
        if isinstance(data, dict):
            return data
    except:
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
    total_awarded, notes = 0.0, []

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_image":
            ok = bool(doc.inline_shapes) or "graphicdata" in "\n".join(p._element.xml for p in doc.paragraphs)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        elif key == "any":
            ok = True
        else:
            ok = normalize_text_no_diacritics(desc) in text

        total_awarded += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")

    return round(total_awarded, 2), notes


# ---------------- PowerPoint grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    total_awarded, notes = 0.0, []

    # -------- Hàm dò hình ảnh mạnh hơn --------
    def detect_any_image_in_pptx(path):
        try:
            with zipfile.ZipFile(path, "r") as z:
                for name in z.namelist():
                    if "media/" in name.lower() and name.lower().endswith((
                        ".png", ".jpg", ".jpeg", ".gif", ".bmp",
                        ".webp", ".svg", ".ico", ".tif", ".tiff", ".heic"
                    )):
                        return True
                    # Cũng kiểm tra ảnh nền trong slide.xml
                    if name.endswith(".xml"):
                        content = z.read(name).decode("utf-8", errors="ignore").lower()
                        if any(tag in content for tag in ["a:blip", "p:pic", "r:embed", "data:image/"]):
                            return True
        except Exception:
            pass
        return False

    has_picture_any = detect_any_image_in_pptx(file_path)
    has_transition_any = any("transition" in slide._element.xml for slide in slides)

    ppt_text = " ".join(
        shape.text for slide in slides for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    )
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_image":
            ok = has_picture_any
        elif key == "has_transition":
            ok = has_transition_any
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms)
        elif key == "any":
            ok = True
        else:
            ok = normalize_text_no_diacritics(desc) in ppt_text_noaccent

        total_awarded += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")

    return round(total_awarded, 2), notes


# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(file_path, "r") as z:
            z.extractall(tempdir)
        proj_path = os.path.join(tempdir, "project.json")
        if not os.path.exists(proj_path):
            return None, ["Không tìm thấy project.json"]
        with open(proj_path, "r", encoding="utf-8") as f:
            proj = json.load(f)
        flags = {"has_loop": False, "has_condition": False, "has_interaction": False, "has_variable": False}
        for t in proj.get("targets", []):
            for block in t.get("blocks", {}).values():
                opcode = block.get("opcode", "").lower()
                if "repeat" in opcode: flags["has_loop"] = True
                if "if" in opcode: flags["has_condition"] = True
                if "touching" in opcode or "keypressed" in opcode: flags["has_interaction"] = True
                if "variable" in opcode: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        shutil.rmtree(tempdir, ignore_errors=True)


def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = key in flags and flags[key]
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes
