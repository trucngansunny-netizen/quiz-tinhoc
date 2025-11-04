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

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
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
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    """
    Try a set of filename patterns and then fallback to scanning the folder
    for something that matches subject prefix and grade.
    """
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    try:
        files = os.listdir(criteria_folder)
    except Exception:
        return None

    patterns = [
        f"{subject_prefix}{grade}.json",
        f"{subject_prefix}_khoi{grade}.json",
        f"{subject_prefix}_khoi_{grade}.json",
        f"{subject_prefix}-khoi{grade}.json",
        f"{subject_prefix}_khoi{grade}.JSON",
        f"{subject_prefix}{grade}.JSON"
    ]
    # exact patterns first
    for p in patterns:
        candidate = os.path.join(criteria_folder, p)
        if os.path.exists(candidate):
            return candidate

    # fallback: case-insensitive scan, accept variants containing both prefix and grade
    lowpref = subject_prefix.lower()
    sgrade = str(grade)
    for fn in files:
        low = fn.lower()
        if low.endswith(".json") and lowpref in low and sgrade in low:
            return os.path.join(criteria_folder, fn)

    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    signature: load_criteria(subject, grade, criteria_folder="criteria")
    subject can be "word", "ppt", "powerpoint", "scratch" (case-insensitive)
    grade can be int or str (3/4/5)
    returns dict with key 'tieu_chi' or None
    """
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = (subject or "").lower()
    if s in ("powerpoint", "ppt", "pptx"):
        pref = "ppt"
    elif s in ("word", "doc", "docx"):
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
    except Exception:
        return None

    # Normalize structure: ensure dict with "tieu_chi"
    if isinstance(data, dict) and "tieu_chi" in data:
        for it in data["tieu_chi"]:
            try:
                it["diem"] = float(it.get("diem", 0) or 0)
            except Exception:
                it["diem"] = 0.0
        return data
    elif isinstance(data, list):
        # allow JSON that's directly a list of criteria
        for it in data:
            try:
                it["diem"] = float(it.get("diem", 0) or 0)
            except Exception:
                it["diem"] = 0.0
        return {"tieu_chi": data}
    return None

# ---------------- Word grading ----------------
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
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image":
            # use inline_shapes (fast) and element xml fallback
            ok = bool(getattr(doc, "inline_shapes", None) and len(doc.inline_shapes) > 0) \
                 or any("graphicdata" in p._element.xml.lower() for p in doc.paragraphs)
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if getattr(run, "bold", False))
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t).strip() in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes

# ---------------- PPT grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)

    # helper: detect picture in a shape quickly and robustly
    def shape_has_picture(shape):
        try:
            st = getattr(shape, "shape_type", None)
            # direct picture
            if st == MSO_SHAPE_TYPE.PICTURE:
                return True
            # group -> check children
            if st == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
                for s in shape.shapes:
                    if shape_has_picture(s):
                        return True
            # fill picture (background or shape fill)
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            # xml scan for blip/pic tags (fast)
            xml = getattr(shape._element, "xml", "").lower()
            if any(tag in xml for tag in ["p:pic", "a:blip", "blipfill", "a:blipfill", "href=\"http"]):
                return True
            # sometimes picture is embedded in chart/smartart as blip in xml
            if "blip" in xml or ".jpg" in xml or ".png" in xml or ".gif" in xml:
                return True
        except Exception:
            return False
        return False

    # fast detection: stop early on first found picture
    has_picture = False
    for sl in slides:
        for sh in sl.shapes:
            if shape_has_picture(sh):
                has_picture = True
                break
        if has_picture:
            break

    # transitions (simple xml search)
    has_transition = any("transition" in getattr(sl._element, "xml", "").lower() for sl in slides)

    # gather text for contains checks
    ppt_text = []
    for sl in slides:
        for sh in sl.shapes:
            if getattr(sh, "has_text_frame", False):
                try:
                    ppt_text.append(sh.text or "")
                except Exception:
                    pass
    ppt_text = " ".join(ppt_text)
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text)

    total_awarded = 0.0
    notes = []

    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "min_slides":
            try:
                req = int(it.get("value", 1))
            except Exception:
                req = 1
            ok = num_slides >= req
        elif key == "has_image":
            if has_picture:
                ok = True
            else:
                # fallback: check any slide has non-text-only content by quick heuristic
                for sl in slides:
                    slide_has_nontext = False
                    for sh in sl.shapes:
                        if getattr(sh, "has_text_frame", False):
                            # if text only and small punctuation, skip
                            txt = normalize_text_no_diacritics(getattr(sh, "text", "") or "").strip()
                            # if txt contains letters/digits it's text; we treat text-only as non-picture
                            # if shape has no text_frame (e.g., table/chart) -> count as non-text
                            if txt == "" or not re.fullmatch(r"[a-z0-9\s.,!?;:\-()]*", txt):
                                slide_has_nontext = True
                                break
                        else:
                            # shape without text_frame -> likely image/chart/shape
                            slide_has_nontext = True
                            break
                    if slide_has_nontext:
                        ok = True
                        break
        elif key == "has_transition":
            ok = has_transition
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(getattr(s, "has_text_frame", False) and (s.text or "").strip() for s in first.shapes)
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
    tempdir = None
    try:
        tempdir = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(file_path, "r") as z:
            z.extractall(tempdir)
        proj_path = os.path.join(tempdir, "project.json")
        if not os.path.exists(proj_path):
            return None, ["Không tìm thấy project.json"]
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
            for block in blocks.values():
                opcode = (block.get("opcode") or "").lower()
                if any(k in opcode for k in ["control_repeat", "control_forever", "control_repeat_until"]):
                    flags["has_loop"] = True
                if any(k in opcode for k in ["control_if", "control_if_else"]):
                    flags["has_condition"] = True
                if any(k in opcode for k in ["sensing_keypressed", "sensing_touchingobject", "event_when"]):
                    flags["has_interaction"] = True
                if any(k in opcode for k in ["data_setvariableto", "data_changevariableby"]):
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
    total_awarded = 0.0
    notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = flags.get(key, False) if key in flags else (key == "any")
        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total_awarded, 2), notes
