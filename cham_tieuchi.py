# cham_tieuchi.py
# Hàm chấm: grade_word, grade_ppt, grade_scratch
# Không scale về 10 — dùng trực tiếp điểm trong tiêu chí

import os
import json
import zipfile
import tempfile
import shutil
import re
import unicodedata

from docx import Document
from pptx import Presentation

# for shape type constants (but we mainly inspect xml)
try:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except Exception:
    MSO_SHAPE_TYPE = None


# ---------------- utilities ----------------
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
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)


def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    # try common patterns you used
    candidates = [
        f"{subject_prefix}{grade}.json",
        f"{subject_prefix}_khoi{grade}.json",
        f"{subject_prefix}_khoi_{grade}.json",
        f"{subject_prefix}-khoi{grade}.json",
        f"{subject_prefix}{grade}.JSON",
    ]
    for c in candidates:
        p = os.path.join(criteria_folder, c)
        if os.path.exists(p):
            return p
    # scan folder for best match
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            low = fn.lower()
            if low.endswith(".json") and low.startswith(subject_prefix.lower()) and str(grade) in low:
                return os.path.join(criteria_folder, fn)
    return None


def load_criteria(subject, grade, criteria_folder="criteria"):
    # subject can be "word"/"ppt"/"scratch" or "Word"/"PowerPoint"
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
    except Exception:
        return None
    # normalize
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
    if isinstance(data, list):
        return {"tieu_chi": data}
    return None


# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    raw_text = "\n".join(p.text for p in doc.paragraphs)
    text = normalize_text_no_diacritics(raw_text)

    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten", "họ tên"])
        elif key == "has_image":
            # detect inline shapes or xml containing graphicData/a:blip/p:pic
            try:
                ok = bool(doc.inline_shapes) or any("graphicdata" in p._element.xml.lower() or "a:blip" in p._element.xml.lower() for p in doc.paragraphs)
            except Exception:
                ok = bool(doc.inline_shapes)
        elif key == "format_text":
            bold = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
            ok = bold
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in text for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes


# ---------------- PPT grading ----------------
def _shape_has_picture(shape):
    """Robust check: xml tags, shape type, fill type, group shapes."""
    try:
        # shape type check if available
        if MSO_SHAPE_TYPE is not None:
            try:
                if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                    return True
                if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                    # group: check children
                    for s in shape.shapes:
                        if _shape_has_picture(s):
                            return True
            except Exception:
                pass
        # check fill type if present (picture fill)
        try:
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) is not None:
                # compare xml: look for blip (image) tags
                xml = getattr(shape, "_element", None)
                if xml is not None:
                    xmls = shape._element.xml.lower()
                    if any(tag in xmls for tag in ("p:pic", "a:blip", "blipfill", "a:blipfill")):
                        return True
        except Exception:
            pass
        # check raw xml for any image markers
        try:
            xmls = shape._element.xml.lower()
            if any(tag in xmls for tag in ("p:pic", "a:blip", "blipfill", "a:blipfill")):
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    # ensure numeric diem
    for it in items:
        try:
            it["diem"] = float(it.get("diem", 0))
        except:
            it["diem"] = 0.0

    num_slides = len(slides)

    # detect pictures robustly across all slides/shapes
    has_picture = False
    for slide in slides:
        try:
            for shape in slide.shapes:
                if _shape_has_picture(shape):
                    has_picture = True
                    break
            if has_picture:
                break
        except Exception:
            continue

    # transitions: check if any slide xml mentions transition
    has_transition = any("transition" in getattr(slide._element, "xml", "").lower() for slide in slides)

    # aggregate text (no-diacritics)
    ppt_text = []
    for slide in slides:
        for shape in slide.shapes:
            try:
                if getattr(shape, "has_text_frame", False):
                    ppt_text.append(shape.text or "")
            except Exception:
                continue
    ppt_text_all = normalize_text_no_diacritics(" ".join(ppt_text))

    total = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
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
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(getattr(s, "has_text_frame", False) and (s.text or "").strip() for s in first.shapes)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_all for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            # fallback: look for words in description
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in ppt_text_all)

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes


# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tmp = None
    try:
        tmp = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(file_path, "r") as z:
            z.extractall(tmp)
        proj = os.path.join(tmp, "project.json")
        if not os.path.exists(proj):
            return None, ["Không tìm thấy project.json trong sb3"]
        with open(proj, "r", encoding="utf-8") as f:
            data = json.load(f)
        targets = data.get("targets", [])
        flags = {
            "has_loop": False,
            "has_condition": False,
            "has_interaction": False,
            "has_variable": False,
            "multiple_sprites_or_animation": False,
        }
        sprite_count = sum(1 for t in targets if not t.get("isStage", False))
        if sprite_count >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            if "variables" in t and len(t.get("variables", []) or []) > 0:
                flags["has_variable"] = True
            blocks = t.get("blocks", {}) or {}
            for bid, block in blocks.items():
                opcode = (block.get("opcode") or "").lower()
                if any(k in opcode for k in ("control_repeat", "control_forever", "control_repeat_until")):
                    flags["has_loop"] = True
                if any(k in opcode for k in ("control_if", "control_if_else")):
                    flags["has_condition"] = True
                if any(k in opcode for k in ("sensing_keypressed", "sensing_touchingobject", "event_whenthisspriteclicked", "event_whenflagclicked", "event_whenbroadcastreceived", "sensing_mousedown")):
                    flags["has_interaction"] = True
                if any(k in opcode for k in ("data_setvariableto", "data_changevariableby", "data_hidevariable", "data_showvariable")):
                    flags["has_variable"] = True
                if "event_broadcast" in opcode:
                    flags["has_interaction"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi phân tích sb3: {e}"]
    finally:
        if tmp and os.path.exists(tmp):
            try:
                shutil.rmtree(tmp)
            except Exception:
                pass


def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err

    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
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
        elif key in ("multiple_sprites_or_animation",):
            ok = flags.get("multiple_sprites_or_animation", False)
        elif key == "any":
            ok = True
        else:
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "broadcast", "phát sóng", "điều kiện"])
        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    total = round(total, 2)
    return total, notes
