# cham_tieuchi.py
# Phiên bản chuẩn: chấm Word / PPT / Scratch theo file tiêu chí JSON
# - Không scale về 10 (lấy điểm gốc trong JSON)
# - Nhận diện ảnh mở rộng (picture, fill, group, smartart, chart...)
# - Hỗ trợ kiểm tra nội dung không dấu (so sánh không phân biệt dấu)
# - load_criteria tìm nhiều biến thể tên file trong folder criteria

import os
import jsonimport os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- UTILITIES ----------------
def normalize_text_no_diacritics(s):
    """Chuẩn hóa: bỏ dấu tiếng Việt, viết thường."""
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
    s = re.sub(r"(\d+)", r" \1 ", s)
    parts = [p.capitalize() for p in s.split() if p.strip()]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    """Tìm file tiêu chí phù hợp"""
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
        lower = fn.lower()
        if lower.startswith(subject_prefix) and str(grade) in lower and lower.endswith(".json"):
            return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """Nạp tiêu chí JSON"""
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = subject.lower()
    if s in ("powerpoint", "ppt", "pptx"):
        pref = "ppt"
    elif s in ("word", "docx"):
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
            data = {"tieu_chi": data}
        for it in data.get("tieu_chi", []):
            try:
                it["diem"] = float(it.get("diem", 0))
            except:
                it["diem"] = 0.0
        return data
    except Exception:
        return None

# ---------------- WORD ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total = 0.0
    notes = []

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
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            ok = any(w in text for w in re.findall(r"\w+", normalize_text_no_diacritics(desc)))

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total, 2), notes

# ---------------- POWERPOINT ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    total = 0.0
    notes = []
    num_slides = len(slides)

    def shape_has_picture(shape):
        """Phát hiện mọi loại ảnh"""
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                return any(shape_has_picture(s) for s in shape.shapes)
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            xml = shape._element.xml.lower()
            if any(tag in xml for tag in ["p:pic", "a:blip", "a:blipfill", "blipfill"]):
                return True
        except Exception:
            pass
        return False

    has_image = any(shape_has_picture(s) for slide in slides for s in slide.shapes)
    has_transition = any("transition" in slide._element.xml for slide in slides)
    ppt_text = normalize_text_no_diacritics(
        " ".join(s.text for slide in slides for s in slide.shapes if getattr(s, "has_text_frame", False))
    )

    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val
        elif key == "has_image":
            ok = has_image
        elif key == "has_transition":
            ok = has_transition
        elif key == "title_first":
            ok = slides and any(s.text.strip() for s in slides[0].shapes if getattr(s, "has_text_frame", False))
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in ppt_text for t in terms)
        else:
            ok = any(w in ppt_text for w in re.findall(r"\w+", normalize_text_no_diacritics(desc)))

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total, 2), notes

# ---------------- SCRATCH ----------------
def analyze_sb3_basic(file_path):
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tmp)
        proj = json.load(open(os.path.join(tmp, "project.json"), encoding="utf-8"))
        flags = {k: False for k in ["has_loop", "has_condition", "has_interaction", "has_variable", "multiple_sprites_or_animation"]}
        sprites = [t for t in proj.get("targets", []) if not t.get("isStage")]
        if len(sprites) >= 2: flags["multiple_sprites_or_animation"] = True
        for t in sprites:
            blocks = t.get("blocks", {})
            for b in blocks.values():
                op = b.get("opcode", "").lower()
                if any(x in op for x in ["control_repeat", "forever"]): flags["has_loop"] = True
                if "if" in op: flags["has_condition"] = True
                if "broadcast" in op or "sensing" in op: flags["has_interaction"] = True
                if "variable" in op or "data_" in op: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None: return None, err
    total, notes = 0.0, []
    for it in criteria.get("tieu_chi", []):
        desc, key = it.get("mo_ta", ""), (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = key == "any" or flags.get(key, False)
        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total, 2), notes

import zipfile
import tempfile
import shutil
import re
import unicodedata

from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- Utilities ----------------
def _no_accent(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s

def pretty_name_from_filename(filename: str) -> str:
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)

def _find_criteria_file(prefix: str, grade, folder="criteria"):
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    names = [
        f"{prefix}{grade}.json",
        f"{prefix}_khoi{grade}.json",
        f"{prefix}_khoi_{grade}.json",
        f"{prefix}-khoi{grade}.json",
        f"{prefix}{grade}.JSON",
        f"{prefix}_khoi{grade}.JSON",
        f"{prefix}_khoi_{grade}.JSON",
    ]
    for n in names:
        p = os.path.join(folder, n)
        if os.path.exists(p):
            return p
    # fallback: scan folder for file starting with prefix and containing grade
    if os.path.isdir(folder):
        for fn in os.listdir(folder):
            low = fn.lower()
            if low.endswith(".json") and low.startswith(prefix.lower()) and str(grade) in low:
                return os.path.join(folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    subject: 'word'/'ppt'/'powerpoint'/'scratch' (case-insensitive)
    grade: 3/4/5 or string
    returns dict: {"tieu_chi": [ {mo_ta, diem, key?, value?}, ... ] } or None
    """
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    s = (subject or "").lower()
    if s in ("powerpoint", "ppt", "pptx"):
        pref = "ppt"
    elif s in ("word", "docx", "doc"):
        pref = "word"
    elif s in ("scratch", "sb3"):
        pref = "scratch"
    else:
        pref = s

    path = _find_criteria_file(pref, grade, criteria_folder)
    if not path:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return None

    # Normalize
    if isinstance(data, dict) and "tieu_chi" in data:
        items = data["tieu_chi"]
    elif isinstance(data, list):
        items = data
    else:
        return None

    # ensure diem numeric
    for it in items:
        if "diem" in it:
            try:
                it["diem"] = float(it["diem"])
            except:
                it["diem"] = 0.0
        else:
            it["diem"] = 0.0
        # ensure key exists
        if "key" in it and isinstance(it["key"], str):
            it["key"] = it["key"].strip()
    return {"tieu_chi": items}

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    # gather textual content
    text = "\n".join(p.text for p in doc.paragraphs)
    text_norm = _no_accent(text)

    # detect images in doc: try inline_shapes, and fallback to xml check
    has_image = False
    try:
        has_image = bool(getattr(doc, "inline_shapes", None) and len(doc.inline_shapes) > 0)
    except Exception:
        has_image = False
    if not has_image:
        try:
            # check underlying xml of paragraphs for graphicData or pict tags
            xml_blob = "\n".join(p._element.xml for p in doc.paragraphs)
            if "graphicData" in xml_blob.lower() or "<p:pic" in xml_blob.lower() or "a:blip" in xml_blob.lower():
                has_image = True
        except Exception:
            pass

    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        val = it.get("value", None)
        pts = float(it.get("diem", 0) or 0)
        ok = False

        try:
            if key == "has_title":
                ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
            elif key == "has_name":
                ok = bool(re.search(r"\b(họ và tên|họ tên|họ tên:|tên học sinh|ho ten|ten hoc sinh)\b", text_norm))
            elif key == "has_image":
                ok = has_image
            elif key == "format_text":
                # check bold/italic somewhere
                bold_found = any(run.bold for p in doc.paragraphs for run in p.runs if getattr(run, "bold", False))
                ok = bold_found
            elif key.startswith("contains:"):
                # contains:term1|term2
                terms = key.split(":",1)[1].split("|")
                ok = any(_no_accent(t.strip()) in text_norm for t in terms if t.strip())
            elif key == "any" or key == "":
                ok = True
            else:
                # fallback: find important words from description
                words = re.findall(r"\w+", _no_accent(desc))
                ok = any(w for w in words if len(w) > 1 and w in text_norm)
        except Exception:
            ok = False

        if ok:
            total += pts
            notes.append(f"✅ {desc} (+{pts})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes

# ---------------- PPT grading ----------------
def _shape_has_picture(shape):
    """
    Robust detection of picture in a shape:
    - shape_type == PICTURE
    - group shapes (recursively)
    - fill picture (type==6)
    - SmartArt/Chart/Table etc containing <a:blip> or <p:pic>
    - XML contains 'a:blip' or 'p:pic' or 'blipfill' or 'a:blipfill'
    """
    try:
        # direct picture
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
            return True
        # group: iterate children
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP and getattr(shape, "shapes", None):
            for s in shape.shapes:
                try:
                    if _shape_has_picture(s):
                        return True
                except Exception:
                    continue
        # fill picture
        try:
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
        except Exception:
            pass
        # other shape types containing embedded pictures (chart, smart art, table, canvas)
        if getattr(shape, "shape_type", None) in (
            MSO_SHAPE_TYPE.CHART,
            MSO_SHAPE_TYPE.SMART_ART,
            MSO_SHAPE_TYPE.TABLE,
            MSO_SHAPE_TYPE.CANVAS,
        ):
            xml = getattr(shape._element, "xml", "") or ""
            xml = xml.lower()
            if any(tag in xml for tag in ("p:pic", "a:blip", "blipfill", "a:blipfill")):
                return True
        # final fallback: any xml tag for picture
        xml = getattr(shape._element, "xml", "") or ""
        xml = xml.lower()
        if any(tag in xml for tag in ("p:pic", "a:blip", "a:blipfill", "blipfill")):
            return True
    except Exception:
        pass
    return False

def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = list(prs.slides)
    num_slides = len(slides)

    # detect pictures across all slides
    has_picture = False
    for slide in slides:
        for shape in slide.shapes:
            try:
                if _shape_has_picture(shape):
                    has_picture = True
                    break
            except Exception:
                continue
        if has_picture:
            break

    # detect transitions: look for transition tag in slide xml (works generally)
    has_transition = False
    try:
        for slide in slides:
            try:
                xml = getattr(slide._element, "xml", "") or ""
                if "transition" in xml.lower() or "p:transition" in xml.lower():
                    has_transition = True
                    break
            except Exception:
                continue
    except Exception:
        has_transition = False

    # collect all text (no-accent)
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                try:
                    ppt_text += " " + (shape.text or "")
                except Exception:
                    continue
    ppt_text_norm = _no_accent(ppt_text)

    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        pts = float(it.get("diem", 0) or 0)
        ok = False

        try:
            if key == "min_slides":
                val = int(it.get("value", 1) or 1)
                ok = num_slides >= val
            elif key == "has_image":
                ok = has_picture
            elif key == "has_transition":
                ok = has_transition
            elif key == "title_first":
                if slides:
                    first = slides[0]
                    ok = any(getattr(shape, "has_text_frame", False) and (shape.text or "").strip() for shape in first.shapes)
            elif key.startswith("contains:"):
                terms = key.split(":",1)[1].split("|")
                ok = any(_no_accent(t.strip()) in ppt_text_norm for t in terms if t.strip())
            elif key == "any" or key == "":
                ok = True
            else:
                words = re.findall(r"\w+", _no_accent(desc))
                ok = any(w for w in words if len(w) > 1 and w in ppt_text_norm)
        except Exception:
            ok = False

        if ok:
            total += pts
            notes.append(f"✅ {desc} (+{pts})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes

# ---------------- Scratch grading ----------------
def analyze_sb3_basic(path):
    tempdir = None
    try:
        tempdir = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(path, "r") as z:
            z.extractall(tempdir)
        proj = None
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
            blocks = t.get("blocks", {}) or {}
            for bid, block in blocks.items():
                opcode = (block.get("opcode") or "").lower()
                if any(k in opcode for k in ("control_repeat","control_forever","control_repeat_until")):
                    flags["has_loop"] = True
                if any(k in opcode for k in ("control_if","control_if_else")):
                    flags["has_condition"] = True
                if any(k in opcode for k in ("sensing_keypressed","sensing_touchingobject",
                                             "event_whenthisspriteclicked","event_whenflagclicked",
                                             "event_whenbroadcastreceived","sensing_mousedown")):
                    flags["has_interaction"] = True
                if any(k in opcode for k in ("data_setvariableto","data_changevariableby","data_hidevariable","data_showvariable")):
                    flags["has_variable"] = True
                if "event_broadcast" in opcode:
                    flags["has_interaction"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi khi phân tích SB3: {e}"]
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

    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        pts = float(it.get("diem", 0) or 0)
        ok = False
        try:
            if key == "has_loop":
                ok = flags.get("has_loop", False)
            elif key == "has_condition":
                ok = flags.get("has_condition", False)
            elif key == "has_interaction":
                ok = flags.get("has_interaction", False)
            elif key == "has_variable":
                ok = flags.get("has_variable", False)
            elif key in ("multiple_sprites_or_animation", "has_multiple_sprites"):
                ok = flags.get("multiple_sprites_or_animation", False)
            elif key == "any" or key == "":
                ok = True
            else:
                ok = any(k in desc.lower() for k in ["vòng","lặp","biến","broadcast","phát sóng","điều kiện","nối"])
        except Exception:
            ok = False

        if ok:
            total += pts
            notes.append(f"✅ {desc} (+{pts})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes

