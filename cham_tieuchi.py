# cham_tieuchi.py
# Phiên bản cập nhật: đọc tiêu chí bền hơn, chấm chính xác (điểm thập phân), scale tổng về 10.
import os
import json
import zipfile
import tempfile
import shutil
import re

from docx import Document
from pptx import Presentation

# ---------------- Utilities ----------------
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
    Try many filename patterns to locate the JSON criteria file.
    subject_prefix: e.g. "ppt", "word", "scratch" (lowercase)
    grade: int 3/4/5
    returns full path or None
    """
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
    # also try uppercase/lowercase variants by scanning folder
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            lower = fn.lower()
            if fn.lower().startswith(subject_prefix.lower()) and str(grade) in lower and fn.lower().endswith(".json"):
                return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    subject: "word" / "ppt" / "scratch" OR "PowerPoint"/"Word"/"Scratch"
    grade: int or str
    returns dict with key 'tieu_chi': list of dicts {mo_ta, diem, key, value?}
    """
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    # normalize subject
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
        # normalize structure
        if isinstance(data, dict) and "tieu_chi" in data:
            # ensure diem numeric
            for it in data["tieu_chi"]:
                if "diem" in it:
                    try:
                        it["diem"] = float(it["diem"])
                    except:
                        it["diem"] = 0.0
                else:
                    it["diem"] = 0.0
            return data
        else:
            # maybe the file is a list directly
            if isinstance(data, list):
                return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ---------------- Scaling helper ----------------
def scale_scores(criteria):
    """
    Ensure total of criteria points maps to 10.0 by computing scale factor.
    Returns list of criteria with original diem and scaled diem (added key 'diem_scaled').
    """
    items = criteria.get("tieu_chi", [])
    total_ori = sum([float(it.get("diem", 0) or 0) for it in items])
    if total_ori == 0:
        # nothing to scale, keep as-is
        for it in items:
            it["diem_scaled"] = float(it.get("diem", 0) or 0)
        return items, 1.0
    scale = 10.0 / total_ori
    for it in items:
        it["diem_scaled"] = round(float(it.get("diem", 0) or 0) * scale, 2)
    # small adjust to ensure sum scaled = 10.0 (fix rounding)
    ssum = round(sum(it["diem_scaled"] for it in items), 2)
    diff = round(10.0 - ssum, 2)
    if abs(diff) >= 0.01:
        # add diff to the largest diem_scaled
        max_it = max(items, key=lambda x: x["diem_scaled"])
        max_it["diem_scaled"] = round(max_it["diem_scaled"] + diff, 2)
    return items, scale

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    """
    Returns (total_scaled_points, notes_list)
    notes_list: lines like "✅ Mo ta (+scaled/ori)" or "❌ Mo ta (+0/ori)"
    """
    # read docx
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = "\n".join([p.text for p in doc.paragraphs]).lower()
    # prepare criteria
    items, scale = scale_scores(criteria)

    total_awarded = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        ori = float(it.get("diem", 0) or 0)
        scaled = float(it.get("diem_scaled", 0) or 0)
        ok = False

        # checks
        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = bool(re.search(r"\b(họ tên|họ và tên|tên học sinh|họ tên:)\b", text))
        elif key == "has_image":
            # detect if any inline shape exists in document.xml parts
            ok = any("graphicData" in p._element.xml for p in doc.paragraphs)
        elif key == "format_text":
            # check for some bold or font size differences
            bold_text = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
            ok = bold_text
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            # key like "contains:nguyễn|năm"
            terms = key.split(":",1)[1].split("|")
            ok = any(t.strip().lower() in text for t in terms)
        else:
            # fallback: look for important words from description
            words = re.findall(r"\w+", desc.lower())
            # require at least one meaningful word match
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total_awarded += scaled
            notes.append(f"✅ {desc} (+{scaled}/{ori})")
        else:
            notes.append(f"❌ {desc} (+0/{ori})")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes

# ---------------- PPT grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items, scale = scale_scores(criteria)
    total_awarded = 0.0
    notes = []

    # gather some quick slide info
    num_slides = len(slides)
    has_picture_any = any(getattr(shape, "shape_type", None) == 13 for slide in slides for shape in slide.shapes)
    has_transition_any = any("transition" in slide._element.xml for slide in slides)
    # aggregate text
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                ppt_text += " " + (shape.text or "")

    ppt_text = ppt_text.lower()

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        ori = float(it.get("diem", 0) or 0)
        scaled = float(it.get("diem_scaled", 0) or 0)
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = num_slides >= val
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(shape.has_text_frame and shape.text.strip() for shape in first.shapes)
            else:
                ok = False
        elif key == "has_image":
            ok = has_picture_any
        elif key == "has_transition":
            ok = has_transition_any
        elif key == "format_text":
            # detect bold or colored runs
            bold_or_colored = False
            for slide in slides:
                for shape in slide.shapes:
                    if not getattr(shape, "has_text_frame", False): continue
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            try:
                                if run.font.bold:
                                    bold_or_colored = True; break
                                if run.font.color and hasattr(run.font.color, "rgb") and run.font.color.rgb:
                                    bold_or_colored = True; break
                            except Exception:
                                continue
                        if bold_or_colored: break
                    if bold_or_colored: break
                if bold_or_colored: break
            ok = bold_or_colored
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(t.strip().lower() in ppt_text for t in terms)
        else:
            words = re.findall(r"\w+", desc.lower())
            ok = any(w for w in words if len(w) > 1 and w in ppt_text)

        if ok:
            total_awarded += scaled
            notes.append(f"✅ {desc} (+{scaled}/{ori})")
        else:
            notes.append(f"❌ {desc} (+0/{ori})")

    total_awarded = round(total_awarded, 2)
    return total_awarded, notes

# ---------------- Scratch grading ----------------
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
            for block_id, block in blocks.items():
                opcode = block.get("opcode", "").lower()
                if any(k in opcode for k in ["control_repeat","control_forever","control_repeat_until"]):
                    flags["has_loop"] = True
                if any(k in opcode for k in ["control_if","control_if_else"]):
                    flags["has_condition"] = True
                if any(k in opcode for k in ["sensing_keypressed","sensing_touchingobject","event_whenthisspriteclicked",
                                             "event_whenflagclicked","event_whenbroadcastreceived","sensing_mousedown"]):
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
            try:
                shutil.rmtree(tempdir)
            except Exception:
                pass

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err

    items, scale = scale_scores(criteria)
    total_awarded = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        ori = float(it.get("diem", 0) or 0)
        scaled = float(it.get("diem_scaled", 0) or 0)
        ok = False
        if key == "has_loop":
            ok = flags.get("has_loop", False)
        elif key == "has_condition":
            ok = flags.get("has_condition", False)
        elif key == "has_interaction":
            ok = flags.get("has_interaction", False)
        elif key == "has_variable":
            ok = flags.get("has_variable", False)
        elif key == "multiple_sprites_or_animation":
            ok = flags.get("multiple_sprites_or_animation", False)
        elif key == "any":
            ok = True
        else:
            # fallback: if description contains 'vòng'/'lặp'/'biến' etc
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "broadcast", "phát sóng", "điều kiện", "nối"])
        if ok:
            total_awarded += scaled
            notes.append(f"✅ {desc} (+{scaled}/{ori})")
        else:
            notes.append(f"❌ {desc} (+0/{ori})")
    total_awarded = round(total_awarded, 2)
    return total_awarded, notes
