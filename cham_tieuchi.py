# cham_tieuchi.py
# Phiên bản: giữ nguyên chức năng đọc tiêu chí và kiểm tra, KHÔNG quy về 10 (trả về điểm raw)

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
    """
    Trích tên học sinh từ tên file (bỏ đuôi, tách camel/case, thay _ và - bằng space)
    Ví dụ: "LeAnhDung.pptx" -> "Le Anh Dung"
    """
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
    # also try scanning folder for variants containing subject and grade
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            lower = fn.lower()
            if lower.endswith(".json") and subject_prefix.lower() in lower and str(grade) in lower:
                return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    Load and normalize criteria file.
    subject: "word"/"ppt"/"scratch" or "Word"/"PowerPoint"/"Scratch"
    grade: int or str
    returns dict like {"tieu_chi": [ {mo_ta, diem, key?, value?}, ... ] } or None
    """
    # normalize grade
    if isinstance(grade, str) and grade.isdigit():
        grade = int(grade)
    # normalize subject prefix
    s = (subject or "").lower()
    if s in ("powerpoint", "ppt", "pptx", "bai thuyet trinh"):
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
    elif isinstance(data, list):
        # convert list to dict format
        normalized = {"tieu_chi": []}
        for it in data:
            if isinstance(it, dict):
                item = dict(it)
                if "diem" in item:
                    try:
                        item["diem"] = float(item["diem"])
                    except:
                        item["diem"] = 0.0
                else:
                    item["diem"] = 0.0
                normalized["tieu_chi"].append(item)
        return normalized
    else:
        return None

# ---------------- Scratch analysis ----------------
def analyze_sb3_basic(file_path):
    """
    Extract simple flags from .sb3 project.json for Scratch checks.
    Returns (flags_dict, []) or (None, [error_message])
    """
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
                opcode = (block.get("opcode") or "").lower()
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

# ---------------- Grading functions (RAW scores, no scaling to 10) ----------------

# WORD
def grade_word(file_path, criteria):
    """
    Chấm Word: trả về (total_awarded_raw, notes_list)
    notes_list chứa từng dòng dạng "✅ Mo ta (+earned/original)" hoặc "❌ Mo ta (+0/original)"
    """
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = "\n".join([p.text for p in doc.paragraphs]).lower()
    items = criteria.get("tieu_chi", [])

    total_awarded = 0.0
    total_max = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        pts = float(it.get("diem", 0) or 0)
        total_max += pts
        ok = False

        # kiểm tra theo key (các key chung hay dùng)
        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = bool(re.search(r"\b(họ tên|họ và tên|tên học sinh|họ tên:)\b", text))
        elif key == "has_image":
            # detect inline shape xml in paragraphs
            ok = any("graphicData" in p._element.xml for p in doc.paragraphs)
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if getattr(run, "bold", False))
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(t.strip().lower() in text for t in terms)
        else:
            # fallback: check words from description
            words = re.findall(r"\w+", desc.lower())
            ok = any(w for w in words if len(w) > 1 and w in text)

        earned = pts if ok else 0.0
        total_awarded += earned
        mark = "✅" if ok else "❌"
        notes.append(f"{mark} {desc} (+{earned}/{pts})")

    notes.append(f"— Tổng điểm: {round(total_awarded,2)}/{round(total_max,2)}")
    return round(total_awarded,2), notes

# POWERPOINT
def grade_ppt(file_path, criteria):
    """
    Chấm PowerPoint: trả về (total_awarded_raw, notes_list)
    """
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    total_max = 0.0
    notes = []

    # aggregate ppt info
    num_slides = len(slides)
    has_picture_any = any(getattr(shape, "shape_type", None) == 13 for slide in slides for shape in slide.shapes)
    has_transition_any = any("transition" in (slide._element.xml or "") for slide in slides)
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                ppt_text += " " + (shape.text or "")
    ppt_text = ppt_text.lower()

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        pts = float(it.get("diem", 0) or 0)
        total_max += pts
        ok = False

        if key == "min_slides":
            try:
                val = int(it.get("value", 1))
            except:
                val = 1
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
            bold_or_colored = False
            for slide in slides:
                for shape in slide.shapes:
                    if not getattr(shape, "has_text_frame", False):
                        continue
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            try:
                                if getattr(run.font, "bold", False):
                                    bold_or_colored = True
                                    break
                                if getattr(run.font, "color", None) and hasattr(run.font.color, "rgb") and run.font.color.rgb:
                                    bold_or_colored = True
                                    break
                            except Exception:
                                continue
                        if bold_or_colored:
                            break
                    if bold_or_colored:
                        break
                if bold_or_colored:
                    break
            ok = bold_or_colored
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(t.strip().lower() in ppt_text for t in terms)
        else:
            words = re.findall(r"\w+", desc.lower())
            ok = any(w for w in words if len(w) > 1 and w in ppt_text)

        earned = pts if ok else 0.0
        total_awarded += earned
        mark = "✅" if ok else "❌"
        notes.append(f"{mark} {desc} (+{earned}/{pts})")

    notes.append(f"— Tổng điểm: {round(total_awarded,2)}/{round(total_max,2)}")
    return round(total_awarded,2), notes

# SCRATCH
def grade_scratch(file_path, criteria):
    """
    Chấm Scratch (sb3): trả về (total_awarded_raw, notes_list)
    """
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    total_max = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        pts = float(it.get("diem", 0) or 0)
        total_max += pts
        ok = False

        # nếu key là 1 trong flags
        if key and key in flags:
            ok = bool(flags.get(key, False))
        elif key == "any":
            ok = True
        else:
            # fallback: check keywords in description
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "broadcast", "phát sóng", "điều kiện", "nối", "sprite"])

        earned = pts if ok else 0.0
        total_awarded += earned
        mark = "✅" if ok else "❌"
        notes.append(f"{mark} {desc} (+{earned}/{pts})")

    notes.append(f"— Tổng điểm: {round(total_awarded,2)}/{round(total_max,2)}")
    return round(total_awarded,2), notes
