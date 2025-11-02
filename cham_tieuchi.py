# cham_tieuchi.py
# Phiên bản hoàn chỉnh: nhận dạng ảnh mạnh, không scale về 10, tương thích ai_tin_web.py
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
def _normalize_no_diacritics(s):
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s

def pretty_name_from_filename(filename):
    """
    Trích tên học sinh từ tên file: 'LeAnhDung.pptx' -> 'Le Anh Dung'
    (Hàm này được ai_tin_web.py gọi)
    """
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)

# ---------------- Load criteria ----------------
def _find_criteria_file(prefix, grade, folder="criteria"):
    """
    Thử các mẫu tên file tiêu chí để tương thích với tên bạn đã dùng.
    Ví dụ: ppt3.json, ppt_khoi3.json, ppt_khoi_3.json, ppt- khoi3.json
    """
    candidates = [
        f"{prefix}{grade}.json",
        f"{prefix}_khoi{grade}.json",
        f"{prefix}_khoi_{grade}.json",
        f"{prefix}-khoi{grade}.json",
        f"{prefix}_khoi{grade}.JSON",
        f"{prefix}{grade}.JSON",
    ]
    for c in candidates:
        p = os.path.join(folder, c)
        if os.path.exists(p):
            return p
    # scan folder for loose matches
    if os.path.isdir(folder):
        for fn in os.listdir(folder):
            lower = fn.lower()
            if lower.endswith(".json") and lower.startswith(prefix.lower()) and str(grade) in lower:
                return os.path.join(folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    subject: 'word' / 'ppt' / 'scratch' or variations like 'PowerPoint'
    grade: int or numeric string
    return dict {'tieu_chi': [...] } or None
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
        # normalize structure and ensure numeric diem kept as-is
        if isinstance(data, dict) and "tieu_chi" in data:
            for it in data["tieu_chi"]:
                if "diem" in it:
                    try:
                        # giữ nguyên điểm gốc (không scale)
                        it["diem"] = float(it["diem"])
                    except:
                        it["diem"] = 0.0
                else:
                    it["diem"] = 0.0
            return data
        elif isinstance(data, list):
            return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    """
    Trả về (total_points, notes_list)
    notes_list: "✅ mo_ta (+diem)" hoặc "❌ mo_ta (+0)"
    So khớp text không phân biệt dấu.
    Phát hiện ảnh dựa trên inline_shapes hoặc XML kiểm tra 'graphicData'
    """
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = _normalize_no_diacritics("\n".join([p.text for p in doc.paragraphs]))
    items = criteria.get("tieu_chi", [])
    total = 0.0
    notes = []

    # helper: detect any image in doc (inline_shapes or xml)
    try:
        has_inline_images = len(getattr(doc, "inline_shapes", [])) > 0
    except Exception:
        has_inline_images = False
    # xml check
    xml_all = ""
    try:
        xml_all = "\n".join(p._element.xml for p in doc.paragraphs).lower()
    except Exception:
        xml_all = ""

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
            ok = has_inline_images or ("graphicdata" in xml_all or "a:blip" in xml_all or "p:pic" in xml_all)
        elif key == "format_text":
            # detect bold runs
            bold_found = any(getattr(run, "bold", False) for p in doc.paragraphs for run in p.runs)
            ok = bold_found
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(_normalize_no_diacritics(t.strip()) in text for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            # fallback: look for words from description
            words = re.findall(r"\w+", _normalize_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes

# ---------------- PPT grading (loại trừ thông minh - nhận dạng ảnh mạnh) ----------------
def _shape_contains_image_by_xml(shape):
    """
    Kiểm tra thô bằng XML của shape: nếu có 'a:blip' / 'p:pic' / 'blipfill' / 'href="http' -> ảnh
    """
    try:
        xml = shape._element.xml.lower()
        for tag in ("a:blip", "p:pic", "blipfill", "a:blipfill", "href=\"http", "r:embed"):
            if tag in xml:
                return True
    except Exception:
        pass
    return False

def _shape_is_text_only(shape):
    """
    Trả True nếu shape chỉ là textbox/shape chứa text (chỉ text, không ảnh)
    """
    try:
        # các loại shape tường minh chứa text
        if getattr(shape, "has_text_frame", False):
            # if has text but xml also contains image tags, treat as not text-only
            if _shape_contains_image_by_xml(shape):
                return False
            # if text present, it's text shape
            return True
        # some shapes types are known non-image (lines, auto shapes, text boxes, tables, charts without image)
        st = getattr(shape, "shape_type", None)
        if st in (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.CONNECTOR):
            return True
    except Exception:
        pass
    return False

def _shape_has_picture(shape):
    """
    Mọi kiểm tra để coi shape này có ảnh theo nhiều cách:
    - shape_type == PICTURE
    - fill.type == picture (6)
    - xml contains blip/pic/blipfill or href
    - group shapes: đệ quy
    - smartart/chart/table may include blip in xml
    """
    try:
        st = getattr(shape, "shape_type", None)
        if st == MSO_SHAPE_TYPE.PICTURE:
            return True
        # group -> check items
        if st == MSO_SHAPE_TYPE.GROUP:
            try:
                for s in shape.shapes:
                    if _shape_has_picture(s):
                        return True
            except Exception:
                pass
        # fill picture
        try:
            fill = getattr(shape, "fill", None)
            # pptx uses numeric constants for type, picture fill is 6 typically
            if getattr(fill, "type", None) == 6:
                return True
        except Exception:
            pass
        # xml check (covers SmartArt, Chart with images, web refs)
        if _shape_contains_image_by_xml(shape):
            return True
    except Exception:
        pass
    return False

def _slide_has_image(slide):
    """
    Slide có ảnh nếu tồn tại shape nào không phải text-only và passes image checks.
    Loại trừ shapes rõ ràng chỉ chứa text/table/chart line...
    """
    try:
        for shape in slide.shapes:
            # if shape explicitly contains image xml -> True
            if _shape_has_picture(shape):
                return True
            # if shape not text-only and xml contains possible image markers
            if not _shape_is_text_only(shape) and _shape_contains_image_by_xml(shape):
                return True
        return False
    except Exception:
        return False

def grade_ppt(file_path, criteria):
    """
    Duyệt slides bằng python-pptx, loại trừ slide chỉ có chữ/số,
    phát hiện ảnh bằng XML/fill/group/SmartArt/Chart.
    Trả về (total_points, notes_list).
    """
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    # đảm bảo diem numeric (giữ nguyên)
    for it in items:
        try:
            it["diem"] = float(it.get("diem", 0) or 0)
        except:
            it["diem"] = 0.0

    total = 0.0
    notes = []
    slides = list(prs.slides)
    num_slides = len(slides)

    # build aggregated text (normalize no accent)
    ppt_text = ""
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                try:
                    ppt_text += " " + (shape.text or "")
                except Exception:
                    pass
    ppt_text_norm = _normalize_no_diacritics(ppt_text)

    # precompute per-slide image existence if needed
    slide_has_image_flags = [ _slide_has_image(s) for s in slides ]

    # detect transitions by xml search (fast)
    has_transition_any = any("transition" in (getattr(s, "_element").xml or "").lower() for s in slides)

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "min_slides":
            val = int(it.get("value", 1))
            ok = (num_slides >= val)
        elif key == "has_image":
            ok = any(slide_has_image_flags)
        elif key == "has_transition":
            ok = has_transition_any
        elif key == "title_first":
            ok = False
            if num_slides >= 1:
                first = slides[0]
                # check any text in first slide shapes
                for shape in first.shapes:
                    try:
                        if getattr(shape, "has_text_frame", False) and (shape.text or "").strip():
                            ok = True
                            break
                    except Exception:
                        continue
        elif key.startswith("contains:"):
            terms = key.split(":",1)[1].split("|")
            ok = any(_normalize_no_diacritics(t.strip()) in ppt_text_norm for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            # fallback: match words from description
            words = re.findall(r"\w+", _normalize_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in ppt_text_norm)

        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    total = round(total, 2)
    return total, notes

# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = None
    try:
        tempdir = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(file_path, 'r') as z:
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
            blocks = t.get("blocks", {}) or {}
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
        elif key in ("multiple_sprites_or_animation", "has_multiple_sprites"):
            ok = flags.get("multiple_sprites_or_animation", False)
        elif key == "any":
            ok = True
        else:
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "broadcast", "phát sóng", "điều kiện", "nối"])
        if ok:
            total += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    total = round(total, 2)
    return total, notes
