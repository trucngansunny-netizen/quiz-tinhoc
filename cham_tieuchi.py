# cham_tieuchi.py
# Phiên bản cuối: quét XML + media + .rels để nhận diện mọi loại ảnh trong PPTX,
# tương thích với ai_tin_web.py (Render/Gunicorn). Giữ nguyên điểm gốc từ JSON.
import os
import json
import zipfile
import tempfile
import shutil
import re
import unicodedata

from docx import Document
from pptx import Presentation

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
    """Chuẩn hóa chuỗi: bỏ dấu tiếng Việt, viết thường."""
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')

def pretty_name_from_filename(filename):
    """Trích tên học sinh từ tên file (giữ tương thích với ai_tin_web.py)."""
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_", " ").replace("-", " ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    parts = [p for p in s.split() if p.strip()]
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    """Tìm file tiêu chí với nhiều biến thể tên (ppt3.json, ppt_khoi3.json, ...)."""
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
    # scan folder for best match
    if os.path.isdir(criteria_folder):
        for fn in os.listdir(criteria_folder):
            lf = fn.lower()
            if lf.endswith(".json") and lf.startswith(subject_prefix.lower()) and str(grade) in lf:
                return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    """
    Đọc tiêu chí. signature tương thích: load_criteria(subject, grade, criteria_folder="criteria")
    subject: "word"/"ppt"/"scratch" hoặc "PowerPoint"/"Word"/...
    grade: int or str
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

    path = find_criteria_file(pref, grade, criteria_folder)
    if not path:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # normalize diem numeric
        if isinstance(data, dict) and "tieu_chi" in data:
            for it in data["tieu_chi"]:
                try:
                    it["diem"] = float(it.get("diem", 0))
                except:
                    it["diem"] = 0.0
            return data
        elif isinstance(data, list):
            # list of criteria directly
            for it in data:
                try:
                    it["diem"] = float(it.get("diem", 0))
                except:
                    it["diem"] = 0.0
            return {"tieu_chi": data}
    except Exception:
        return None
    return None

# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    """
    Chấm Word. Trả về (total_points, notes_list).
    notes_list: các chuỗi '✅ ... (+diem)' hoặc '❌ ... (+0)'
    """
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []

    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten", "họ và tên"])
        elif key == "has_image":
            # doc.inline_shapes works for many Word images
            try:
                ok = bool(getattr(doc, "inline_shapes", None) and len(doc.inline_shapes) > 0)
            except Exception:
                ok = False
            # fallback: check xml for 'graphicData' or 'a:blip'
            if not ok:
                try:
                    xml_all = "\n".join(p._element.xml for p in doc.paragraphs)
                    if "graphicdata" in xml_all.lower() or "a:blip" in xml_all.lower():
                        ok = True
                except Exception:
                    pass
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if getattr(run, "bold", False))
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            # fallback: search words from description
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w for w in words if len(w) > 1 and w in text)

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes

# ---------------- Helper: PPTX XML + media scanner ----------------
_IMAGE_EXTS = (
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff",
    ".svg", ".webp", ".ico", ".emf", ".wmf", ".jfif", ".heic", ".avif"
)

def pptx_contains_image_by_zip(path):
    """
    Phương pháp an toàn: mở zip, quét:
      - ppt/media/* với đuôi ảnh
      - tất cả XML trong ppt/slides/, ppt/slideLayouts/, ppt/slideMasters/ tìm <a:blip r:link= hoặc <a:blip r:embed= hoặc p:pic hoặc a:blipfill
      - .rels files tìm Target (nếu target là http hoặc file image)
    Trả về True nếu phát hiện khả nghi có ảnh.
    """
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            # 1) check ppt/media
            for n in names:
                ln = n.lower()
                if ln.startswith("ppt/media/") and any(ln.endswith(ext) for ext in _IMAGE_EXTS):
                    return True
            # 2) scan xmls under slides/layouts/masters and their .rels for blip or link
            xml_candidates = [n for n in names if n.lower().startswith(("ppt/slides/", "ppt/slidelayouts/", "ppt/slidemasters/")) and n.lower().endswith((".xml", ".xml.rels"))]
            for n in xml_candidates:
                try:
                    raw = z.read(n).decode("utf-8", errors="ignore").lower()
                except Exception:
                    try:
                        raw = z.read(n).decode("cp1252", errors="ignore").lower()
                    except Exception:
                        raw = ""
                if not raw:
                    continue
                # look for blip tags or pic tags
                if "<a:blip" in raw or "p:pic" in raw or "a:blipfill" in raw or "blipfill" in raw or "r:link" in raw or "r:embed" in raw:
                    return True
                # check .rels explicit targets (external links or media)
                if n.endswith(".rels"):
                    # find Target="..."
                    for m in re.finditer(r'target="([^"]+)"', raw):
                        tgt = m.group(1).lower()
                        if tgt.startswith("http://") or tgt.startswith("https://"):
                            # external link -> likely online image
                            # consider that as image presence (we treat external link as image presence)
                            # but check extension too
                            if any(tgt.endswith(ext) for ext in _IMAGE_EXTS) or "bing.net" in tgt or "microsoft" in tgt or "cdn" in tgt:
                                return True
                            # even if no extension, r:link often indicates online picture - treat as image
                            return True
                        if any(tgt.endswith(ext) for ext in _IMAGE_EXTS):
                            return True
            # 3) fallback: check any xml containing "<p:pic" or "<a:blip" anywhere
            for n in names:
                ln = n.lower()
                if ln.endswith(".xml"):
                    try:
                        raw = z.read(n).decode("utf-8", errors="ignore").lower()
                    except Exception:
                        raw = ""
                    if "<a:blip" in raw or "<p:pic" in raw or "a:blipfill" in raw or "blipfill" in raw:
                        return True
    except Exception:
        # if any error, fallback to False
        return False
    return False

# ---------------- PowerPoint grading ----------------
def grade_ppt(file_path, criteria):
    """
    Chấm PPT:
     - dùng python-pptx để đọc text/slide count/shape properties
     - dùng pptx_contains_image_by_zip() để phát hiện ảnh mọi dạng
    """
    # Try to open with python-pptx (may fail if file corrupted)
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)
    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []

    # detect images: combination of python-pptx shape checks + deep zip/xml scan
    has_image_by_shapes = False
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        for sl in slides:
            for shape in sl.shapes:
                try:
                    # picture shape
                    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                        has_image_by_shapes = True
                        break
                    # group
                    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
                        for s2 in shape.shapes:
                            if getattr(s2, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                                has_image_by_shapes = True
                                break
                        if has_image_by_shapes:
                            break
                    # fill with picture - type 6 commonly indicates picture fill
                    try:
                        if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                            has_image_by_shapes = True
                            break
                    except Exception:
                        pass
                    # xml snippet: check for blip tags in shape xml
                    try:
                        if "p:pic" in shape._element.xml.lower() or "a:blip" in shape._element.xml.lower():
                            has_image_by_shapes = True
                            break
                    except Exception:
                        pass
                except Exception:
                    continue
            if has_image_by_shapes:
                break
    except Exception:
        has_image_by_shapes = False

    # deep xml/zip scan
    try:
        has_image_by_zip = pptx_contains_image_by_zip(file_path)
    except Exception:
        has_image_by_zip = False

    has_picture_any = bool(has_image_by_shapes or has_image_by_zip)

    # detect transitions simple way (search xml or python-pptx)
    try:
        has_transition_any = any("transition" in sl._element.xml.lower() for sl in slides)
    except Exception:
        has_transition_any = False

    # gather text (no-accent)
    ppt_text = []
    try:
        for sl in slides:
            for shape in sl.shapes:
                if getattr(shape, "has_text_frame", False):
                    try:
                        txt = shape.text or ""
                        if txt:
                            ppt_text.append(txt)
                    except Exception:
                        continue
    except Exception:
        pass
    ppt_text_join = " ".join(ppt_text)
    ppt_text_noaccent = normalize_text_no_diacritics(ppt_text_join)

    # evaluate criteria
    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False

        try:
            if key == "min_slides":
                val = int(it.get("value", 1))
                ok = num_slides >= val
            elif key == "has_image":
                ok = has_picture_any
            elif key == "has_transition":
                ok = has_transition_any
            elif key == "title_first":
                if slides:
                    try:
                        first = slides[0]
                        ok = any(getattr(s, "has_text_frame", False) and (s.text or "").strip() for s in first.shapes)
                    except Exception:
                        ok = False
            elif key.startswith("contains:"):
                terms = key.split(":", 1)[1].split("|")
                ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_noaccent for t in terms if t.strip())
            elif key == "any":
                ok = True
            else:
                words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
                ok = any(w for w in words if len(w) > 1 and w in ppt_text_noaccent)
        except Exception:
            ok = False

        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")

    return round(total_awarded, 2), notes

# ---------------- Scratch grading ----------------
def analyze_sb3_basic(file_path):
    tempdir = tempfile.mkdtemp(prefix="sb3_")
    try:
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
                if any(k in opcode for k in ["control_repeat", "control_forever", "control_repeat_until"]):
                    flags["has_loop"] = True
                if any(k in opcode for k in ["control_if", "control_if_else"]):
                    flags["has_condition"] = True
                if any(k in opcode for k in ["sensing_keypressed", "sensing_touchingobject",
                                             "event_whenthisspriteclicked", "event_whenflagclicked",
                                             "event_whenbroadcastreceived", "sensing_mousedown"]):
                    flags["has_interaction"] = True
                if any(k in opcode for k in ["data_setvariableto", "data_changevariableby", "data_hidevariable", "data_showvariable"]):
                    flags["has_variable"] = True
                if "event_broadcast" in opcode:
                    flags["has_interaction"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi khi phân tích file Scratch: {e}"]
    finally:
        shutil.rmtree(tempdir, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None:
        return None, err
    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    notes = []
    for it in items:
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0) or 0)
        ok = False
        if key in flags:
            ok = flags.get(key, False)
        elif key == "any":
            ok = True
        else:
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "phát sóng", "điều kiện"])
        if ok:
            total_awarded += diem
            notes.append(f"✅ {desc} (+{diem})")
        else:
            notes.append(f"❌ {desc} (+0)")
    return round(total_awarded, 2), notes
