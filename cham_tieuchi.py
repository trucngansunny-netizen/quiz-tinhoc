import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
    if not isinstance(s, str): return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_"," ").replace("-"," ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])',' ',s)
    s = re.sub(r'(\d+)',r' \1 ',s)
    return " ".join(p.capitalize() for p in s.split() if p.strip())

def find_criteria_file(subject_prefix, grade, folder="criteria"):
    for fn in os.listdir(folder):
        low = fn.lower()
        if low.startswith(subject_prefix) and str(grade) in low and low.endswith(".json"):
            return os.path.join(folder, fn)
    return None

def load_criteria(subject, grade, folder="criteria"):
    s = subject.lower()
    pref = "ppt" if "ppt" in s else "word" if "word" in s else "scratch"
    path = find_criteria_file(pref, grade, folder)
    if not path: return None
    with open(path,"r",encoding="utf-8") as f: data = json.load(f)
    if isinstance(data, dict) and "tieu_chi" in data:
        for it in data["tieu_chi"]:
            it["diem"] = float(it.get("diem", 0) or 0)
        return data
    elif isinstance(data, list):
        return {"tieu_chi": data}
    return None

# ---------------- WORD ----------------
def grade_word(path, criteria):
    try: doc = Document(path)
    except Exception as e: return None, [f"Lỗi đọc file Word: {e}"]
    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total = 0; notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", ""); key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0)); ok = False
        if key == "has_title": ok = any(p.text.strip() for p in doc.paragraphs[:2])
        elif key == "has_name": ok = any(k in text for k in ["ho ten", "ten hoc sinh", "ho va ten"])
        elif key == "has_image": ok = bool(doc.inline_shapes)
        elif key == "format_text": ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key == "any": ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in text for w in words if len(w) > 1)
        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes

# ---------------- POWERPOINT ----------------
def grade_ppt(path, criteria):
    try: prs = Presentation(path)
    except Exception as e: return None, [f"Lỗi đọc PowerPoint: {e}"]

    slides = prs.slides
    num_slides = len(slides)

    # nhanh hơn: chỉ kiểm tra các tag cần thiết
    def shape_has_picture(shape):
        try:
            st = shape.shape_type
            if st == MSO_SHAPE_TYPE.PICTURE:
                return True
            if st == MSO_SHAPE_TYPE.GROUP:
                return any(shape_has_picture(s) for s in shape.shapes)
            if hasattr(shape, "fill") and getattr(shape.fill, "type", None) == 6:
                return True
            xml = shape._element.xml
            # kiểm tra nhanh (đoạn đầu XML)
            if "<p:pic" in xml or "<a:blip" in xml or "<a:blipFill" in xml:
                return True
        except Exception:
            pass
        return False

    has_pic = False
    for sl in slides:
        for s in sl.shapes:
            if shape_has_picture(s):
                has_pic = True
                break
        if has_pic: break

    has_transition = any("transition" in sl._element.xml for sl in slides)
    ppt_text = " ".join(s.text for sl in slides for s in sl.shapes if getattr(s, "has_text_frame", False))
    ppt_text_no = normalize_text_no_diacritics(ppt_text)

    total = 0; notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", "")
        key = (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False

        if key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "has_image":
            if has_pic:
                ok = True
            else:
                # loại trừ slide toàn chữ/số → còn lại là có hình hoặc khung
                for sl in slides:
                    slide_ok = False
                    for s in sl.shapes:
                        if getattr(s, "has_text_frame", False):
                            txt = normalize_text_no_diacritics(s.text).strip()
                            if re.fullmatch(r"[a-z0-9\s.,!?;:-]*", txt):
                                continue
                        else:
                            slide_ok = True
                            break
                    if slide_ok:
                        ok = True
                        break
        elif key == "has_transition":
            ok = has_transition
        elif key == "title_first":
            if slides:
                first = slides[0]
                ok = any(s.has_text_frame and s.text.strip() for s in first.shapes)
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(normalize_text_no_diacritics(t.strip()) in ppt_text_no for t in terms if t.strip())
        elif key == "any":
            ok = True
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text_no for w in words if len(w) > 1)
        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")

    return round(total, 2), notes

# ---------------- SCRATCH ----------------
def analyze_sb3_basic(path):
    tmp = None
    try:
        tmp = tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(path, "r") as z: z.extractall(tmp)
        pj = os.path.join(tmp, "project.json")
        if not os.path.exists(pj): return None, ["Không tìm thấy project.json"]
        data = json.load(open(pj, "r", encoding="utf-8"))
        flags = {"has_loop": False, "has_condition": False, "has_interaction": False, "has_variable": False, "multiple_sprites_or_animation": False}
        sprites = [t for t in data.get("targets", []) if not t.get("isStage", False)]
        if len(sprites) >= 2: flags["multiple_sprites_or_animation"] = True
        for t in sprites:
            if t.get("variables"): flags["has_variable"] = True
            for b in t.get("blocks", {}).values():
                op = b.get("opcode", "").lower()
                if any(k in op for k in ["control_repeat", "forever", "repeat_until"]): flags["has_loop"] = True
                if any(k in op for k in ["control_if", "control_if_else"]): flags["has_condition"] = True
                if any(k in op for k in ["sensing_", "event_when", "broadcast"]): flags["has_interaction"] = True
        return flags, []
    except Exception as e:
        return None, [str(e)]
    finally:
        if tmp: shutil.rmtree(tmp, ignore_errors=True)

def grade_scratch(path, criteria):
    flags, err = analyze_sb3_basic(path)
    if flags is None: return None, err
    total = 0; notes = []
    for it in criteria.get("tieu_chi", []):
        desc = it.get("mo_ta", ""); key = (it.get("key") or "").lower(); diem = float(it.get("diem", 0)); ok = False
        if key in flags: ok = flags[key]
        elif key == "any": ok = True
        if ok: total += diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes
