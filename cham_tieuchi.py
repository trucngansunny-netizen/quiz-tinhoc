import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation

def normalize_text_no_diacritics(s):
    if not isinstance(s, str): return ""
    s = unicodedata.normalize('NFD', s.lower())
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = re.sub(r'[_\-]+', ' ', base)
    s = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', s)
    s = re.sub(r'(\d+)', r' \1 ', s)
    return " ".join(p.capitalize() for p in s.split())

def find_criteria_file(prefix, grade, folder="criteria"):
    for f in os.listdir(folder):
        lf = f.lower()
        if lf.startswith(prefix) and str(grade) in lf and lf.endswith(".json"):
            return os.path.join(folder, f)
    return None

def load_criteria(subject, grade, folder="criteria"):
    s = subject.lower()
    if "ppt" in s: pref = "ppt"
    elif "word" in s: pref = "word"
    elif "scratch" in s: pref = "scratch"
    else: pref = s
    path = find_criteria_file(pref, grade, folder)
    if not path: return None
    with open(path, encoding="utf-8") as f: data = json.load(f)
    return data if isinstance(data, dict) else {"tieu_chi": data}

# ---------------- WORD ----------------
def grade_word(file_path, criteria):
    try:
        doc = Document(file_path)
    except Exception as e:
        return None, [f"Lỗi Word: {e}"]

    text = normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total, notes = 0, []
    for it in criteria.get("tieu_chi", []):
        desc, key = it.get("mo_ta", ""), (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False
        if key == "has_image":
            ok = bool(doc.inline_shapes) or "a:blip" in "\n".join(p._element.xml for p in doc.paragraphs)
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in text for w in words if len(w) > 1)
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes

# ---------------- PPT ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
    except Exception as e:
        return None, [f"Lỗi PowerPoint: {e}"]

    has_picture_any = False
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            for name in z.namelist():
                low = name.lower()
                # Bước 1: Tìm trong ppt/media/
                if low.startswith("ppt/media/") and low.endswith((
                    ".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",
                    ".svg",".webp",".ico",".emf",".wmf",".jfif",".heic",".avif"
                )):
                    has_picture_any = True
                    break
            # Bước 2: Nếu chưa, quét toàn bộ XML trong ppt/slides/, slideLayouts/, slideMasters/
            if not has_picture_any:
                for name in z.namelist():
                    if name.lower().startswith(("ppt/slides/", "ppt/slidelayouts/", "ppt/slidemasters/")) and name.endswith(".xml"):
                        xml = z.read(name).decode("utf-8", errors="ignore").lower()
                        if "<a:blip" in xml or "r:link=" in xml or "svgblip" in xml or "blipfill" in xml:
                            has_picture_any = True
                            break
    except Exception:
        has_picture_any = False

    # Số slide, hiệu ứng
    slides = prs.slides
    num_slides = len(slides)
    has_transition = any("transition" in s._element.xml for s in slides)
    ppt_text = normalize_text_no_diacritics(" ".join(
        shape.text for s in slides for shape in s.shapes if getattr(shape, "has_text_frame", False)
    ))

    total, notes = 0, []
    for it in criteria.get("tieu_chi", []):
        desc, key = it.get("mo_ta",""), (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = False
        if key == "has_image": ok = has_picture_any
        elif key == "min_slides": ok = num_slides >= int(it.get("value", 1))
        elif key == "has_transition": ok = has_transition
        else:
            words = re.findall(r"\w+", normalize_text_no_diacritics(desc))
            ok = any(w in ppt_text for w in words if len(w) > 1)
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes

# ---------------- SCRATCH ----------------
def analyze_sb3_basic(file_path):
    tmp = tempfile.mkdtemp(prefix="sb3_")
    try:
        with zipfile.ZipFile(file_path, 'r') as z: z.extractall(tmp)
        with open(os.path.join(tmp, "project.json"), encoding="utf-8") as f: proj = json.load(f)
        flags = {k: False for k in ["has_loop","has_condition","has_interaction","has_variable","multiple_sprites_or_animation"]}
        targets = proj.get("targets", [])
        if sum(1 for t in targets if not t.get("isStage")) >= 2:
            flags["multiple_sprites_or_animation"] = True
        for t in targets:
            for b in t.get("blocks", {}).values():
                op = b.get("opcode", "")
                if "repeat" in op: flags["has_loop"] = True
                if "if" in op: flags["has_condition"] = True
                if "event_when" in op or "sensing_" in op: flags["has_interaction"] = True
                if "data_" in op: flags["has_variable"] = True
        return flags, []
    except Exception as e:
        return None, [f"Lỗi Scratch: {e}"]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def grade_scratch(file_path, criteria):
    flags, err = analyze_sb3_basic(file_path)
    if flags is None: return None, err
    total, notes = 0, []
    for it in criteria.get("tieu_chi", []):
        desc, key = it.get("mo_ta",""), (it.get("key") or "").lower()
        diem = float(it.get("diem", 0))
        ok = flags.get(key, False)
        total += diem if ok else 0
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total, 2), notes
