import os, json, zipfile, tempfile, shutil, re, unicodedata
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------- Utilities ----------------
def normalize_text_no_diacritics(s):
    if not isinstance(s, str): return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s

def pretty_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    s = base.replace("_"," ").replace("-"," ")
    s = re.sub(r'(?<=[a-z])(?=[A-Z])',' ',s)
    s = re.sub(r'(\d+)',r' \1 ',s)
    parts = [p.capitalize() for p in s.split() if p.strip()]
    return " ".join(parts)

def find_criteria_file(subject_prefix, grade, criteria_folder="criteria"):
    patterns = [
        f"{subject_prefix}{grade}.json",
        f"{subject_prefix}_khoi{grade}.json",
        f"{subject_prefix}-khoi{grade}.json"
    ]
    for p in patterns:
        fpath = os.path.join(criteria_folder, p)
        if os.path.exists(fpath): return fpath
    for fn in os.listdir(criteria_folder):
        if fn.lower().startswith(subject_prefix) and str(grade) in fn and fn.lower().endswith(".json"):
            return os.path.join(criteria_folder, fn)
    return None

def load_criteria(subject, grade, criteria_folder="criteria"):
    s = subject.lower()
    if s in ("powerpoint","ppt","pptx"): pref = "ppt"
    elif s in ("word","docx","doc"): pref = "word"
    elif s in ("scratch","sb3"): pref = "scratch"
    else: pref = s
    path = find_criteria_file(pref, grade, criteria_folder)
    if not path: return None
    with open(path,"r",encoding="utf-8") as f: data=json.load(f)
    if isinstance(data,dict) and "tieu_chi" in data:
        for it in data["tieu_chi"]:
            it["diem"]=float(it.get("diem",0) or 0)
        return data
    elif isinstance(data,list):
        return {"tieu_chi":data}
    return None

# ---------------- WORD ----------------
def grade_word(path,criteria):
    try: doc=Document(path)
    except Exception as e: return None,[f"Lỗi đọc file Word: {e}"]
    text=normalize_text_no_diacritics("\n".join(p.text for p in doc.paragraphs))
    total=0; notes=[]
    for it in criteria.get("tieu_chi",[]):
        desc=it.get("mo_ta",""); key=(it.get("key") or "").lower(); diem=float(it.get("diem",0))
        ok=False
        if key=="has_title": ok=any(len(p.text.strip())>0 for p in doc.paragraphs[:2])
        elif key=="has_name": ok=any(k in text for k in["ho ten","ten hoc sinh","ho va ten"])
        elif key=="has_image": ok=bool(doc.inline_shapes) or "graphicdata" in "\n".join(p._element.xml for p in doc.paragraphs)
        elif key=="format_text": ok=any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key=="any": ok=True
        elif key.startswith("contains:"):
            terms=key.split(":",1)[1].split("|"); ok=any(normalize_text_no_diacritics(t) in text for t in terms)
        else:
            words=re.findall(r"\w+",normalize_text_no_diacritics(desc)); ok=any(w in text for w in words if len(w)>1)
        if ok: total+=diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total,2),notes

# ---------------- POWERPOINT ----------------
def grade_ppt(path,criteria):
    try: prs=Presentation(path)
    except Exception as e: return None,[f"Lỗi đọc PowerPoint: {e}"]
    slides=prs.slides
    def shape_has_picture(shape):
        try:
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE: return True
            if shape.shape_type==MSO_SHAPE_TYPE.GROUP:
                return any(shape_has_picture(s) for s in shape.shapes)
            if hasattr(shape,"fill") and getattr(shape.fill,"type",None)==6: return True
            xml=shape._element.xml.lower()
            if any(tag in xml for tag in["p:pic","a:blip","blipfill","a:blipfill"]): return True
        except Exception: return False
        return False
    has_pic=any(shape_has_picture(s) for sl in slides for s in sl.shapes)
    has_transition=any("transition" in sl._element.xml for sl in slides)
    ppt_text=" ".join(s.text for sl in slides for s in sl.shapes if getattr(s,"has_text_frame",False))
    ppt_text_no=normalize_text_no_diacritics(ppt_text)
    total=0; notes=[]
    for it in criteria.get("tieu_chi",[]):
        desc=it.get("mo_ta",""); key=(it.get("key") or "").lower(); diem=float(it.get("diem",0)); ok=False
        if key=="min_slides": ok=len(slides)>=int(it.get("value",1))
        elif key=="has_image":
            if not has_pic:
                # loại trừ slide chỉ toàn chữ, số, ký tự
                for sl in slides:
                    for s in sl.shapes:
                        if getattr(s,"has_text_frame",False):
                            txt=normalize_text_no_diacritics(s.text).strip()
                            if re.fullmatch(r"[a-z0-9\s.,!?;:-]*",txt): continue
                        else:
                            ok=True; break
                    if ok: break
            else: ok=True
        elif key=="has_transition": ok=has_transition
        elif key=="title_first": ok=slides and any(s.has_text_frame and s.text.strip() for s in slides[0].shapes)
        elif key.startswith("contains:"):
            terms=key.split(":",1)[1].split("|")
            ok=any(normalize_text_no_diacritics(t.strip()) in ppt_text_no for t in terms)
        else:
            words=re.findall(r"\w+",normalize_text_no_diacritics(desc))
            ok=any(w in ppt_text_no for w in words if len(w)>1)
        if ok: total+=diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total,2),notes

# ---------------- SCRATCH ----------------
def analyze_sb3_basic(path):
    tmp=None
    try:
        tmp=tempfile.mkdtemp(prefix="sb3_")
        with zipfile.ZipFile(path,"r") as z:z.extractall(tmp)
        pj=os.path.join(tmp,"project.json")
        if not os.path.exists(pj): return None,["Không tìm thấy project.json"]
        data=json.load(open(pj,"r",encoding="utf-8"))
        flags={"has_loop":False,"has_condition":False,"has_interaction":False,"has_variable":False,"multiple_sprites_or_animation":False}
        sprites=[t for t in data.get("targets",[]) if not t.get("isStage",False)]
        if len(sprites)>=2: flags["multiple_sprites_or_animation"]=True
        for t in sprites:
            if t.get("variables"): flags["has_variable"]=True
            for b in t.get("blocks",{}).values():
                op=b.get("opcode","").lower()
                if any(k in op for k in["control_repeat","forever","repeat_until"]): flags["has_loop"]=True
                if any(k in op for k in["control_if","control_if_else"]): flags["has_condition"]=True
                if any(k in op for k in["sensing_","event_when","broadcast"]): flags["has_interaction"]=True
        return flags,[]
    except Exception as e: return None,[str(e)]
    finally:
        if tmp: shutil.rmtree(tmp,ignore_errors=True)

def grade_scratch(path,criteria):
    flags,err=analyze_sb3_basic(path)
    if flags is None: return None,err
    total=0; notes=[]
    for it in criteria.get("tieu_chi",[]):
        desc=it.get("mo_ta",""); key=(it.get("key") or "").lower(); diem=float(it.get("diem",0)); ok=False
        if key in flags: ok=flags[key]
        elif key=="any": ok=True
        if ok: total+=diem
        notes.append(f"{'✅' if ok else '❌'} {desc} (+{diem if ok else 0})")
    return round(total,2),notes
