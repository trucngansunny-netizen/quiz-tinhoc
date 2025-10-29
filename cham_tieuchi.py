# ---------------- Word grading ----------------
def grade_word(file_path, criteria):
    """
    Returns (total_points, notes_list)
    notes_list: lines like "✅ Mo ta (+1/1)" or "❌ Mo ta (+0/1)"
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

        if key == "has_title":
            ok = any(len(p.text.strip()) > 0 for p in doc.paragraphs[:2])
        elif key == "has_name":
            ok = bool(re.search(r"\b(họ tên|họ và tên|tên học sinh|họ tên:)\b", text))
        elif key == "has_image":
            ok = any("graphicData" in p._element.xml for p in doc.paragraphs)
        elif key == "format_text":
            ok = any(run.bold for p in doc.paragraphs for run in p.runs if run.bold)
        elif key == "any":
            ok = True
        elif key.startswith("contains:"):
            terms = key.split(":", 1)[1].split("|")
            ok = any(t.strip().lower() in text for t in terms)
        else:
            words = re.findall(r"\w+", desc.lower())
            ok = any(w for w in words if len(w) > 1 and w in text)

        earned = pts if ok else 0.0
        total_awarded += earned
        mark = "✅" if ok else "❌"
        notes.append(f"{mark} {desc} (+{earned}/{pts})")

    notes.append(f"— Tổng điểm: {round(total_awarded,2)}/{round(total_max,2)}")
    return round(total_awarded,2), notes


# ---------------- PPT grading ----------------
def grade_ppt(file_path, criteria):
    try:
        prs = Presentation(file_path)
        slides = prs.slides
    except Exception as e:
        return None, [f"Lỗi đọc file PowerPoint: {e}"]

    items = criteria.get("tieu_chi", [])
    total_awarded = 0.0
    total_max = 0.0
    notes = []

    num_slides = len(slides)
    has_picture_any = any(getattr(shape, "shape_type", None) == 13 for slide in slides for shape in slide.shapes)
    has_transition_any = any("transition" in slide._element.xml for slide in slides)
    ppt_text = " ".join(
        shape.text for slide in slides for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ).lower()

    for it in items:
        desc = it.get("mo_ta", "").strip()
        key = it.get("key", "").lower() if it.get("key") else ""
        pts = float(it.get("diem", 0) or 0)
        total_max += pts
        ok = False

        if key == "min_slides":
            ok = num_slides >= int(it.get("value", 1))
        elif key == "title_first":
            ok = slides and any(shape.has_text_frame and shape.text.strip() for shape in slides[0].shapes)
        elif key == "has_image":
            ok = has_picture_any
        elif key == "has_transition":
            ok = has_transition_any
        elif key == "format_text":
            ok = any(
                run.font.bold
                for slide in slides
                for shape in slide.shapes
                if getattr(shape, "has_text_frame", False)
                for paragraph in shape.text_frame.paragraphs
                for run in paragraph.runs
                if run.font and run.font.bold
            )
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


# ---------------- Scratch grading ----------------
def grade_scratch(file_path, criteria):
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

        if key and key in flags:
            ok = flags.get(key, False)
        elif key == "any":
            ok = True
        else:
            ok = any(k in desc.lower() for k in ["vòng", "lặp", "biến", "broadcast", "phát sóng", "điều kiện", "nối"])

        earned = pts if ok else 0.0
        total_awarded += earned
        mark = "✅" if ok else "❌"
        notes.append(f"{mark} {desc} (+{earned}/{pts})")

    notes.append(f"— Tổng điểm: {round(total_awarded,2)}/{round(total_max,2)}")
    return round(total_awarded,2), notes
