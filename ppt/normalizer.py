try:
    from .layout_engine import pick_layout_variant
except Exception:
    from layout_engine import pick_layout_variant


def clamp_text(text: str, max_len: int) -> str:
    text = (text or "").strip()
    if len(text) <= max_len:
        return text

    cut_chars = ["。", "！", "？", ".", "!", "?", "；", ";", "：", ":", "，", ",", "、", " "]
    candidate = text[:max_len]
    last_pos = -1
    for ch in cut_chars:
        pos = candidate.rfind(ch)
        if pos > last_pos:
            last_pos = pos
    if last_pos >= max(8, int(max_len * 0.55)):
        return candidate[:last_pos + 1].strip()

    shortened = candidate.rstrip("，,。.!！?？:：;；、 ")
    return shortened + "…"


def smart_shorten(text: str, max_len: int) -> str:
    text = (text or "").strip()
    if len(text) <= max_len:
        return text

    sentence_breaks = ["。", "！", "？", ".", "!", "?"]
    for sep in sentence_breaks:
        if sep in text:
            parts = [p.strip() for p in text.split(sep) if p.strip()]
            built = ""
            for part in parts:
                candidate = (built + part + sep).strip()
                if len(candidate) <= max_len:
                    built = candidate
                else:
                    break
            if built:
                return built

    soft_breaks = ["；", ";", "：", ":", "，", ",", "、"]
    candidate = text[:max_len]
    last_pos = -1
    for sep in soft_breaks:
        pos = candidate.rfind(sep)
        if pos > last_pos:
            last_pos = pos
    if last_pos >= max(8, int(max_len * 0.6)):
        return candidate[:last_pos + 1].strip()

    return clamp_text(text, max_len)


def avoid_incomplete_tail(text: str) -> str:
    text = (text or "").strip()
    if not text:
        return text
    if text.endswith("…"):
        base = text[:-1].rstrip()
        cut_chars = ["。", "！", "？", ".", "!", "?", "；", ";", "：", ":", "，", ",", "、"]
        last_pos = -1
        for ch in cut_chars:
            pos = base.rfind(ch)
            if pos > last_pos:
                last_pos = pos
        if last_pos >= 0:
            return base[:last_pos + 1].strip()
    return text


def _norm_text(text: str) -> str:
    return " ".join((text or "").replace("•", "").replace("·", "").split()).strip().lower()


def dedupe_preserve_order(items: list) -> list:
    seen = set()
    out = []
    for item in items or []:
        s = str(item or "").strip()
        key = _norm_text(s)
        if not s or not key or key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def expand_bullets(bullets: list, target: int = 5, *, allow_templates: bool = True) -> list:
    items = dedupe_preserve_order([str(b).strip() for b in (bullets or []) if str(b).strip()])
    return items[:target]


def split_bullets_to_multiple_slides(slide: dict, max_items: int = 4) -> list:
    bullets = [str(b).strip() for b in (slide.get("bullets") or []) if str(b).strip()]
    bullets = dedupe_preserve_order(bullets)
    if len(bullets) <= max_items:
        return [slide]

    pages = []
    for i in range(0, len(bullets), max_items):
        new_slide = dict(slide)
        new_slide["bullets"] = bullets[i:i + max_items]
        if i > 0:
            new_slide["title"] = clamp_text((slide.get("title") or "继续") + "（续）", 34)
            new_slide["_continued"] = True
        pages.append(new_slide)

    if len(pages) >= 2 and len(pages[-1].get("bullets") or []) <= 1:
        prev = dict(pages[-2])
        prev_bullets = list(prev.get("bullets") or [])
        prev_bullets.extend(pages[-1].get("bullets") or [])
        prev["bullets"] = prev_bullets[:5]
        pages[-2] = prev
        pages.pop()

    return pages


def _clean_text(text: str, max_len: int) -> str:
    return avoid_incomplete_tail(smart_shorten(text, max_len))


def _remove_redundant_points(points: list, *reference_texts: str) -> list:
    ref_keys = {_norm_text(x) for x in reference_texts if _norm_text(x)}
    out = []
    seen = set(ref_keys)
    for item in points or []:
        s = str(item or "").strip()
        key = _norm_text(s)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def _apply_layout_variant(slide: dict) -> dict:
    slide = dict(slide)
    slide["layout_variant"] = pick_layout_variant(slide)
    return slide


def enforce_slide_content_budget(ppt_data: dict) -> dict:
    slides = ppt_data.get("slides", [])
    new_slides = []

    for slide in slides:
        slide = dict(slide)
        slide_type = slide.get("type", "bullet")

        if "title" in slide:
            slide["title"] = _clean_text(slide.get("title", ""), 34)
        if "subtitle" in slide:
            slide["subtitle"] = _clean_text(slide.get("subtitle", ""), 56)
        if "summary" in slide:
            slide["summary"] = _clean_text(slide.get("summary", ""), 96)
        if "highlight" in slide:
            slide["highlight"] = _clean_text(slide.get("highlight", ""), 78)
        if "closing" in slide:
            slide["closing"] = _clean_text(slide.get("closing", ""), 82)
        if "speaker_note" in slide:
            slide["speaker_note"] = clamp_text(slide.get("speaker_note", ""), 260)

        if slide_type == "bullet":
            raw_bullets = [str(b).strip() for b in (slide.get("bullets") or []) if str(b).strip()]
            raw_bullets = dedupe_preserve_order([_clean_text(b, 64) for b in raw_bullets])

            slide["bullets"] = raw_bullets
            split_pages = split_bullets_to_multiple_slides(slide, max_items=4)

            for page in split_pages:
                is_continued = bool(page.get("_continued"))
                target = 3 if is_continued else 5
                page_bullets = expand_bullets(page.get("bullets", []), target=target, allow_templates=not is_continued)
                page_bullets = dedupe_preserve_order([_clean_text(b, 64) for b in page_bullets])

                if is_continued and len(page_bullets) <= 1 and len(new_slides) > 0:
                    prev = dict(new_slides[-1])
                    prev_bullets = dedupe_preserve_order(list(prev.get("bullets") or []) + page_bullets)
                    prev["bullets"] = prev_bullets[:5]
                    new_slides[-1] = _apply_layout_variant(prev)
                    continue

                page["bullets"] = page_bullets
                new_slides.append(_apply_layout_variant(page))
            continue

        elif slide_type in ["image_left", "image_right"]:
            summary = slide.get("summary", "") or ""
            caption = slide.get("image_caption", "") or ""
            bullets = [_clean_text(b, 36) for b in (slide.get("bullets") or []) if str(b).strip()]
            bullets = dedupe_preserve_order(bullets)
            bullets = _remove_redundant_points(bullets, summary, caption)

            slide["bullets"] = bullets[:3]
            if not caption:
                base = summary or slide.get("title", "视觉图像")
                slide["image_caption"] = _clean_text(base, 40)
            else:
                slide["image_caption"] = _clean_text(caption, 40)

        elif slide_type == "compare":
            slide["left_title"] = _clean_text(slide.get("left_title", "方案 A"), 18)
            slide["right_title"] = _clean_text(slide.get("right_title", "方案 B"), 18)
            slide["left_points"] = dedupe_preserve_order([
                _clean_text(p, 38)
                for p in (slide.get("left_points") or [])
                if str(p).strip()
            ])[:4]
            slide["right_points"] = dedupe_preserve_order([
                _clean_text(p, 38)
                for p in (slide.get("right_points") or [])
                if str(p).strip()
            ])[:4]

        elif slide_type == "two_column":
            left_vals = slide.get("left_points") or slide.get("left") or []
            right_vals = slide.get("right_points") or slide.get("right") or []
            slide["left_title"] = _clean_text(slide.get("left_title", "左侧观点"), 18)
            slide["right_title"] = _clean_text(slide.get("right_title", "右侧观点"), 18)
            slide["left_points"] = dedupe_preserve_order([
                _clean_text(p, 40)
                for p in left_vals
                if str(p).strip()
            ])[:4]
            slide["right_points"] = dedupe_preserve_order([
                _clean_text(p, 40)
                for p in right_vals
                if str(p).strip()
            ])[:4]

        elif slide_type == "highlight":
            slide["summary"] = _clean_text(slide.get("summary", ""), 100)
            slide["highlight"] = _clean_text(slide.get("highlight") or slide.get("summary", ""), 72)
            slide["closing"] = _clean_text(slide.get("closing") or slide.get("summary", ""), 80)

        new_slides.append(_apply_layout_variant(slide))

    ppt_data["slides"] = new_slides
    if "summary" in ppt_data:
        ppt_data["summary"] = _clean_text(ppt_data["summary"], 120)
    return ppt_data
