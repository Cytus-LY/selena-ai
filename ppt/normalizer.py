import re

try:
    from .layout_engine import finalize_slide_layouts, maybe_promote_bullets_to_timeline
except Exception:
    from layout_engine import finalize_slide_layouts, maybe_promote_bullets_to_timeline


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
    if not shortened:
        shortened = candidate.strip()
    has_cjk = any("\u4e00" <= c <= "\u9fff" for c in shortened)
    if has_cjk:
        if shortened[-1] not in "。！？":
            return shortened + "。"
        return shortened
    if shortened[-1] not in ".!?":
        return shortened + "."
    return shortened


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
    if text.endswith("..."):
        base = re.sub(r"\.{2,}$", "", text).rstrip()
        cut_chars = ["。", "！", "？", ".", "!", "?", "；", ";", "：", ":", "，", ",", "、"]
        last_pos = -1
        for ch in cut_chars:
            pos = base.rfind(ch)
            if pos > last_pos:
                last_pos = pos
        if last_pos >= 0:
            return base[:last_pos + 1].strip()
        return base
    return text


def _strip_trailing_ellipsis(text: str) -> str:
    s = (text or "").strip()
    s = re.sub(r"[…\.]{1,}\s*$", "", s).strip()
    return s


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


def _clauses_from_summary(summary: str, *, min_len: int = 12, max_len: int = 68, cap: int = 3) -> list[str]:
    """Split summary into short clauses for use as supporting bullets (no LLM call)."""
    if not (summary or "").strip():
        return []
    parts = re.split(r"[。；]+", str(summary).strip())
    out = []
    for p in parts:
        p = p.strip().strip("，、 ")
        if len(p) < min_len:
            continue
        line = _clean_text(p, max_len) if len(p) > max_len else (p if p[-1] in "。！？" else p + "。")
        out.append(line)
        if len(out) >= cap:
            break
    return out


def _density_fallback_bullets(title: str, need: int) -> list[str]:
    """Honest placeholders when模型只给了过少要点（避免空泛形容词句）。"""
    short = (title or "").strip()[:24] or "本页主题"
    pool = [
        f"汇报时可补充与「{short}」相关的数据口径、统计周期或样本范围，便于听众对齐理解。",
        "建议明确责任角色、时间窗口与交付物，便于会后执行与复盘。",
        "可补充风险与前提假设：哪些结论在何种边界条件下成立。",
    ]
    return pool[:need]


def ensure_bullet_slide_density(page: dict, *, is_continued: bool) -> None:
    """正文页至少 3 条可见要点：优先从 summary 拆句，否则使用结构化占位句。"""
    if is_continued:
        return
    bullets = [str(b).strip() for b in (page.get("bullets") or []) if str(b).strip()]
    if len(bullets) >= 3:
        page["bullets"] = dedupe_preserve_order([_clean_text(b, 64) for b in bullets])[:5]
        return
    seen = {_norm_text(b) for b in bullets}
    for clause in _clauses_from_summary(page.get("summary") or ""):
        k = _norm_text(clause)
        if k and k not in seen:
            bullets.append(clause)
            seen.add(k)
        if len(bullets) >= 3:
            break
    if len(bullets) < 3:
        for fb in _density_fallback_bullets(page.get("title") or "", 3 - len(bullets)):
            k = _norm_text(fb)
            if k not in seen:
                bullets.append(fb)
                seen.add(k)
            if len(bullets) >= 3:
                break
    page["bullets"] = dedupe_preserve_order([_clean_text(b, 64) for b in bullets])[:5]


def ensure_timeline_density(slide: dict) -> None:
    """时间轴至少 3 个阶段：从 summary 拆句补入，避免大留白。"""
    steps = [str(s).strip() for s in (slide.get("timeline_steps") or slide.get("bullets") or []) if str(s).strip()]
    if len(steps) >= 3:
        slide["timeline_steps"] = dedupe_preserve_order([_clean_text(s, 72) for s in steps])[:5]
        slide["bullets"] = slide["timeline_steps"]
        return
    seen = {_norm_text(s) for s in steps}
    for clause in _clauses_from_summary(slide.get("summary") or "", min_len=10, max_len=72, cap=3):
        k = _norm_text(clause)
        if k not in seen:
            steps.append(clause)
            seen.add(k)
        if len(steps) >= 3:
            break
    if len(steps) < 3:
        for fb in _density_fallback_bullets(slide.get("title") or "", 3 - len(steps)):
            k = _norm_text(fb)
            if k not in seen:
                steps.append(_clean_text(fb, 72))
                seen.add(k)
    slide["timeline_steps"] = dedupe_preserve_order([_clean_text(s, 72) for s in steps])[:5]
    slide["bullets"] = slide["timeline_steps"]


def ensure_image_slide_density(slide: dict) -> None:
    """图文页右侧至少 2 条要点，减少『只有一句说明』的空白感。"""
    bullets = [str(b).strip() for b in (slide.get("bullets") or []) if str(b).strip()]
    if len(bullets) >= 2:
        slide["bullets"] = bullets[:4]
        return
    seen = {_norm_text(b) for b in bullets}
    cap = slide.get("image_caption") or ""
    for clause in _clauses_from_summary(slide.get("summary") or "", min_len=10, max_len=52, cap=2):
        k = _norm_text(clause)
        if k and k != _norm_text(cap) and k not in seen:
            bullets.append(_clean_text(clause, 56))
            seen.add(k)
        if len(bullets) >= 2:
            break
    if len(bullets) < 2:
        for fb in _density_fallback_bullets(slide.get("title") or "", 2 - len(bullets)):
            k = _norm_text(fb)
            if k not in seen:
                bullets.append(_clean_text(fb, 56))
                seen.add(k)
    slide["bullets"] = dedupe_preserve_order(bullets)[:4]


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
    return _strip_trailing_ellipsis(avoid_incomplete_tail(smart_shorten(text, max_len)))


def _scrub_meta_instruction_caption(text: str) -> str:
    t = (text or "").strip()
    if not t:
        return t
    if "这一页" in t[:16] and ("总结" in t[:40] or "汇总" in t[:40] or "旨在" in t[:40]):
        return ""
    if "本页" in t[:12] and ("总结" in t[:36] or "汇总" in t[:36] or "旨在" in t[:36] or "主要" in t[:36]):
        return ""
    meta_snippets = (
        "总结报告的核心要点",
        "报告的核心结论",
        "强调行动的紧迫性",
        "强调行动紧迫性",
        "本页将展示",
        "本页将",
        "如图所示",
        "以下将从",
        "详见下页",
        "本幻灯片",
    )
    for m in meta_snippets:
        if m in t and len(t) < 160:
            return ""
    return t


def _strip_meta_leading_sentence(text: str) -> str:
    """Remove opening slide-meta sentence from bullets (e.g. '这一页总结了…。')."""
    s = (text or "").strip()
    if not s:
        return s
    meta_open = (
        "这一页总结了",
        "本页总结了",
        "这一页汇总",
        "本页旨在",
        "本页主要",
        "这一页主要",
    )
    for m in meta_open:
        if s.startswith(m):
            for sep in ("。", "！", "？"):
                if sep in s:
                    idx = s.index(sep)
                    rest = s[idx + 1 :].strip(" ，。、；")
                    return rest if len(rest) > 10 else ""
            return ""
    return s


def _finalize_visible_line(text: str, max_len: int) -> str:
    t = (text or "").strip()
    if not t:
        return t
    t = _clean_text(t, max_len)
    t = re.sub(r"(\.{2,}|…+)\s*$", "", t).strip()
    t = avoid_incomplete_tail(t)
    t = _strip_trailing_ellipsis(t).strip()
    if not t:
        return t
    if t[-1] not in "。！？.!?":
        t = t + ("。" if any("\u4e00" <= c <= "\u9fff" for c in t) else ".")
    if len(t) > max_len:
        t = _clean_text(t, max_len)
    return t


def _dedupe_bullets_against_image_caption(caption: str, bullets: list) -> list:
    cap = re.sub(r"(\.{2,}|…+)\s*$", "", (caption or "").strip()).strip()
    nc = _norm_text(cap)
    out = []
    for b in bullets or []:
        s = str(b).strip()
        if not s:
            continue
        ns = _norm_text(s)
        if nc and ns == nc:
            continue
        if cap and len(cap) >= 8 and s.startswith(cap):
            rest = s[len(cap) :].lstrip(" ，。、；：．.")
            if len(rest) > 8:
                out.append(rest)
            continue
        out.append(s)
    return dedupe_preserve_order(out)


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


def _dedupe_summary_against_title(slide: dict, slide_type: str) -> None:
    if slide_type == "section":
        return
    title = (slide.get("title") or "").strip()
    summary = (slide.get("summary") or "").strip()
    if not title or not summary:
        return
    nt, ns = _norm_text(title), _norm_text(summary)
    if nt != ns and nt not in ns and ns not in nt:
        return
    bullets = slide.get("bullets") or []
    for b in bullets:
        nb = _norm_text(str(b))
        if nb and nb != nt:
            slide["summary"] = _clean_text(str(b), 96)
            return
    for key in ("left_points", "right_points"):
        for b in slide.get(key) or []:
            nb = _norm_text(str(b))
            if nb and nb != nt:
                slide["summary"] = _clean_text(str(b), 96)
                return
    slide["summary"] = ""


def _apply_layout_variant(slide: dict) -> dict:
    """Clone slide; `layout_variant` is assigned later by `finalize_slide_layouts`."""
    return dict(slide)


def _timeline_slide_fingerprint(slide: dict) -> tuple:
    steps = slide.get("timeline_steps") or slide.get("bullets") or []
    return (
        _norm_text(slide.get("title", "")),
        tuple(_norm_text(str(x)) for x in steps),
        _norm_text(slide.get("summary", "")),
    )


def dedupe_consecutive_duplicate_slides(slides: list) -> list:
    """Drop consecutive timeline slides with identical visible content (LLM often repeats the same页)."""
    out = []
    prev_tl_fp = None
    for s in slides or []:
        if not isinstance(s, dict):
            continue
        if s.get("type") == "timeline":
            fp = _timeline_slide_fingerprint(s)
            if fp == prev_tl_fp:
                continue
            prev_tl_fp = fp
        else:
            prev_tl_fp = None
        out.append(s)
    return out


def enforce_slide_content_budget(ppt_data: dict) -> dict:
    slides = ppt_data.get("slides", [])
    new_slides = []

    for slide in slides:
        slide = dict(slide)
        slide = maybe_promote_bullets_to_timeline(slide)
        slide_type = slide.get("type", "bullet")

        if "title" in slide:
            slide["title"] = _clean_text(slide.get("title", ""), 34)
        if "subtitle" in slide:
            slide["subtitle"] = _clean_text(slide.get("subtitle", ""), 56)
        if "summary" in slide:
            if slide_type != "section":
                slide["summary"] = _clean_text(slide.get("summary", ""), 96)
        if "highlight" in slide:
            slide["highlight"] = _clean_text(slide.get("highlight", ""), 78)
        if "closing" in slide:
            slide["closing"] = _clean_text(slide.get("closing", ""), 82)
        if "speaker_note" in slide:
            slide["speaker_note"] = _clean_text(slide.get("speaker_note", ""), 260)

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
                ensure_bullet_slide_density(page, is_continued=is_continued)
                _dedupe_summary_against_title(page, "bullet")
                new_slides.append(_apply_layout_variant(page))
            continue

        elif slide_type == "timeline":
            raw_steps = [str(s).strip() for s in (slide.get("timeline_steps") or slide.get("bullets") or []) if str(s).strip()]
            raw_steps = dedupe_preserve_order([_clean_text(s, 72) for s in raw_steps])
            slide["timeline_steps"] = raw_steps
            slide["bullets"] = list(raw_steps)
            ensure_timeline_density(slide)
            _dedupe_summary_against_title(slide, slide_type)
            new_slides.append(_apply_layout_variant(slide))
            continue

        elif slide_type in ["image_left", "image_right"]:
            title = slide.get("title", "") or ""
            summary = slide.get("summary", "") or ""
            caption = _scrub_meta_instruction_caption((slide.get("image_caption") or "").strip())
            bullets = []
            for b in slide.get("bullets") or []:
                if not str(b).strip():
                    continue
                line = _strip_meta_leading_sentence(_clean_text(str(b), 56))
                if line and str(line).strip():
                    bullets.append(line)
            bullets = dedupe_preserve_order(bullets)
            bullets = _remove_redundant_points(bullets, summary, caption or (slide.get("image_caption") or ""), title)

            slide["bullets"] = bullets[:4]
            if not caption:
                sn = (slide.get("speaker_note") or "").strip()
                chunk = ""
                if sn:
                    for sep in ("。", "！", "？", ".", "!", "?"):
                        if sep in sn:
                            chunk = sn.split(sep)[0].strip() + sep
                            break
                    else:
                        chunk = _clean_text(sn, 52)
                    chunk = _scrub_meta_instruction_caption(chunk)
                if chunk and _norm_text(chunk) != _norm_text(summary):
                    slide["image_caption"] = _finalize_visible_line(chunk, 56)
                else:
                    base = summary or slide.get("title", "视觉图像")
                    slide["image_caption"] = _finalize_visible_line(base, 56)
            else:
                slide["image_caption"] = _finalize_visible_line(caption, 56)
            slide["bullets"] = _dedupe_bullets_against_image_caption(slide["image_caption"], slide["bullets"])
            slide["bullets"] = [_finalize_visible_line(b, 56) for b in (slide["bullets"] or []) if str(b).strip()]
            slide["bullets"] = dedupe_preserve_order(slide["bullets"])[:4]
            ensure_image_slide_density(slide)
            slide["bullets"] = _dedupe_bullets_against_image_caption(slide["image_caption"], slide["bullets"])
            slide["bullets"] = [_finalize_visible_line(b, 56) for b in (slide["bullets"] or []) if str(b).strip()]
            slide["bullets"] = dedupe_preserve_order(slide["bullets"])[:4]
            _dedupe_summary_against_title(slide, slide_type)

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
            _dedupe_summary_against_title(slide, slide_type)

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
            _dedupe_summary_against_title(slide, slide_type)

        elif slide_type == "highlight":
            slide["summary"] = _clean_text(slide.get("summary", ""), 100)
            slide["highlight"] = _clean_text(slide.get("highlight") or slide.get("summary", ""), 72)
            slide["closing"] = _clean_text(slide.get("closing") or slide.get("summary", ""), 80)
            _dedupe_summary_against_title(slide, slide_type)

        elif slide_type == "section":
            slide["summary"] = _clean_text(slide.get("summary", "") or "", 240)
        else:
            _dedupe_summary_against_title(slide, slide_type)

        new_slides.append(_apply_layout_variant(slide))

    ppt_data["slides"] = dedupe_consecutive_duplicate_slides(new_slides)
    finalize_slide_layouts(ppt_data["slides"])
    if "subtitle" in ppt_data:
        ppt_data["subtitle"] = _clean_text(ppt_data.get("subtitle", ""), 120)
    if "summary" in ppt_data:
        ppt_data["summary"] = _clean_text(ppt_data["summary"], 120)
    return ppt_data
