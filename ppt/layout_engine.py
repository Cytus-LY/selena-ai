"""
Layout selection for PPT slides: structural signals + deck-level variety.

`layout_variant` is consumed by renderers to pick geometry and chrome (cards vs plain,
image hero ratio, compare panels vs open columns, etc.).
"""
from __future__ import annotations

import re
from typing import Any

# --- Structural signals (content-based, not title-only heuristics) ------------

_TIMELINE_RE = re.compile(
    r"(阶段|里程碑|路线图|时间线|排期|实施计划|落地计划|滚动|迭代)"
    r"|(第[一二三四五六七八九十百千万\d]+[步周天月季])"
    r"|(Q[1-4]|H[12]|P\d+|S\d+|M\d+)"
    r"|(\d{4}\s*年|\d{1,2}\s*月)"
    r"|(T\+\d+|T-\d+|D\+\d+|Week\s*\d+)",
    re.I,
)


def _joined_bullets(slide: dict) -> str:
    parts = [str(slide.get("title") or ""), str(slide.get("summary") or "")]
    for b in slide.get("bullets") or []:
        parts.append(str(b))
    return " ".join(parts)


def slide_suggests_timeline(slide: dict) -> bool:
    """True when bullet lines collectively look like phases / schedule (not keyword on title only)."""
    if (slide.get("type") or "bullet") != "bullet":
        return False
    bullets = [str(b).strip() for b in (slide.get("bullets") or []) if str(b).strip()]
    if len(bullets) < 3:
        return False
    hits = sum(1 for b in bullets if _TIMELINE_RE.search(b))
    if hits >= 2:
        return True
    if hits >= 1 and _TIMELINE_RE.search(_joined_bullets(slide)):
        return len(bullets) <= 5
    return False


def maybe_promote_bullets_to_timeline(slide: dict) -> dict:
    """Upgrade bullet → timeline when content structure matches a plan / roadmap."""
    slide = dict(slide)
    if not slide_suggests_timeline(slide):
        return slide
    bullets = [str(b).strip() for b in (slide.get("bullets") or []) if str(b).strip()]
    slide["type"] = "timeline"
    slide["timeline_steps"] = bullets[:5]
    slide["bullets"] = bullets[:5]
    return slide


def _count_compare_pairs(slide: dict) -> tuple[int, int]:
    """How many structured points exist on left/right (for compare / two_column)."""
    lp = [str(x).strip() for x in (slide.get("left_points") or []) if str(x).strip()]
    rp = [str(x).strip() for x in (slide.get("right_points") or []) if str(x).strip()]
    return len(lp), len(rp)


def build_layout_context(slides: list[dict]) -> list[dict[str, Any]]:
    """Per-slide context: ordinals within type + previous variant for alternation."""
    bullet_i = 0
    section_i = 0
    compare_i = 0
    two_i = 0
    image_i = 0
    timeline_i = 0
    highlight_i = 0
    prev_bullet_variant: str | None = None
    out: list[dict[str, Any]] = []
    for idx, raw in enumerate(slides or []):
        st = (raw.get("type") or "bullet") if isinstance(raw, dict) else "bullet"
        ctx: dict[str, Any] = {
            "deck_index": idx,
            "slide_type": st,
            "prev_bullet_variant": prev_bullet_variant,
        }
        if st == "bullet":
            bullet_i += 1
            ctx["bullet_ordinal"] = bullet_i
        if st == "section":
            section_i += 1
            ctx["section_ordinal"] = section_i
        if st == "compare":
            compare_i += 1
            ctx["compare_ordinal"] = compare_i
        if st == "two_column":
            two_i += 1
            ctx["two_column_ordinal"] = two_i
        if st in {"image_left", "image_right"}:
            image_i += 1
            ctx["image_ordinal"] = image_i
        if st == "timeline":
            timeline_i += 1
            ctx["timeline_ordinal"] = timeline_i
        if st == "highlight":
            highlight_i += 1
            ctx["highlight_ordinal"] = highlight_i
        out.append(ctx)
        # update prev after we will compute variant — done in finalize pass
    return out


def pick_layout_variant(slide: dict, ctx: dict[str, Any] | None = None) -> str:
    """
    Return a renderer key. Backward-compatible names (card_clean, …) are still produced
    where old code might expect them; new keys extend the set.
    """
    ctx = ctx or {}
    slide_type = slide.get("type", "bullet")
    visual_priority = int(slide.get("visual_priority", 0) or 0)
    bullets = slide.get("bullets", []) or []
    is_continued = bool(slide.get("_continued"))

    if slide_type == "highlight":
        ho = int(ctx.get("highlight_ordinal", 1) or 1)
        if visual_priority >= 2:
            return "hero_center"
        return "band_top" if ho % 2 == 0 else "hero_center"

    if slide_type == "timeline":
        return "timeline_rail"

    if slide_type == "bullet":
        n = len(bullets)
        total_chars = sum(len(str(b)) for b in bullets)
        ordi = int(ctx.get("bullet_ordinal", 0) or 0)
        prev = ctx.get("prev_bullet_variant")
        sparse = n <= 2 or total_chars < 95

        if is_continued:
            return "bullet_plain_bar"

        # 内容少：避免大卡片 + 优先紧凑条带/竖线版式
        if sparse or (n <= 3 and total_chars < 130):
            return "bullet_plain_bar"

        # 中等信息量：用小卡片而非满幅标准卡
        if n <= 3 and total_chars < 220:
            return "bullet_card_compact"

        # Variety: every 2nd bullet page uses two-column when enough lines
        if n >= 4 and ordi % 2 == 0:
            return "bullet_plain_two_col"

        if ordi % 3 == 0 and prev not in ("bullet_plain_bar", "bullet_plain_two_col"):
            return "bullet_plain_bar"

        if n <= 3:
            return "bullet_card_compact"
        return "bullet_card_standard"

    if slide_type == "compare":
        lc, rc = _count_compare_pairs(slide)
        # Strong structured contrast → open columns read better
        if lc >= 3 and rc >= 3 and int(ctx.get("compare_ordinal", 1) or 1) % 2 == 0:
            return "compare_open"
        return "compare_panels"

    if slide_type == "two_column":
        if int(ctx.get("two_column_ordinal", 1) or 1) % 2 == 0:
            return "two_column_open"
        return "two_column_cards"

    if slide_type in {"image_left", "image_right"}:
        # Hero ratio + subtle frame; bold = stronger divider on text side
        if visual_priority >= 2 or len(bullets) <= 2:
            return "image_hero_bold"
        return "image_hero_soft"

    if slide_type == "section":
        # Alternate centered “card” section vs editorial strip (less big blue block)
        so = int(ctx.get("section_ordinal", 1) or 1)
        return "section_editorial" if so % 2 == 1 else "section_center_card"

    # Fallback
    return "bullet_card_standard"


def finalize_slide_layouts(slides: list[dict]) -> None:
    """
    Mutates slides in place: sets layout_variant using deck-level context
    (alternation, ordinals). Updates prev_bullet_variant as we go.
    """
    if not slides:
        return
    ctx_list = build_layout_context(slides)
    prev_bullet_variant: str | None = None
    for i, slide in enumerate(slides):
        if not isinstance(slide, dict):
            continue
        ctx = dict(ctx_list[i])
        ctx["prev_bullet_variant"] = prev_bullet_variant
        slide["layout_variant"] = pick_layout_variant(slide, ctx)
        lv = slide.get("layout_variant") or ""
        if slide.get("type") == "bullet" or lv.startswith("bullet_"):
            prev_bullet_variant = lv


# Legacy map (for documentation / tests)
LAYOUT_VARIANTS = {
    "highlight": ["hero_center", "band_top"],
    "bullet": ["bullet_card_standard", "bullet_card_compact", "bullet_plain_bar", "bullet_plain_two_col"],
    "compare": ["compare_panels", "compare_open"],
    "two_column": ["two_column_cards", "two_column_open"],
    "image_left": ["image_hero_soft", "image_hero_bold"],
    "image_right": ["image_hero_soft", "image_hero_bold"],
    "timeline": ["timeline_rail"],
    "section": ["section_center_card", "section_editorial"],
}
