LAYOUT_VARIANTS = {
    "highlight": ["hero_center", "band_top"],
    "bullet": ["card_clean", "card_dense"],
    "compare": ["dual_cards"],
    "two_column": ["split_clean"],
    "image_left": ["image_left_soft", "image_left_bold"],
    "image_right": ["image_right_soft", "image_right_bold"],
}


def pick_layout_variant(slide: dict) -> str:
    slide_type = slide.get("type", "bullet")
    visual_priority = int(slide.get("visual_priority", 0) or 0)
    bullets = slide.get("bullets", []) or []
    is_continued = bool(slide.get("_continued"))

    if slide_type == "highlight":
        return "hero_center" if visual_priority >= 2 else "band_top"

    if slide_type == "bullet":
        if is_continued:
            return "card_dense"
        return "card_clean" if len(bullets) <= 3 else "card_dense"

    if slide_type == "compare":
        return "dual_cards"

    if slide_type == "two_column":
        return "split_clean"

    if slide_type in {"image_left", "image_right"}:
        if visual_priority >= 2 or len(bullets) <= 2:
            return f"{slide_type}_bold"
        return f"{slide_type}_soft"

    return LAYOUT_VARIANTS.get(slide_type, ["card_clean"])[0]
