"""
HWP ParameterSet value mappings (string ↔ int).

These mappings are used by ``MappedProperty`` descriptors to translate
user-friendly strings (e.g. "left", "center") to/from the integer codes
expected by the HWP COM API.

All mappings in this module are public and exported via ``__all__`` of the
parent ``hwpapi.parametersets`` package.

Usage:
    >>> from hwpapi.parametersets import DIRECTION_MAP, ALIGN_MAP
    >>> DIRECTION_MAP["left"]
    0
    >>> ALL_MAPPINGS["align"]
    {'left': 0, 'center': 1, 'right': 2}
"""
from __future__ import annotations


# ── Direction ────────────────────────────────────────────────────────────
DIRECTION_MAP = {"left": 0, "right": 1, "top": 2, "bottom": 3}

# ── Size / Alignment ─────────────────────────────────────────────────────
CAP_FULL_SIZE_MAP = {"exclude": 0, "include": 1}
ALIGNMENT_MAP = {"left": 0, "center": 1, "right": 2}
VERT_ALIGN_MAP = {"top": 0, "center": 1, "bottom": 2}
VERT_REL_TO_MAP = {"paper": 0, "page": 1, "paragraph": 2}
HORZ_REL_TO_MAP = {"paper": 0, "page": 1, "column": 2, "paragraph": 3}
HORZ_ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "inside": 3, "outside": 4}
ALIGN_MAP = {"left": 0, "center": 1, "right": 2}
ALIGN_TYPE_MAP = {"between": 0, "left": 1, "right": 2, "center": 3, "ratio": 4, "shared": 5}

# ── Font / Text ──────────────────────────────────────────────────────────
FONTTYPE_MAP = {"don't care": 0, "TTF": 1, "HFT": 2, "dontcare": 0, "ttf": 1, "htf": 2}
TEXT_DIRECTION_MAP = {"horizontal": 0, "vertical": 1}
TEXT_ALIGN_MAP = {"font": 0, "up": 1, "middle": 2, "down": 3}
LINE_SPACING_TYPE_MAP = {"font": 0, "fixed": 1, "space": 2}

# ── Line / Text Wrapping ─────────────────────────────────────────────────
LINE_WRAP_MAP = {"basic": 0, "no_newline": 1, "forced": 2}
TEXT_WRAP_MAP = {"square": 0, "top_bottom": 1, "behind": 2, "front": 3, "tight": 4, "through": 5}
TEXT_FLOW_MAP = {"both": 0, "left": 1, "right": 2, "largest": 3}
LATIN_LINE_BREAK_MAP = {"word": 0, "hyphen": 1, "letter": 2}
NONLATIN_LINE_BREAK_MAP = {"word": 0, "letter": 1}

# ── Style / Effect ───────────────────────────────────────────────────────
SHADOW_TYPE_MAP = {"none": 0, "drop": 1, "continuous": 2}
BACKGROUND_TYPE_MAP = {"empty": 0, "fill": 1, "picture": 2, "gradation": 3}
GRADATION_TYPE_MAP = {"stripe": 1, "circle": 2, "cone": 3, "square": 4}
ROTATION_SETTING_MAP = {"none": 0, "setted_rotation": 1, "picture_centered_rotation": 2, "rotation_and_centered": 3}
PIC_EFFECT_MAP = {"none": 0, "bw": 1, "sepia": 2}

# ── Search / Direction ───────────────────────────────────────────────────
SEARCH_DIRECTION_MAP = {"down": 0, "up": 1, "doc": 2}

# ── Border / Outline ─────────────────────────────────────────────────────
BORDER_TEXT_MAP = {"column": 0, "text": 1}
UNDERLINE_TYPE_MAP = {"none": 0, "bottom": 1, "center": 2, "top": 3}
OUTLINE_TYPE_MAP = {
    "none": 0, "solid": 1, "dot": 2, "thick": 3,
    "dash": 4, "dashdot": 5, "dashdotdot": 6,
}
STRIKEOUT_TYPE_MAP = {
    "none": 0, "red single": 1, "red double": 2,
    "text single": 3, "text double": 4,
}

# ── Special Character / Formatting ───────────────────────────────────────
USE_KERNING_MAP = {"off": 0, "on": 1}
DIAC_SYM_MARK_MAP = {"none": 0, "black circle": 1, "empty circle": 2}
USE_FONT_SPACE_MAP = {"off": 0, "on": 1}
HEADING_TYPE_MAP = {"none": 0, "outline": 1, "number": 2, "bullet": 3}

# ── Numbering / Formatting ───────────────────────────────────────────────
NUMBERING_TYPE_MAP = {"none": 0, "picture": 1, "table": 2, "equation": 3}
NUMBER_FORMAT_MAP = {
    "1": 0, "circled 1": 1, "I": 2, "i": 3, "A": 4, "a": 5,
    "circled A": 6, "circled a": 7, "가": 8, "동그라미 가": 9,
    "ㄱ": 10, "동그라미 ㄱ": 11, "일": 12, "一": 13, "동그라미 一": 14
}

# ── Page / Table ─────────────────────────────────────────────────────────
PAGE_BREAK_MAP = {"none": 0, "cell": 1, "text": 2}


# ── Unified registry ─────────────────────────────────────────────────────
ALL_MAPPINGS = {
    "direction": DIRECTION_MAP,
    "cap_full_size": CAP_FULL_SIZE_MAP,
    "alignment": ALIGNMENT_MAP,
    "fonttype": FONTTYPE_MAP,
    "shadow_type": SHADOW_TYPE_MAP,
    "background_type": BACKGROUND_TYPE_MAP,
    "gradation_type": GRADATION_TYPE_MAP,
    "rotation_setting": ROTATION_SETTING_MAP,
    "search_direction": SEARCH_DIRECTION_MAP,
    "text_direction": TEXT_DIRECTION_MAP,
    "line_wrap": LINE_WRAP_MAP,
    "text_wrap": TEXT_WRAP_MAP,
    "text_flow": TEXT_FLOW_MAP,
    "vert_align": VERT_ALIGN_MAP,
    "vert_rel_to": VERT_REL_TO_MAP,
    "horz_rel_to": HORZ_REL_TO_MAP,
    "horz_align": HORZ_ALIGN_MAP,
    "align": ALIGN_MAP,
    "align_type": ALIGN_TYPE_MAP,
    "latin_line_break": LATIN_LINE_BREAK_MAP,
    "nonlatin_line_break": NONLATIN_LINE_BREAK_MAP,
    "text_align": TEXT_ALIGN_MAP,
    "heading_type": HEADING_TYPE_MAP,
    "border_text": BORDER_TEXT_MAP,
    "underline_type": UNDERLINE_TYPE_MAP,
    "outline_type": OUTLINE_TYPE_MAP,
    "strikeout_type": STRIKEOUT_TYPE_MAP,
    "use_kerning": USE_KERNING_MAP,
    "diac_sym_mark": DIAC_SYM_MARK_MAP,
    "use_font_space": USE_FONT_SPACE_MAP,
    "numbering_type": NUMBERING_TYPE_MAP,
    "number_format": NUMBER_FORMAT_MAP,
    "line_spacing_type": LINE_SPACING_TYPE_MAP,
    "pic_effect": PIC_EFFECT_MAP,
    "page_break": PAGE_BREAK_MAP,
}


# ════════════════════════════════════════════════════════════════════
# v0.0.25+ — ENUM 통일 (BorderType, HatchStyle, ApplyTo, DiagonalFlag)
# ════════════════════════════════════════════════════════════════════

#: Border / line type — BorderFill.BorderType*, set_cell_border 의 top/bottom
BORDER_TYPE_MAP = {
    "none": 0,                 # 선 없음
    "solid": 1,                # 실선
    "dash": 2,                 # 짧은 점선
    "dot": 3,                  # 작은 점
    "dash_dot": 4,             # 1점쇄선
    "dash_dot_dot": 5,         # 2점쇄선
    "long_dash": 6,            # 긴 점선
    "circle": 7,               # 원형 점
    "double": 8,               # 이중선
    "thick_thin": 9,           # 얇음+굵음
    "thin_thick": 10,          # 굵음+얇음
    "thin_thick_thin": 11,     # 얇음+굵음+얇음
    "wave": 12,                # 물결
    "double_wave": 13,         # 이중 물결
    "thick_3d": 14,            # 두꺼운 3D
    "thick_3d_inset": 15,      # 두꺼운 3D inset
}

#: Hatch style — set_cell_color 의 hatch_style, FillAttr.WinBrushFaceStyle
HATCH_STYLE_MAP = {
    "none": 0,                 # 채우기 없음
    "horizontal": 1,
    "vertical": 2,
    "diagonal_down": 3,
    "diagonal_up": 4,
    "cross": 5,
    "diagonal_cross": 6,
    "dense_horizontal": 7,
    "dense_vertical": 8,
    "solid": 9,
    "dense_diagonal_down": 10,
    "dense_diagonal_up": 11,
    "dense_cross": 12,
}

#: Cell ApplyTo — CellBorderFill 등의 ApplyTo
CELL_APPLY_TO_MAP = {
    "current": 1,              # 현재 셀만
    "selected": 2,             # 선택된 셀들
    "all": 3,                  # 표 전체
}

#: Diagonal flag (bit-flag) — BorderFill.SlashFlag / CounterSlashFlag
DIAGONAL_FLAG_MAP = {
    "none": 0,
    "bottom": 1,
    "middle": 2,
    "top": 4,
    "all": 7,                  # bottom | middle | top
}


def resolve_enum(map_dict: dict, value):
    """
    Enum/MAP 값 해석 — 문자열이면 dict 조회, 정수면 그대로 반환.

    Examples
    --------
    >>> resolve_enum(BORDER_TYPE_MAP, "solid")
    1
    >>> resolve_enum(BORDER_TYPE_MAP, 1)
    1
    >>> resolve_enum(BORDER_TYPE_MAP, "double")
    8
    """
    if isinstance(value, str):
        if value in map_dict:
            return map_dict[value]
        # case-insensitive fallback
        lower = {k.lower(): v for k, v in map_dict.items()}
        if value.lower() in lower:
            return lower[value.lower()]
        raise ValueError(
            f"Unknown enum value {value!r}. Valid: {sorted(map_dict.keys())}"
        )
    return value


__all__ = [
    "PARAMETERSET_REGISTRY",  # re-declared in __init__.py, just for clarity
    "DIRECTION_MAP", "CAP_FULL_SIZE_MAP", "ALIGNMENT_MAP", "VERT_ALIGN_MAP",
    "VERT_REL_TO_MAP", "HORZ_REL_TO_MAP", "HORZ_ALIGN_MAP", "ALIGN_MAP",
    "ALIGN_TYPE_MAP", "FONTTYPE_MAP", "TEXT_DIRECTION_MAP", "TEXT_ALIGN_MAP",
    "LINE_SPACING_TYPE_MAP", "LINE_WRAP_MAP", "TEXT_WRAP_MAP", "TEXT_FLOW_MAP",
    "LATIN_LINE_BREAK_MAP", "NONLATIN_LINE_BREAK_MAP", "SHADOW_TYPE_MAP",
    "BACKGROUND_TYPE_MAP", "GRADATION_TYPE_MAP", "ROTATION_SETTING_MAP",
    "PIC_EFFECT_MAP", "SEARCH_DIRECTION_MAP", "BORDER_TEXT_MAP",
    "UNDERLINE_TYPE_MAP", "OUTLINE_TYPE_MAP", "STRIKEOUT_TYPE_MAP",
    "USE_KERNING_MAP", "DIAC_SYM_MARK_MAP", "USE_FONT_SPACE_MAP",
    "HEADING_TYPE_MAP", "NUMBERING_TYPE_MAP", "NUMBER_FORMAT_MAP",
    "PAGE_BREAK_MAP", "ALL_MAPPINGS",
    # v0.0.25+
    "BORDER_TYPE_MAP", "HATCH_STYLE_MAP", "CELL_APPLY_TO_MAP",
    "DIAGONAL_FLAG_MAP", "resolve_enum",
]
# Note: PARAMETERSET_REGISTRY is defined in __init__.py, not here
__all__.remove("PARAMETERSET_REGISTRY")
