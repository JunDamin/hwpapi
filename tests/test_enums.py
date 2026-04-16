"""Test v0.0.25+ enum maps and Color class constants."""
import pytest

from hwpapi.parametersets.mappings import (
    BORDER_TYPE_MAP,
    HATCH_STYLE_MAP,
    CELL_APPLY_TO_MAP,
    DIAGONAL_FLAG_MAP,
    resolve_enum,
)
from hwpapi.parametersets.properties import Color


# ─── BORDER_TYPE_MAP ──────────────────────────────────────────────

def test_border_type_map_basic_keys():
    assert BORDER_TYPE_MAP["solid"] == 1
    assert BORDER_TYPE_MAP["double"] == 8
    assert BORDER_TYPE_MAP["wave"] == 12
    assert BORDER_TYPE_MAP["none"] == 0


def test_border_type_map_all_unique_values():
    """0~15 까지 모든 값이 고유해야 함."""
    values = list(BORDER_TYPE_MAP.values())
    assert len(values) == len(set(values))


# ─── HATCH_STYLE_MAP ──────────────────────────────────────────────

def test_hatch_style_map():
    assert HATCH_STYLE_MAP["none"] == 0
    assert HATCH_STYLE_MAP["horizontal"] == 1
    assert HATCH_STYLE_MAP["diagonal_cross"] == 6
    assert HATCH_STYLE_MAP["solid"] == 9


# ─── CELL_APPLY_TO_MAP ────────────────────────────────────────────

def test_cell_apply_to_map():
    assert CELL_APPLY_TO_MAP["current"] == 1
    assert CELL_APPLY_TO_MAP["selected"] == 2
    assert CELL_APPLY_TO_MAP["all"] == 3


# ─── DIAGONAL_FLAG_MAP ────────────────────────────────────────────

def test_diagonal_flag_map_bitwise():
    """bottom | middle | top == all."""
    assert DIAGONAL_FLAG_MAP["bottom"] | DIAGONAL_FLAG_MAP["middle"] | DIAGONAL_FLAG_MAP["top"] \
        == DIAGONAL_FLAG_MAP["all"]


# ─── resolve_enum ─────────────────────────────────────────────────

def test_resolve_enum_string():
    assert resolve_enum(BORDER_TYPE_MAP, "solid") == 1
    assert resolve_enum(BORDER_TYPE_MAP, "double") == 8


def test_resolve_enum_int_passthrough():
    assert resolve_enum(BORDER_TYPE_MAP, 5) == 5
    assert resolve_enum(BORDER_TYPE_MAP, 0) == 0


def test_resolve_enum_case_insensitive():
    assert resolve_enum(BORDER_TYPE_MAP, "SOLID") == 1
    assert resolve_enum(BORDER_TYPE_MAP, "Double") == 8


def test_resolve_enum_unknown_raises():
    with pytest.raises(ValueError):
        resolve_enum(BORDER_TYPE_MAP, "not_a_real_style")


# ─── Color class constants ────────────────────────────────────────

def test_color_red_class_constant():
    assert Color.RED.hex == "#ff0000"


def test_color_blue_class_constant():
    assert Color.BLUE.hex == "#0000ff"


def test_color_yellow_class_constant():
    assert Color.YELLOW.hex == "#ffff00"


def test_color_all_16_constants_present():
    """16 표준 색상 모두 정의됨."""
    expected = ["RED", "GREEN", "BLUE", "BLACK", "WHITE",
                 "YELLOW", "CYAN", "MAGENTA", "ORANGE", "PURPLE",
                 "PINK", "BROWN", "GRAY", "LIGHT_GRAY", "DARK_GRAY", "NAVY"]
    for name in expected:
        assert hasattr(Color, name), f"Color.{name} missing"
        assert isinstance(getattr(Color, name), Color)


def test_color_class_constants_independent():
    """클래스 상수는 매번 같은 값이지만 새 인스턴스 동일 비교 가능."""
    assert Color.RED == Color.from_hex("#FF0000")
    assert Color.RED.raw == Color.from_rgb(255, 0, 0).raw
