"""Unit tests for hwpapi.functions utility functions (no HWP required)."""
import pytest
from hwpapi.functions import (
    hex_to_rgb, get_rgb_tuple, convert_to_hwp_color, convert_hwp_color_to_hex,
    mili2unit, unit2mili, point2unit, unit2point,
    parse_unit_string, to_hwpunit, from_hwpunit,
    get_value, get_key, convert2int, get_font_name,
    HWPUNIT_PER_MM, HWPUNIT_PER_PT, HWPUNIT_PER_CM, HWPUNIT_PER_INCH,
)


# ── Unit constants ───────────────────────────────────────────────────────

class TestUnitConstants:
    def test_mm_constant(self):
        assert HWPUNIT_PER_MM == 283

    def test_pt_constant(self):
        assert HWPUNIT_PER_PT == 100

    def test_cm_constant(self):
        assert HWPUNIT_PER_CM == 2830

    def test_inch_constant(self):
        assert HWPUNIT_PER_INCH == 7200


# ── Basic unit conversion (legacy API) ──────────────────────────────────

class TestLegacyUnitConversion:
    def test_mili2unit_basic(self):
        assert mili2unit(1) == 283
        assert mili2unit(10) == 2830

    def test_mili2unit_zero(self):
        # Zero is falsy, returns as-is
        assert mili2unit(0) == 0

    def test_mili2unit_none(self):
        assert mili2unit(None) is None

    def test_unit2mili_basic(self):
        assert unit2mili(283) == 1.0
        assert unit2mili(2830) == 10.0

    def test_point2unit_basic(self):
        assert point2unit(1) == 100
        assert point2unit(12) == 1200

    def test_unit2point_basic(self):
        assert unit2point(100) == 1.0
        assert unit2point(1200) == 12.0

    def test_roundtrip_mm(self):
        for val in [1, 5, 10, 100, 210, 297]:
            assert unit2mili(mili2unit(val)) == val


# ── parse_unit_string ────────────────────────────────────────────────────

class TestParseUnitString:
    def test_mm_suffix(self):
        assert parse_unit_string("210mm") == (210.0, "mm")

    def test_cm_suffix(self):
        assert parse_unit_string("21cm") == (21.0, "cm")

    def test_inch_suffix(self):
        assert parse_unit_string("8.27in") == (8.27, "in")

    def test_pt_suffix(self):
        assert parse_unit_string("12pt") == (12.0, "pt")

    def test_bare_number_uses_default(self):
        assert parse_unit_string(210) == (210.0, "mm")
        assert parse_unit_string(12, default_unit="pt") == (12.0, "pt")

    def test_float_input(self):
        assert parse_unit_string(12.5) == (12.5, "mm")


# ── to_hwpunit / from_hwpunit ────────────────────────────────────────────

class TestHwpUnitConversion:
    def test_mm_to_hwpunit(self):
        assert to_hwpunit("210mm") == 210 * 283
        assert to_hwpunit(210, "mm") == 210 * 283

    def test_cm_to_hwpunit(self):
        assert to_hwpunit("21cm") == 21 * 2830

    def test_inch_to_hwpunit(self):
        assert to_hwpunit("1in") == 7200

    def test_pt_to_hwpunit(self):
        assert to_hwpunit("12pt") == 1200

    def test_from_hwpunit_mm(self):
        assert from_hwpunit(59430, "mm") == pytest.approx(210.0)

    def test_from_hwpunit_pt(self):
        assert from_hwpunit(1200, "pt") == 12.0

    def test_roundtrip_mm(self):
        for mm in [10, 50, 100, 210, 297]:
            assert from_hwpunit(to_hwpunit(f"{mm}mm"), "mm") == pytest.approx(mm)


# ── Color conversion ─────────────────────────────────────────────────────

class TestColorConversion:
    def test_hex_to_rgb_with_hash(self):
        assert hex_to_rgb("#FF0000") == (255, 0, 0)

    def test_hex_to_rgb_without_hash(self):
        assert hex_to_rgb("00FF00") == (0, 255, 0)

    def test_hex_to_rgb_blue(self):
        assert hex_to_rgb("#0000FF") == (0, 0, 255)

    def test_get_rgb_tuple_passthrough(self):
        assert get_rgb_tuple((100, 150, 200)) == (100, 150, 200)

    def test_get_rgb_tuple_invalid_length(self):
        with pytest.raises(ValueError):
            get_rgb_tuple((1, 2, 3, 4))

    def test_get_rgb_tuple_invalid_value(self):
        with pytest.raises(ValueError):
            get_rgb_tuple((300, 0, 0))

    def test_get_rgb_tuple_named_color(self):
        assert get_rgb_tuple("red") == (255, 0, 0)
        assert get_rgb_tuple("blue") == (0, 0, 255)

    def test_get_rgb_tuple_hex(self):
        assert get_rgb_tuple("#FF0000") == (255, 0, 0)

    def test_get_rgb_tuple_invalid_string(self):
        with pytest.raises(ValueError):
            get_rgb_tuple("not_a_color")

    def test_get_rgb_tuple_invalid_type(self):
        with pytest.raises(TypeError):
            get_rgb_tuple(123)

    def test_convert_to_hwp_color_passthrough_int(self):
        assert convert_to_hwp_color(0xFF0000) == 0xFF0000

    def test_convert_to_hwp_color_named(self):
        # "red" in HWP = 0x0000FF (BBGGRR) = 255
        assert convert_to_hwp_color("red") == 0x0000FF

    def test_convert_to_hwp_color_hex(self):
        # #FF0000 → BGR → 0x0000FF
        assert convert_to_hwp_color("#FF0000") == 0x0000FF

    def test_convert_to_hwp_color_tuple(self):
        # (R, G, B) = (255, 0, 0) → B*65536 + G*256 + R = 255
        assert convert_to_hwp_color((255, 0, 0)) == 255

    def test_convert_hwp_color_to_hex(self):
        # 0x0000FF (BBGGRR: red) → #ff0000
        assert convert_hwp_color_to_hex(0x0000FF) == "#ff0000"

    def test_convert_hwp_color_zero(self):
        assert convert_hwp_color_to_hex(0) == 0

    def test_color_roundtrip(self):
        # hex → hwp → hex should be identity (lowercase)
        original = "#ff00aa"
        hwp = convert_to_hwp_color(original)
        back = convert_hwp_color_to_hex(hwp)
        assert back == original


# ── Dict helpers ─────────────────────────────────────────────────────────

class TestDictHelpers:
    def test_get_value_basic(self):
        d = {"a": 1, "b": 2}
        assert get_value(d, "a") == 1
        assert get_value(d, "b") == 2

    def test_get_value_missing(self):
        d = {"a": 1}
        # get_value may raise or return None for missing keys
        try:
            result = get_value(d, "missing")
            assert result is None or result == "missing"
        except KeyError:
            pass  # acceptable

    def test_get_key_basic(self):
        d = {"a": 1, "b": 2}
        assert get_key(d, 1) == "a"
        assert get_key(d, 2) == "b"

    def test_convert2int_from_int(self):
        d = {"left": 0, "center": 1, "right": 2}
        assert convert2int(d, 1) == 1
        assert convert2int(d, 2) == 2

    def test_convert2int_from_string(self):
        d = {"left": 0, "center": 1, "right": 2}
        assert convert2int(d, "left") == 0
        assert convert2int(d, "center") == 1


# ── Regex helpers ────────────────────────────────────────────────────────

class TestRegexHelpers:
    def test_get_font_name_with_hft(self):
        assert get_font_name("Arial A1B2C3.HFT") == "Arial"

    def test_get_font_name_korean(self):
        assert get_font_name("바탕 1234.HFT") == "바탕"

    def test_get_font_name_no_match(self):
        assert get_font_name("plain text") is None

    def test_get_font_name_empty(self):
        assert get_font_name("") is None
