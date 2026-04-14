"""Tests for hwpapi.constants - enums, font lists, field lists."""
import pytest
from enum import Enum
from hwpapi.constants import (
    char_fields, para_fields,
    korean_fonts, english_fonts, chinese_fonts, japanese_fonts,
    symbol_fonts, user_fonts, other_fonts,
    MaskOption, ScanStartPosition, ScanEndPosition, ScanDirection,
    MoveId, SizeOption, Effect, SelectionOption, Direction, Thickness,
)


# ── Field lists ──────────────────────────────────────────────────────────

class TestFieldLists:
    def test_char_fields_is_list(self):
        assert isinstance(char_fields, list)
        assert len(char_fields) > 0

    def test_char_fields_no_duplicates(self):
        assert len(char_fields) == len(set(char_fields))

    def test_char_fields_contains_core_attributes(self):
        for name in ["Bold", "Italic", "FaceNameHangul", "Height"]:
            assert name in char_fields, f"{name} missing from char_fields"

    def test_para_fields_is_list(self):
        assert isinstance(para_fields, list)
        assert len(para_fields) > 0

    def test_para_fields_no_duplicates(self):
        assert len(para_fields) == len(set(para_fields))

    def test_char_fields_all_strings(self):
        assert all(isinstance(f, str) for f in char_fields)

    def test_para_fields_all_strings(self):
        assert all(isinstance(f, str) for f in para_fields)


# ── Font lists ───────────────────────────────────────────────────────────

class TestFontLists:
    @pytest.mark.parametrize("font_list,name", [
        (korean_fonts, "korean_fonts"),
        (english_fonts, "english_fonts"),
        (chinese_fonts, "chinese_fonts"),
        (japanese_fonts, "japanese_fonts"),
        (symbol_fonts, "symbol_fonts"),
        (user_fonts, "user_fonts"),
        (other_fonts, "other_fonts"),
    ])
    def test_font_list_is_list(self, font_list, name):
        assert isinstance(font_list, list), f"{name} should be a list"

    @pytest.mark.parametrize("font_list,name", [
        (korean_fonts, "korean_fonts"),
        (english_fonts, "english_fonts"),
        (chinese_fonts, "chinese_fonts"),
        (japanese_fonts, "japanese_fonts"),
        (symbol_fonts, "symbol_fonts"),
    ])
    def test_font_list_all_strings(self, font_list, name):
        assert all(isinstance(f, str) for f in font_list), f"{name} should contain only strings"


# ── Enum classes ─────────────────────────────────────────────────────────

class TestMaskOption:
    def test_is_enum(self):
        assert issubclass(MaskOption, Enum)

    def test_normal_value(self):
        assert MaskOption.Normal.value == 0x00

    def test_char_value(self):
        assert MaskOption.Char.value == 0x01

    def test_has_all_values(self):
        names = [m.name for m in MaskOption]
        assert "Normal" in names
        assert "Char" in names


class TestScanPosition:
    def test_scan_start_current(self):
        assert ScanStartPosition.Current.value == 0x0000

    def test_scan_start_document(self):
        assert ScanStartPosition.Document.value == 0x0070

    def test_scan_end_current(self):
        assert ScanEndPosition.Current.value == 0x0000

    def test_scan_end_document(self):
        assert ScanEndPosition.Document.value == 0x0007

    def test_scan_direction_forward(self):
        assert ScanDirection.Forward.value == 0x0000

    def test_scan_direction_backward(self):
        assert ScanDirection.Backward.value == 0x0100


class TestMoveId:
    def test_main_value(self):
        assert MoveId.Main.value == 0

    def test_top_of_file_value(self):
        assert MoveId.TopOfFile.value == 2

    def test_cell_offset(self):
        """Cell movements should start at 100."""
        assert MoveId.LeftOfCell.value == 100
        assert MoveId.RightOfCell.value == 101

    def test_screen_pos_offset(self):
        """Screen positions start at 200."""
        assert MoveId.ScrPos.value == 200
        assert MoveId.ScanPos.value == 201


class TestSizeOption:
    def test_real_size(self):
        assert SizeOption.RealSize.value == 0

    def test_specific_size(self):
        assert SizeOption.SpecificSize.value == 1

    def test_cell_size(self):
        assert SizeOption.CellSize.value == 2


class TestEffect:
    def test_real_picture(self):
        assert Effect.RealPicture.value == 0

    def test_grayscale(self):
        assert Effect.GrayScale.value == 1

    def test_black_white(self):
        assert Effect.BlackWhite.value == 2


class TestSelectionOption:
    def test_doc_selection(self):
        assert SelectionOption.Doc.value == ("MoveDocBegin", "MoveSelDocEnd")

    def test_para_selection(self):
        assert SelectionOption.Para.value == ("MoveParaBegin", "MoveSelParaEnd")

    def test_line_selection(self):
        assert SelectionOption.Line.value == ("MoveLineBegin", "MoveSelLineEnd")

    def test_word_selection(self):
        assert SelectionOption.Word.value == ("MoveWordBegin", "MoveSelWordEnd")


class TestDirection:
    def test_forward(self):
        assert Direction.Forward.value == 0

    def test_backward(self):
        assert Direction.Backward.value == 1

    def test_all(self):
        assert Direction.All.value == 2


class TestThickness:
    def test_is_enum(self):
        assert issubclass(Thickness, Enum)

    def test_0_1_mm(self):
        assert Thickness._0_1_mm.value == 0

    def test_max_value(self):
        assert Thickness.최대값.value == 16

    def test_has_description_method(self):
        assert hasattr(Thickness, 'get_thickness_description')


# ── Enum integrity ───────────────────────────────────────────────────────

class TestEnumIntegrity:
    @pytest.mark.parametrize("enum_cls", [
        MaskOption, ScanStartPosition, ScanEndPosition, ScanDirection,
        MoveId, SizeOption, Effect, SelectionOption, Direction, Thickness,
    ])
    def test_enum_is_subclass(self, enum_cls):
        assert issubclass(enum_cls, Enum)

    @pytest.mark.parametrize("enum_cls", [
        MaskOption, ScanStartPosition, ScanEndPosition, ScanDirection,
        SizeOption, Effect, Direction,
    ])
    def test_enum_has_members(self, enum_cls):
        assert len(list(enum_cls)) > 0
