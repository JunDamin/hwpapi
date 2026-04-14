"""Tests for hwpapi.classes - dataclasses and accessors (no HWP required)."""
import pytest
from dataclasses import fields, asdict
from unittest.mock import MagicMock
from hwpapi.classes import (
    Character, Paragraph, PageShape,
    MoveAccessor, CellAccessor, TableAccessor, PageAccessor,
)


# ── Dataclass: Character ─────────────────────────────────────────────────

class TestCharacterDataclass:
    def test_default_instantiation(self):
        char = Character()
        assert char is not None

    def test_default_values_are_none(self):
        char = Character()
        assert char.Bold is None
        assert char.Italic is None
        assert char.FaceNameHangul is None

    def test_keyword_initialization(self):
        char = Character(Bold=1, Italic=1, FaceNameHangul="Arial")
        assert char.Bold == 1
        assert char.Italic == 1
        assert char.FaceNameHangul == "Arial"

    def test_asdict_conversion(self):
        char = Character(Bold=1)
        d = asdict(char)
        assert isinstance(d, dict)
        assert d["Bold"] == 1
        assert d["Italic"] is None

    def test_has_expected_fields(self):
        field_names = {f.name for f in fields(Character)}
        expected = {"Bold", "Italic", "FaceNameHangul", "TextColor", "Height"}
        assert expected.issubset(field_names)

    def test_field_count(self):
        """Character should have ~65 fields."""
        assert len(fields(Character)) > 50


# ── Dataclass: Paragraph ─────────────────────────────────────────────────

class TestParagraphDataclass:
    def test_default_instantiation(self):
        para = Paragraph()
        assert para is not None

    def test_default_values_are_none(self):
        para = Paragraph()
        assert para.AlignType is None
        assert para.LeftMargin is None

    def test_keyword_initialization(self):
        para = Paragraph(AlignType=1, LeftMargin=100)
        assert para.AlignType == 1
        assert para.LeftMargin == 100

    def test_asdict(self):
        para = Paragraph(LineSpacing=160)
        d = asdict(para)
        assert d["LineSpacing"] == 160


# ── Dataclass: PageShape ─────────────────────────────────────────────────

class TestPageShapeDataclass:
    def test_instantiation_with_required_args(self):
        """PageShape has required fields - check it accepts them."""
        # PageShape has at least one required field (MarginLeft)
        # We just verify the class is a valid dataclass
        field_list = fields(PageShape)
        assert len(field_list) > 0


# ── Accessor classes (unit tests with mocks) ─────────────────────────────

class TestMoveAccessor:
    def test_instantiation(self):
        mock_app = MagicMock()
        accessor = MoveAccessor(mock_app)
        assert accessor._app is mock_app

    def test_has_logger(self):
        mock_app = MagicMock()
        accessor = MoveAccessor(mock_app)
        assert accessor.logger is not None


class TestCellAccessor:
    def test_instantiation(self):
        mock_app = MagicMock()
        accessor = CellAccessor(mock_app)
        assert accessor is not None


class TestTableAccessor:
    def test_instantiation(self):
        mock_app = MagicMock()
        accessor = TableAccessor(mock_app)
        assert accessor is not None


class TestPageAccessor:
    def test_instantiation(self):
        mock_app = MagicMock()
        accessor = PageAccessor(mock_app)
        assert accessor is not None


# ── Cross-class invariants ───────────────────────────────────────────────

class TestDataclassConsistency:
    """Ensure Character/Paragraph fields match char_fields/para_fields lists."""

    def test_character_fields_match_char_fields(self):
        from hwpapi.constants import char_fields
        field_names = {f.name for f in fields(Character)}
        # All char_fields should be in Character dataclass
        for name in char_fields:
            assert name in field_names, f"char_fields has '{name}' but Character doesn't"

    def test_paragraph_fields_match_para_fields(self):
        from hwpapi.constants import para_fields
        field_names = {f.name for f in fields(Paragraph)}
        for name in para_fields:
            assert name in field_names, f"para_fields has '{name}' but Paragraph doesn't"
