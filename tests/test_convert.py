"""Domain-grouped tests: convert."""
from hwpapi.classes.convert import Convert, _int_to_korean
from hwpapi.classes.view import View
from hwpapi.presets import Presets
from unittest.mock import MagicMock
import pytest


@pytest.mark.parametrize("n,expected", [
    (0, "영"),
    (1, "일"),
    (10, "십"),
    (100, "백"),
    (1000, "천"),
    (10000, "일만"),
    (1234, "천이백삼십사"),
    (100000000, "일억"),
    (-5, "음오"),
])
def test_int_to_korean(n, expected):
    assert _int_to_korean(n) == expected


def test_number_to_korean_with_text():
    app = MagicMock()
    result = Convert(app).number_to_korean("1,234,567")
    assert "백이십삼만" in result
    assert "사천오백육십칠" in result


def test_number_to_korean_multiple_numbers():
    app = MagicMock()
    result = Convert(app).number_to_korean("금액: 1,000 원, 수량: 5 개")
    assert "천" in result
    assert "오" in result


def test_wrap_by_word_creates_action():
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).wrap_by_word()
    app.api.CreateAction.assert_called_with("ParagraphShape")


def test_wrap_by_char_creates_action():
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).wrap_by_char()
    app.api.CreateAction.assert_called_with("ParagraphShape")


def test_replace_font_selects_all_and_sets_charshape():
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).replace_font("맑은 고딕", "함초롬바탕")
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "SelectAll" in calls


