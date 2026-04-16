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


def test_replace_font_targets_old_font_only():
    """v0.0.24+: replace_font 가 old_font 만 매칭해서 교체."""
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).replace_font("맑은 고딕", "함초롬바탕")
    # AllReplace 액션 + HFindReplace pset 사용
    app.api.HAction.GetDefault.assert_any_call("AllReplace", app.api.HParameterSet.HFindReplace.HSet)
    app.api.HAction.Execute.assert_any_call("AllReplace", app.api.HParameterSet.HFindReplace.HSet)


def test_replace_font_with_replace_all_legacy():
    """replace_all=True 일 때만 SelectAll → 전체 덮어쓰기."""
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).replace_font("any", "함초롬바탕", replace_all=True)
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "SelectAll" in calls
    app.set_charshape.assert_called()


def test_replace_font_empty_old_warns():
    """빈 old_font 는 경고 + 0 반환."""
    app = MagicMock()
    app.logger = MagicMock()
    result = Convert(app).replace_font("", "함초롬바탕")
    assert result == 0
    app.logger.warning.assert_called()


