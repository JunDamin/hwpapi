"""Test v0.0.17 — Convert, View, page_border, highlight_yellow, summary_box."""
from unittest.mock import MagicMock

import pytest

from hwpapi.classes.convert import Convert, _int_to_korean
from hwpapi.classes.view import View
from hwpapi.presets import Presets


# ═════════════════════════════════════════════════════════════════
# _int_to_korean
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# Convert.number_to_korean
# ═════════════════════════════════════════════════════════════════

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


def test_number_to_korean_from_selection_replaces_text():
    app = MagicMock()
    app.selection = "1,234"
    app.logger = MagicMock()
    out = Convert(app).number_to_korean()
    assert out is not None
    # Should call insert_text to replace selection
    app.insert_text.assert_called()


def test_number_to_korean_empty_selection_returns_none():
    app = MagicMock()
    app.selection = ""
    app.logger = MagicMock()
    assert Convert(app).number_to_korean() is None


# ═════════════════════════════════════════════════════════════════
# Convert.wrap_by_word / wrap_by_char
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# Convert.replace_font
# ═════════════════════════════════════════════════════════════════

def test_replace_font_selects_all_and_sets_charshape():
    app = MagicMock()
    app.logger = MagicMock()
    Convert(app).replace_font("맑은 고딕", "함초롬바탕")
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "SelectAll" in calls


# ═════════════════════════════════════════════════════════════════
# View.zoom
# ═════════════════════════════════════════════════════════════════

def test_view_zoom_default():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).zoom(150)
    # PictureScale CreateAction attempted
    app.api.CreateAction.assert_called_with("PictureScale")


def test_view_zoom_clamps_low():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).zoom(5)   # below 10 → clamped to 10
    pset = app.api.CreateAction.return_value.CreateSet.return_value
    # Item calls should include Zoom=10
    calls = [c.args for c in pset.SetItem.call_args_list if c.args]
    assert ("Zoom", 10) in calls


def test_view_zoom_clamps_high():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).zoom(1000)
    pset = app.api.CreateAction.return_value.CreateSet.return_value
    calls = [c.args for c in pset.SetItem.call_args_list if c.args]
    assert ("Zoom", 500) in calls


def test_view_zoom_actual():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).zoom_actual()
    pset = app.api.CreateAction.return_value.CreateSet.return_value
    calls = [c.args for c in pset.SetItem.call_args_list if c.args]
    assert ("Zoom", 100) in calls


def test_view_zoom_fit_page():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    View(app).zoom_fit_page()
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "ViewZoomFitPage" in calls or "MoveViewFitPage" in calls


def test_view_full_screen():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).full_screen()
    app.api.Run.assert_any_call("FullScreen")


def test_view_page_mode():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).page_mode()
    app.api.Run.assert_any_call("ViewPage")


def test_view_draft_mode():
    app = MagicMock()
    app.logger = MagicMock()
    View(app).draft_mode()
    app.api.Run.assert_any_call("ViewDraft")


def test_view_scroll_to_cursor():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    View(app).scroll_to_cursor()
    # Either MoveScrollCurPos or fallback
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "MoveScrollCurPos" in calls or "MoveRight" in calls


# ═════════════════════════════════════════════════════════════════
# Preset — page_border / highlight_yellow / summary_box
# ═════════════════════════════════════════════════════════════════

def test_page_border_enable():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    Presets(app).page_border(enable=True)
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "MasterPage" in calls


def test_page_border_disable():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    Presets(app).page_border(enable=False)
    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "MasterPage" in calls
    assert "SelectAll" in calls


def test_highlight_yellow_applies_shade():
    app = MagicMock()
    app.logger = MagicMock()
    cs = MagicMock()
    cs.shade_color = None
    app.charshape = cs
    Presets(app).highlight_yellow()
    # set_charshape should have been called with shade_color
    app.set_charshape.assert_called()
    call_kwargs = app.set_charshape.call_args.kwargs
    assert call_kwargs.get("shade_color") == "#FFFF00"


def test_highlight_yellow_toggles_off_when_yellow():
    app = MagicMock()
    app.logger = MagicMock()
    cs = MagicMock()
    cs.shade_color = "#FFFF00"
    app.charshape = cs
    Presets(app).highlight_yellow(toggle=True)
    # Should have called set_charshape(shade_color=None)
    call_kwargs = app.set_charshape.call_args.kwargs
    assert call_kwargs.get("shade_color") is None


def test_summary_box_creates_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    Presets(app).summary_box("test content")
    app.api.CreateAction.assert_any_call("TableCreate")
