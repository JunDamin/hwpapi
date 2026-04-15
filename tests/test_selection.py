"""Domain-grouped tests: selection."""
from hwpapi.classes.convert import Convert, _int_to_korean
from hwpapi.classes.debug import Debug
from hwpapi.classes.images import Image, Images
from hwpapi.classes.selection import Selection
from hwpapi.classes.view import View
from hwpapi.presets import Presets
from unittest.mock import MagicMock
from unittest.mock import MagicMock, call, patch
from unittest.mock import MagicMock, patch
import pytest


def _mk_ctrl(ctrl_id="gso ", desc="그림"):
    c = MagicMock()
    c.CtrlID = ctrl_id
    c.UserDesc = desc
    c.Width = 20000
    c.Height = 15000
    return c


@pytest.fixture
def sel_app():
    app = MagicMock()
    app.selection = "current selection text"
    app.api.Run.return_value = True
    app.api.GetTextFile.return_value = "fallback text"
    return app


@pytest.fixture
def img_app():
    app = MagicMock()
    app.logger = MagicMock()
    # Linked list: ctrl1 → ctrl2 → None
    ctrl2 = _mk_ctrl(desc="그림")
    ctrl2.Next = None
    ctrl1 = _mk_ctrl(desc="그림")
    ctrl1.Next = ctrl2
    app.api.HeadCtrl = ctrl1
    app.api.Run.return_value = True
    return app


@pytest.fixture
def table_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    # TableLowerCell returns True 3x then False → 3 rows
    app.api.Run.side_effect = [True] * 50 + [False] * 50
    return app


@pytest.fixture
def dbg_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.GetPos.return_value = (5, 10, 17)
    app.current_page = 3
    app.page_count = 10
    app.selection = "hello world"
    cs = MagicMock()
    cs.fontsize = 1100
    cs.bold = True
    cs.italic = False
    cs.text_color = "#000000"
    cs.shade_color = None
    app.charshape = cs
    app.in_table.return_value = False
    app.documents = [MagicMock(), MagicMock()]
    app.visible = True
    app.version = "13.0.0"
    app.get_filepath.return_value = "test.hwp"
    return app


@pytest.fixture
def tbl_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    app.api.Run.return_value = True
    return app


def test_selection_text_uses_app_selection(sel_app):
    s = Selection(sel_app)
    assert s.text == "current selection text"


def test_selection_text_empty_means_is_empty(sel_app):
    sel_app.selection = ""
    assert Selection(sel_app).is_empty


def test_selection_clear(sel_app):
    Selection(sel_app).clear()
    sel_app.api.Run.assert_any_call("Cancel")


def test_selection_current_word_tries_selectword(sel_app):
    Selection(sel_app).current_word()
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list]
    assert "SelectWord" in calls or "MoveWordBegin" in calls


def test_selection_current_line(sel_app):
    Selection(sel_app).current_line()
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list]
    assert "MoveLineBegin" in calls
    assert "MoveSelLineEnd" in calls


def test_selection_current_paragraph(sel_app):
    Selection(sel_app).current_paragraph()
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list]
    assert "MoveParaBegin" in calls
    assert "MoveSelParaEnd" in calls


def test_selection_to_paragraph_end(sel_app):
    Selection(sel_app).to_paragraph_end()
    sel_app.api.Run.assert_any_call("MoveSelParaEnd")


def test_selection_to_document_end(sel_app):
    Selection(sel_app).to_document_end()
    sel_app.api.Run.assert_any_call("MoveSelDocEnd")


def test_selection_expand_char_positive(sel_app):
    Selection(sel_app).expand_char(3)
    run_calls = [c.args[0] for c in sel_app.api.Run.call_args_list]
    assert run_calls.count("MoveSelRight") == 3


def test_selection_expand_char_negative(sel_app):
    Selection(sel_app).expand_char(-2)
    run_calls = [c.args[0] for c in sel_app.api.Run.call_args_list]
    assert run_calls.count("MoveSelLeft") == 2


def test_selection_chainable(sel_app):
    s = Selection(sel_app)
    assert s.clear() is s
    assert s.current_paragraph() is s
    assert s.to_line_end() is s


def test_compress_calls_decrease_actions(sel_app):
    Selection(sel_app).compress()
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list if c.args]
    assert "CharShapeSpacingDecrease" in calls
    assert "CharShapeScaleDecrease" in calls


def test_compress_step_3(sel_app):
    Selection(sel_app).compress(step=3)
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list if c.args]
    assert calls.count("CharShapeSpacingDecrease") == 3
    assert calls.count("CharShapeScaleDecrease") == 3


def test_expand_calls_increase_actions(sel_app):
    Selection(sel_app).expand()
    calls = [c.args[0] for c in sel_app.api.Run.call_args_list if c.args]
    assert "CharShapeSpacingIncrease" in calls
    assert "CharShapeScaleIncrease" in calls


def test_compress_expand_chainable(sel_app):
    s = Selection(sel_app)
    assert s.compress(1).expand(1) is s


def test_debug_state_has_selection(dbg_app):
    s = Debug(dbg_app).state()
    assert s["selection"]["text"] == "hello world"
    assert s["selection"]["length"] == 11


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


