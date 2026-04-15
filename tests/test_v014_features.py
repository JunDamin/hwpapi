"""Test v0.0.14 additions — Selection, Images, Presets, clean_excel_paste."""
from unittest.mock import MagicMock, call, patch

import pytest

from hwpapi.classes.selection import Selection
from hwpapi.classes.images import Image, Images
from hwpapi.presets import Presets


# ═════════════════════════════════════════════════════════════════
# Selection
# ═════════════════════════════════════════════════════════════════

@pytest.fixture
def sel_app():
    app = MagicMock()
    app.selection = "current selection text"
    app.api.Run.return_value = True
    app.api.GetTextFile.return_value = "fallback text"
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


# ═════════════════════════════════════════════════════════════════
# Images
# ═════════════════════════════════════════════════════════════════

def _mk_ctrl(ctrl_id="gso ", desc="그림"):
    c = MagicMock()
    c.CtrlID = ctrl_id
    c.UserDesc = desc
    c.Width = 20000
    c.Height = 15000
    return c


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


def test_images_iter(img_app):
    imgs = list(Images(img_app))
    assert len(imgs) == 2
    assert all(isinstance(i, Image) for i in imgs)


def test_images_len(img_app):
    assert len(Images(img_app)) == 2


def test_images_non_image_filtered():
    app = MagicMock()
    app.logger = MagicMock()
    ctrl_txt = _mk_ctrl(ctrl_id="tbl ", desc="표")  # 표는 이미지 아님
    ctrl_txt.Next = None
    app.api.HeadCtrl = ctrl_txt
    assert len(Images(app)) == 0


def test_images_getitem(img_app):
    img = Images(img_app)[0]
    assert img.index == 0


def test_images_resize_all_calls_action(img_app):
    with patch("hwpapi.classes.controls.Control") as MockCtrl:
        MockCtrl.return_value.select.return_value = True
        Images(img_app).resize_all(width="100mm")
    # ShapeObjectSize CreateAction should have been attempted
    assert img_app.api.CreateAction.called


# ═════════════════════════════════════════════════════════════════
# Presets.striped_rows
# ═════════════════════════════════════════════════════════════════

@pytest.fixture
def table_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    # TableLowerCell returns True 3x then False → 3 rows
    app.api.Run.side_effect = [True] * 50 + [False] * 50
    return app


def test_striped_rows_requires_table():
    app = MagicMock()
    app.in_table.return_value = False
    app.logger = MagicMock()
    Presets(app).striped_rows()
    app.logger.warning.assert_called()
    # No Run should be invoked after the guard
    # (beyond the in_table check)


def test_striped_rows_runs_cell_block(table_app):
    Presets(table_app).striped_rows()
    run_calls = [c.args[0] for c in table_app.api.Run.call_args_list if c.args]
    assert "TableCellBlock" in run_calls
    assert "TableCellBlockRow" in run_calls


def test_striped_rows_custom_colors(table_app):
    result = Presets(table_app).striped_rows(colors=["#AAAAAA", "#BBBBBB"])
    # returns the app (chainable)
    assert result is table_app
