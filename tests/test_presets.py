"""Domain-grouped tests: presets."""
from hwpapi.classes.convert import Convert, _int_to_korean
from hwpapi.classes.images import Image, Images
from hwpapi.classes.selection import Selection
from hwpapi.classes.view import View
from hwpapi.presets import Presets
from unittest.mock import MagicMock
from unittest.mock import MagicMock, call, patch
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


def test_striped_rows_requires_table():
    app = MagicMock()
    app.in_table.return_value = False
    app.logger = MagicMock()
    Presets(app).striped_rows()
    app.logger.warning.assert_called()


def test_striped_rows_runs_cell_block(table_app):
    Presets(table_app).striped_rows()
    run_calls = [c.args[0] for c in table_app.api.Run.call_args_list if c.args]
    assert "TableCellBlock" in run_calls
    assert "TableCellBlockRow" in run_calls


def test_striped_rows_custom_colors(table_app):
    result = Presets(table_app).striped_rows(colors=["#AAAAAA", "#BBBBBB"])
    # returns the app (chainable)
    assert result is table_app


def test_title_box_creates_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    Presets(app).title_box(text="Report", subtitle="Dept.")
    # TableCreate action should have been created
    app.api.CreateAction.assert_any_call("TableCreate")


def test_subtitle_bar_creates_table():
    app = MagicMock()
    app.logger = MagicMock()
    Presets(app).subtitle_bar("1. 개요")
    app.api.CreateAction.assert_any_call("TableCreate")


def test_table_header_requires_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = False
    Presets(app).table_header()
    app.logger.warning.assert_called()


def test_table_header_resolves_color_preset(table_app):
    Presets(table_app).table_header(color="sky")
    # CellFill CreateAction should have been attempted
    table_app.api.CreateAction.assert_any_call("CellFill")


def test_table_header_accepts_hex_color(table_app):
    Presets(table_app).table_header(color="#FF6600")
    table_app.api.CreateAction.assert_any_call("CellFill")


def test_table_header_multiple_rows(table_app):
    Presets(table_app).table_header(rows=3)
    # Run("TableLowerCell") should have been called (moving to next header row)
    calls = [c.args[0] for c in table_app.api.Run.call_args_list if c.args]
    assert "TableLowerCell" in calls


def test_table_footer_requires_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = False
    Presets(app).table_footer()
    app.logger.warning.assert_called()


def test_table_footer_runs(table_app):
    Presets(table_app).table_footer(color="gray")
    table_app.api.CreateAction.assert_any_call("CellFill")


def test_toc_calls_makeindex():
    app = MagicMock()
    app.logger = MagicMock()
    Presets(app).toc()
    app.api.CreateAction.assert_any_call("MakeIndex")


def test_toc_without_bookmarks():
    app = MagicMock()
    app.logger = MagicMock()
    result = Presets(app).toc(with_bookmarks=False, dot_leader=False, levels=2)
    assert result is app


def test_page_numbers_default():
    app = MagicMock()
    app.logger = MagicMock()
    Presets(app).page_numbers()
    app.api.CreateAction.assert_any_call("InsertAutoNum")


def test_page_numbers_with_header_filename():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    Presets(app).page_numbers(header_filename=True)
    run_calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "HeaderFooter" in run_calls


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


