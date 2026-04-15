"""Test v0.0.15 additions — title_box, subtitle_bar, table_header/footer, toc, page_numbers, compress/expand."""
from unittest.mock import MagicMock

import pytest

from hwpapi.classes.selection import Selection
from hwpapi.presets import Presets


@pytest.fixture
def table_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    app.api.Run.return_value = True
    return app


# ═════════════════════════════════════════════════════════════════
# title_box / subtitle_bar
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# table_header / table_footer
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# toc
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# page_numbers
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# Selection.compress / .expand
# ═════════════════════════════════════════════════════════════════

@pytest.fixture
def sel_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.Run.return_value = True
    app.selection = "text"
    return app


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
