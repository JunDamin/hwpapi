"""Test v0.0.16 — batch_mode, undo_group, debug, table.header_row/footer_row/current_row/align."""
from unittest.mock import MagicMock, patch

import pytest

from hwpapi.classes.debug import Debug


# ═════════════════════════════════════════════════════════════════
# Debug
# ═════════════════════════════════════════════════════════════════

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


def test_debug_state_has_cursor(dbg_app):
    s = Debug(dbg_app).state()
    assert s["cursor"] == {"doc_id": 5, "para": 10, "pos": 17}


def test_debug_state_has_page(dbg_app):
    s = Debug(dbg_app).state()
    assert s["page"] == {"current": 3, "total": 10}


def test_debug_state_has_selection(dbg_app):
    s = Debug(dbg_app).state()
    assert s["selection"]["text"] == "hello world"
    assert s["selection"]["length"] == 11


def test_debug_state_has_charshape_summary(dbg_app):
    s = Debug(dbg_app).state()
    assert s["charshape_summary"]["fontsize"] == 1100
    assert s["charshape_summary"]["bold"] is True


def test_debug_state_has_in_table(dbg_app):
    s = Debug(dbg_app).state()
    assert s["in_table"] is False


def test_debug_state_has_documents_open(dbg_app):
    s = Debug(dbg_app).state()
    assert s["documents_open"] == 2


def test_debug_state_has_version(dbg_app):
    s = Debug(dbg_app).state()
    assert s["version"] == "13.0.0"


def test_debug_state_has_filepath(dbg_app):
    s = Debug(dbg_app).state()
    assert s["filepath"] == "test.hwp"


def test_debug_state_error_handled():
    """Broken props should not crash state()."""
    app = MagicMock()
    app.logger = MagicMock()
    app.api.GetPos.side_effect = Exception("boom")
    app.current_page = property(lambda s: (_ for _ in ()).throw(Exception("boom")))
    app.selection = ""
    app.charshape = None
    app.in_table.side_effect = Exception("boom")
    # Should not raise
    s = Debug(app).state()
    assert isinstance(s, dict)


def test_debug_print_does_not_crash(dbg_app, capsys):
    Debug(dbg_app).print()
    out = capsys.readouterr().out
    assert "hwpapi debug state" in out
    assert "cursor" in out


def test_debug_timing_success(dbg_app):
    def work():
        return 42

    r = Debug(dbg_app).timing(work)
    assert r["result"] == 42
    assert r["success"] is True
    assert "elapsed_ms" in r
    assert r["exception"] is None


def test_debug_timing_failure(dbg_app):
    def boom():
        raise ValueError("x")

    r = Debug(dbg_app).timing(boom)
    assert r["success"] is False
    assert isinstance(r["exception"], ValueError)


# ═════════════════════════════════════════════════════════════════
# Debug.trace context manager
# ═════════════════════════════════════════════════════════════════

def test_debug_trace_wraps_run(capsys):
    """trace() should override api.Run temporarily."""
    from hwpapi.classes.debug import Debug

    class FakeApi:
        def Run(self, name):
            return True

    class FakeApp:
        api = FakeApi()

    app = FakeApp()
    d = Debug(app)
    with d.trace(verbose=True):
        app.api.Run("InsertText")
    out = capsys.readouterr().out
    assert "Run('InsertText')" in out
    # After exit, further Run calls should not emit trace output
    app.api.Run("NotTraced")
    after = capsys.readouterr().out
    assert "NotTraced" not in after


# ═════════════════════════════════════════════════════════════════
# TableAccessor batch methods
# ═════════════════════════════════════════════════════════════════

@pytest.fixture
def tbl_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    app.api.Run.return_value = True
    return app


def test_header_row_needs_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = False
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(app).header_row(bold=True)
    app.logger.warning.assert_called()


def test_header_row_applies_bg(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).header_row(bg="#E8F4F8")
    # CellFill CreateAction attempted
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_footer_row_applies_bg(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).footer_row(bg="#EEEEEE")
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_current_row_applies(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).current_row(bg="#FFFF00")
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_align_center(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(horz="center", scope="current_cell")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "ParagraphAlignCenter" in calls


def test_align_vertical_middle(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(vert="middle", scope="all")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "CellAlignCenter" in calls


def test_align_scope_col(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(horz="right", scope="current_col")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "TableCellBlockCol" in calls
    assert "ParagraphAlignRight" in calls


def test_batch_chain(tbl_app):
    """Methods return self for chaining."""
    from hwpapi.classes.accessors import TableAccessor
    t = TableAccessor(tbl_app)
    r = t.header_row(bg="#CCC").align(horz="center", scope="all")
    assert r is t


# ═════════════════════════════════════════════════════════════════
# batch_mode / undo_group (context managers)
# ═════════════════════════════════════════════════════════════════

def test_batch_mode_restores_mode():
    from hwpapi.core.app import App

    # Real App too heavy — mock the internals we touch
    app = MagicMock(spec=App)
    app.SILENCE_ALL_YES = 0x111111
    app.get_message_box_mode.return_value = 0
    app.api.XHwpWindows.Active_XHwpWindow.Visible = True

    # Run the actual context manager function
    from hwpapi.core.app import App as RealApp
    fn = RealApp.batch_mode
    with fn(app, hide=False):
        pass

    # Should have set + restored
    calls = [c.args for c in app.set_message_box_mode.call_args_list]
    assert (0x111111,) in calls or (app.SILENCE_ALL_YES,) in calls
    # Restore to prev (0)
    assert (0,) in calls


def test_undo_group_wraps_actions():
    from hwpapi.core.app import App as RealApp

    app = MagicMock()
    app.api.Run.return_value = True
    fn = RealApp.undo_group

    with fn(app, "test"):
        app.api.Run("InsertText")

    calls = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "SetUndoBegin" in calls
    assert "SetUndoEnd" in calls
    assert "InsertText" in calls
