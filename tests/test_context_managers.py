"""Domain-grouped tests: context_managers."""
from hwpapi.classes.debug import Debug
from unittest.mock import MagicMock, patch
import pytest


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


