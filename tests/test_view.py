"""Domain-grouped tests: view."""
from hwpapi.classes.convert import Convert, _int_to_korean
from hwpapi.classes.view import View
from hwpapi.presets import Presets
from unittest.mock import MagicMock
import pytest


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


