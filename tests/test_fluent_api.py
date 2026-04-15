"""Test Fluent API — return self on action methods (v0.0.13)."""
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def mock_app():
    """Return an App-like mock that doesn't actually invoke HWP."""
    with patch("hwpapi.core.app.dispatch"), \
         patch("hwpapi.core.app.check_dll"), \
         patch("hwpapi.core.app.get_absolute_path", lambda p: str(p)):
        from hwpapi.core.app import App

        # Bypass __init__ — attach a minimal shape for testing.
        # `api` is a property that reads `self.engine.impl`, so give it an
        # engine with the mock attached.
        app = App.__new__(App)
        app.engine = MagicMock()
        app.engine.impl = MagicMock()
        app.actions = MagicMock()
        app.actions.InsertText = MagicMock()
        app.actions.InsertText.pset = MagicMock()
        app.actions.InsertText.run = MagicMock(return_value=True)

        # Provide stubs for the few methods our Fluent tests touch
        app.set_charshape = MagicMock(return_value=True)
        app.get_charshape = MagicMock(return_value=MagicMock())
        app._snapshot_charshape = MagicMock(return_value=None)
        app._restore_charshape = MagicMock()
        yield app


def test_insert_text_returns_self(mock_app):
    result = mock_app.insert_text("hello")
    assert result is mock_app


def test_styled_text_empty_returns_self(mock_app):
    result = mock_app.styled_text("")
    assert result is mock_app


def test_insert_page_break_returns_self(mock_app):
    result = mock_app.insert_page_break()
    assert result is mock_app


def test_insert_line_break_returns_self(mock_app):
    result = mock_app.insert_line_break()
    assert result is mock_app


def test_insert_paragraph_break_returns_self(mock_app):
    result = mock_app.insert_paragraph_break()
    assert result is mock_app


def test_insert_tab_returns_self(mock_app):
    result = mock_app.insert_tab()
    assert result is mock_app


def test_insert_bookmark_returns_self(mock_app):
    # insert_bookmark requires CreateAction pathway — patch it
    mock_app.engine.impl.CreateAction.return_value.CreateSet.return_value = MagicMock()
    result = mock_app.insert_bookmark("ch1")
    assert result is mock_app


def test_insert_hyperlink_returns_self(mock_app):
    mock_app.engine.impl.CreateAction.return_value.CreateSet.return_value = MagicMock()
    result = mock_app.insert_hyperlink("Anthropic", "https://anthropic.com")
    assert result is mock_app


def test_fluent_chain(mock_app):
    """End-to-end: chain multiple Fluent methods."""
    result = (
        mock_app.insert_text("A")
                .insert_paragraph_break()
                .insert_text("B")
                .insert_tab()
                .insert_text("C")
    )
    assert result is mock_app
    # InsertText.pset.Text should have been touched 3 times
    assert mock_app.actions.InsertText.run.call_count == 3
