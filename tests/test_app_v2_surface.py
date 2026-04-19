"""
Acceptance tests for PRD P2-002 / P2-004 — v2 slim :class:`App` surface.

These tests verify the **clean-cut break** to the v2 facade:
- ≤15 public members total
- Every v1-only method raises :class:`AttributeError`
- Every v2 ``keep_in_App`` member is present

The tests never instantiate a real :class:`App` (no COM / HWP needed).
We rely on ``App.__new__`` or class-level attribute inspection to probe
the public surface without touching the engine.
"""
from __future__ import annotations

import pytest
from unittest.mock import MagicMock

import hwpapi
from hwpapi.core.app import App


# ---------------------------------------------------------------------
# Public-surface size gate (P2-002)
# ---------------------------------------------------------------------

def _class_public_members() -> list[str]:
    """Non-underscore members defined on the App class object itself."""
    return sorted(m for m in dir(App) if not m.startswith("_"))


def test_public_surface_is_at_most_fifteen():
    """``len(dir(App) non-underscore) <= 15`` per plan acceptance."""
    members = _class_public_members()
    assert len(members) <= 15, (
        f"App surface is {len(members)} members (>15): {members}"
    )


def test_package_exports_app():
    """``import hwpapi; hwpapi.App`` still works."""
    assert hwpapi.App is App


# ---------------------------------------------------------------------
# Kept members present (P2-002 audit §1.1)
# ---------------------------------------------------------------------

CLASS_LEVEL_KEEP = {
    "__init__",
    "__enter__",
    "__exit__",
    "api",
    "close",
    "doc",
    "new",
    "open",
    "quit",
    "reload",
    "save",
    "save_as",
    "visible",
}


@pytest.mark.parametrize("name", sorted(CLASS_LEVEL_KEEP))
def test_kept_member_present_on_class(name: str):
    """Every audited keep_in_App member is defined on the class."""
    assert hasattr(App, name), f"App is missing kept member {name!r}"


@pytest.fixture
def fake_app(monkeypatch):
    """App wired to a MagicMock engine — no COM / HWP required."""
    fake_api = MagicMock()
    fake_api.XHwpWindows.Item.return_value.Visible = True
    fake_engine = MagicMock()
    fake_engine.impl = fake_api

    monkeypatch.setattr("hwpapi.core.app.Engine", lambda: fake_engine)
    monkeypatch.setattr("hwpapi.core.app.Engines", lambda: [])
    monkeypatch.setattr("hwpapi.core.app.check_dll", lambda p: None)
    monkeypatch.setattr("hwpapi.functions.get_hwp_dll_path", lambda: None)

    return App(new_app=True, is_visible=False), fake_engine, fake_api


def test_engine_is_instance_attribute(fake_app):
    """``engine`` lands in ``__init__`` — verify via patched _load path."""
    app, fake_engine, fake_api = fake_app
    assert app.engine is fake_engine
    assert app.api is fake_api


def test_actions_is_exposed(fake_app):
    """``app.actions`` is the low-level escape hatch — audit §1.5."""
    app, _, _ = fake_app
    assert hasattr(app, "actions")


# ---------------------------------------------------------------------
# Removed v1 surface → AttributeError (P2-004 + P2-006)
# ---------------------------------------------------------------------

V1_REMOVED_MEMBERS = [
    # P2-004 explicit duplication fixes
    "charshape",
    "set_charshape",
    "set_visible",
    # P2-006 examples in the plan acceptance criteria
    "insert_text",
    "get_text",
    "fields",
    "move",
    "sel",
    "view",
    "convert",
    "help",
    "status",
    # Other representative members from audit §1.6 (delete bucket)
    "version",
    "config",
    "debug",
    "lint",
    "preset",
    "template",
    "documents",
    "hwpunit_to_mm",
    "mm_to_hwpunit",
    "rgb_color",
    "get_field",
    "set_field",
    "field_names",
    "save_page_image",
    "replace_brackets_with_fields",
    "logger",  # renamed to _logger
    # audit §1.2/§1.3/§1.4 moves — AttributeError on App (still callable via .doc)
    "select_all",
    "undo",
    "redo",
    "copy",
    "cut",
    "paste",
    "find_text",
    "replace_all",
    "insert_picture",
    "setup_page",
    "page_count",
    "current_page",
    "selection",
    "text",
    "insert_table",
    "insert_bookmark",
    "insert_hyperlink",
    "new_document",  # renamed → new
    "save_block",    # renamed → save_as
]


@pytest.mark.parametrize("name", V1_REMOVED_MEMBERS)
def test_v1_member_is_attribute_error(name: str):
    """Accessing a removed v1 member on the class raises AttributeError."""
    assert not hasattr(App, name), (
        f"App still exposes v1 surface member {name!r} — should have been removed"
    )


# ---------------------------------------------------------------------
# doc / new / context manager protocol (behavioural smoke)
# ---------------------------------------------------------------------

def test_doc_returns_document_and_is_cached():
    """``app.doc`` returns a :class:`Document` and is cached."""
    from hwpapi.document import Document

    app = App.__new__(App)
    app._doc_cache = None
    app._logger = MagicMock()
    # Simulate a bound engine for .doc property access.
    app.engine = MagicMock()

    first = app.doc
    second = app.doc
    assert isinstance(first, Document)
    assert first is second  # cached


def test_context_manager_calls_quit():
    """``with App(...) as app`` calls ``quit`` on __exit__."""
    app = App.__new__(App)
    app._logger = MagicMock()
    app.engine = MagicMock()
    app.quit = MagicMock()

    with app as got:
        assert got is app

    app.quit.assert_called_once()


def test_new_is_classmethod():
    """``App.new`` is a classmethod (v2 rename of v1 ``new_document``)."""
    assert "new" in App.__dict__
    assert isinstance(App.__dict__["new"], classmethod)


def test_visible_is_property():
    """``App.visible`` is a property (not a method). ``set_visible`` removed."""
    assert isinstance(App.__dict__["visible"], property)
    assert not hasattr(App, "set_visible")
