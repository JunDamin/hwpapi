"""Domain-grouped tests: discovery."""
from hwpapi.core.app import App
from unittest.mock import MagicMock, patch
import pytest


def test_accessor_map_has_all_categories():
    categories = set(App._ACCESSOR_MAP.keys())
    expected = {
        "Navigation & Selection",
        "Collections",
        "Structure",
        "Transform & View",
        "Quality & Templates",
        "Presets & Debug",
    }
    assert categories == expected


def test_accessor_map_entries_are_tuples():
    for category, items in App._ACCESSOR_MAP.items():
        for entry in items:
            assert isinstance(entry, tuple)
            assert len(entry) == 2
            name, desc = entry
            assert isinstance(name, str)
            assert isinstance(desc, str)
            assert len(desc) >= 5  # non-empty description


def test_accessor_map_covers_main_accessors():
    """핵심 accessor 들이 누락되지 않았는지."""
    all_names = {n for items in App._ACCESSOR_MAP.values() for n, _ in items}
    essential = {
        "move", "sel", "documents", "fields", "bookmarks", "hyperlinks",
        "images", "cell", "table", "page", "convert", "view",
        "lint", "template", "config", "preset", "debug",
    }
    missing = essential - all_names
    assert not missing, f"Missing from accessor map: {missing}"


def test_context_manager_list_non_empty():
    assert len(App._CONTEXT_MANAGERS) >= 6
    for sig, desc in App._CONTEXT_MANAGERS:
        assert "=" in sig or "(" in sig or " " in sig
        assert len(desc) > 5


def test_help_prints_categories(capsys):
    """help() prints all _ACCESSOR_MAP categories."""
    # Create a minimal mock instance without calling __init__
    app = App.__new__(App)
    App.help(app)
    out = capsys.readouterr().out
    for category in App._ACCESSOR_MAP.keys():
        assert category in out


def test_help_prints_context_managers(capsys):
    app = App.__new__(App)
    App.help(app)
    out = capsys.readouterr().out
    assert "Context managers" in out
    assert "silenced" in out
    assert "batch_mode" in out


def test_help_prints_main_properties(capsys):
    app = App.__new__(App)
    App.help(app)
    out = capsys.readouterr().out
    for p in ("text", "visible", "version", "charshape"):
        assert p in out


def test_repr_empty_uninitialized():
    app = App.__new__(App)
    r = App.__repr__(app)
    assert r.startswith("App(")
    assert ")" in r


def test_repr_resilient_to_broken_props():
    """__repr__ must not raise even if every property fails."""
    app = App.__new__(App)
    r = App.__repr__(app)
    assert isinstance(r, str)
    assert r.startswith("App(")


