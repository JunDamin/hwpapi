"""Test Fields/Bookmarks/Hyperlinks accessors (v0.0.12)."""
import warnings
from unittest.mock import MagicMock

import pytest

from hwpapi.classes.fields import (
    Fields, Field, Bookmarks, Hyperlinks, Hyperlink,
)


@pytest.fixture
def mock_app():
    app = MagicMock()
    app.api.GetFieldList.return_value = "name\x02date\x02total"
    app.field_exists = lambda n: n in ("name", "date", "total")
    app.get_field = lambda n: f"<{n}>"

    set_calls = []
    create_calls = []
    delete_calls = []
    rename_calls = []

    app.set_field = lambda n, v: set_calls.append((n, v)) or True
    app.create_field = lambda n, **kw: create_calls.append((n, kw))
    app.delete_field = lambda n: delete_calls.append(n) or True
    app.delete_all_fields = lambda: 3
    app.rename_field = lambda o, n: rename_calls.append((o, n)) or True
    app.replace_brackets_with_fields = lambda **kw: ["a", "b"]
    app.move_to_field = lambda *a, **kw: True

    # expose recorders for assertions
    app._set_calls = set_calls
    app._create_calls = create_calls
    app._delete_calls = delete_calls
    app._rename_calls = rename_calls
    return app


# ─── Backward-compatible list-like ────────────────────────────────

def test_fields_iter_yields_names(mock_app):
    f = Fields(mock_app)
    assert list(f) == ["name", "date", "total"]


def test_fields_len(mock_app):
    assert len(Fields(mock_app)) == 3


def test_fields_contains(mock_app):
    f = Fields(mock_app)
    assert "name" in f
    assert "missing" not in f
    assert 123 not in f  # non-str → False, not TypeError


def test_fields_bool(mock_app):
    assert bool(Fields(mock_app))


def test_fields_int_index_deprecated(mock_app):
    f = Fields(mock_app)
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        assert f[0] == "name"
        assert any(issubclass(x.category, DeprecationWarning) for x in w)


# ─── New dict-like + collection methods ───────────────────────────

def test_fields_getitem_by_name(mock_app):
    f = Fields(mock_app)
    fld = f["name"]
    assert isinstance(fld, Field)
    assert fld.name == "name"
    assert fld.value == "<name>"


def test_fields_getitem_missing_raises(mock_app):
    with pytest.raises(KeyError):
        Fields(mock_app)["nonexistent"]


def test_fields_setitem_existing(mock_app):
    Fields(mock_app)["name"] = "홍길동"
    assert ("name", "홍길동") in mock_app._set_calls
    assert mock_app._create_calls == []  # already exists


def test_fields_setitem_creates_if_missing(mock_app):
    Fields(mock_app)["new_field"] = "val"
    assert mock_app._create_calls == [("new_field", {})]
    assert ("new_field", "val") in mock_app._set_calls


def test_fields_delitem(mock_app):
    del Fields(mock_app)["name"]
    assert "name" in mock_app._delete_calls


def test_fields_add(mock_app):
    fld = Fields(mock_app).add("new", memo="m", direction="d")
    assert isinstance(fld, Field)
    assert fld.name == "new"
    assert mock_app._create_calls == [("new", {"memo": "m", "direction": "d"})]


def test_fields_remove(mock_app):
    assert Fields(mock_app).remove("name") is True


def test_fields_remove_all(mock_app):
    assert Fields(mock_app).remove_all() == 3


def test_fields_find_existing(mock_app):
    fld = Fields(mock_app).find("name")
    assert isinstance(fld, Field)


def test_fields_find_missing(mock_app):
    assert Fields(mock_app).find("missing") is None


def test_fields_rename(mock_app):
    assert Fields(mock_app).rename("old", "new") is True
    assert ("old", "new") in mock_app._rename_calls


def test_fields_to_dict(mock_app):
    d = Fields(mock_app).to_dict()
    assert d == {"name": "<name>", "date": "<date>", "total": "<total>"}


def test_fields_update_with_dict(mock_app):
    Fields(mock_app).update({"name": "x", "date": "y"})
    assert ("name", "x") in mock_app._set_calls
    assert ("date", "y") in mock_app._set_calls


def test_fields_update_with_kwargs(mock_app):
    Fields(mock_app).update(name="v1", date="v2")
    assert ("name", "v1") in mock_app._set_calls
    assert ("date", "v2") in mock_app._set_calls


def test_fields_from_brackets(mock_app):
    out = Fields(mock_app).from_brackets()
    assert out == ["a", "b"]


# ─── Field value object ───────────────────────────────────────────

def test_field_value_get(mock_app):
    fld = Field(mock_app, "name")
    assert fld.value == "<name>"


def test_field_value_set(mock_app):
    fld = Field(mock_app, "name")
    fld.value = "X"
    assert ("name", "X") in mock_app._set_calls


def test_field_remove(mock_app):
    Field(mock_app, "name").remove()
    assert "name" in mock_app._delete_calls


# ─── Bookmarks ────────────────────────────────────────────────────

def test_bookmarks_add():
    app = MagicMock()
    app.insert_bookmark.return_value = True
    assert Bookmarks(app).add("ch1") is True
    app.insert_bookmark.assert_called_once_with("ch1")


def test_bookmarks_remove():
    app = MagicMock()
    app.api.DeleteBookMark.return_value = True
    assert Bookmarks(app).remove("ch1") is True


def test_bookmarks_goto():
    app = MagicMock()
    app.api.SelectBookMark.return_value = True
    assert Bookmarks(app).goto("ch1") is True


def test_bookmarks_contains():
    app = MagicMock()
    app.api.ExistBookMark.return_value = True
    assert "ch1" in Bookmarks(app)
    assert 123 not in Bookmarks(app)


# ─── Hyperlinks ───────────────────────────────────────────────────

def test_hyperlinks_add():
    app = MagicMock()
    h = Hyperlinks(app).add("Anthropic", "https://anthropic.com")
    assert isinstance(h, Hyperlink)
    assert h.text == "Anthropic"
    assert h.url == "https://anthropic.com"
    app.insert_hyperlink.assert_called_once_with("Anthropic", "https://anthropic.com")
