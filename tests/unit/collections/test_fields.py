"""Unit tests for :class:`hwpapi.collections.fields.FieldCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from hwpapi.collections import Collection
from hwpapi.collections.fields import Field, FieldCollection

from ._helpers import is_collection_shaped, make_app


def _app(field_list: str = ""):
    impl = MagicMock()
    impl.GetFieldList.return_value = field_list
    impl.GetFieldText.return_value = ""
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _app()
    coll = FieldCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _app("")
    assert len(FieldCollection(app)) == 0


def test_iter_empty():
    app, _ = _app("")
    assert list(FieldCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _app("")
    assert "missing" not in FieldCollection(app)
    assert 123 not in FieldCollection(app)


def test_iter_yields_field_names_stx_separated():
    app, _ = _app("customer\x02date\x02amount")
    fields = list(FieldCollection(app))
    assert [f.name for f in fields] == ["customer", "date", "amount"]
    assert all(isinstance(f, Field) for f in fields)


def test_getitem_returns_field_object():
    app, _ = _app("name\x02date")
    coll = FieldCollection(app)
    f = coll["name"]
    assert isinstance(f, Field)
    assert f.name == "name"
    assert coll[0].name == "name"
    assert coll[1].name == "date"


def test_getitem_missing_raises_keyerror():
    app, _ = _app("name")
    with pytest.raises(KeyError):
        FieldCollection(app)["missing"]


def test_setitem_calls_put_field_text():
    app, impl = _app("name")
    FieldCollection(app)["name"] = "홍길동"
    impl.PutFieldText.assert_called_once_with("name", "홍길동")


def test_names_deduplicates_in_order():
    app, _ = _app("a\x02b\x02a\x02c")
    assert FieldCollection(app).names() == ["a", "b", "c"]
