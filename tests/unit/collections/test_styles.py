"""Unit tests for :class:`hwpapi.collections.styles.StyleCollection`.

Styles are a Phase 3 stub — enumeration requires a live HWP runtime
(HWPML export + XML parse). Tests validate Protocol compliance and
the stable lookup paths (``__getitem__`` by name, ``current``).
"""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from hwpapi.collections import Collection
from hwpapi.collections.styles import Style, StyleCollection

from ._helpers import is_collection_shaped, make_app


def _empty_app():
    impl = MagicMock()
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = StyleCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(StyleCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(StyleCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert "missing" not in StyleCollection(app)


def test_getitem_by_name_returns_style():
    app, _ = _empty_app()
    s = StyleCollection(app)["제목 1"]
    assert isinstance(s, Style)
    assert s.name == "제목 1"


def test_getitem_out_of_range_raises():
    app, _ = _empty_app()
    with pytest.raises(IndexError):
        StyleCollection(app)[0]
