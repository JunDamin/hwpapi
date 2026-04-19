"""Unit tests for :class:`hwpapi.collections.tables.TableCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.collections import Collection
from hwpapi.collections.tables import Table, TableCollection

from ._helpers import chain_ctrls, is_collection_shaped, make_app, make_ctrl


def _empty_app():
    impl = MagicMock()
    impl.HeadCtrl = None
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = TableCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(TableCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(TableCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert "missing" not in TableCollection(app)
    assert 99 not in TableCollection(app)


def test_iter_yields_tables_by_ordinal():
    head = chain_ctrls(
        make_ctrl(ctrl_id="tbl ", UserDesc="표 1: 매출"),
        make_ctrl(ctrl_id="gso "),
        make_ctrl(ctrl_id="tbl ", UserDesc=""),
    )
    impl = MagicMock()
    impl.HeadCtrl = head
    app = make_app(impl)
    tables = list(TableCollection(app))
    assert len(tables) == 2
    assert all(isinstance(t, Table) for t in tables)
    assert tables[0].index == 0
    assert tables[0].caption == "표 1: 매출"
    assert tables[1].caption == ""


def test_getitem_by_int_returns_table():
    head = chain_ctrls(make_ctrl(ctrl_id="tbl ", UserDesc="A"))
    impl = MagicMock()
    impl.HeadCtrl = head
    coll = TableCollection(make_app(impl))
    assert coll[0].caption == "A"


def test_getitem_by_missing_caption_returns_none():
    head = chain_ctrls(make_ctrl(ctrl_id="tbl ", UserDesc="A"))
    impl = MagicMock()
    impl.HeadCtrl = head
    assert TableCollection(make_app(impl))["missing"] is None
