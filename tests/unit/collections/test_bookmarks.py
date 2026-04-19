"""Unit tests for :class:`hwpapi.collections.bookmarks.BookmarkCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.collections import Collection
from hwpapi.collections.bookmarks import Bookmark, BookmarkCollection

from ._helpers import chain_ctrls, is_collection_shaped, make_app, make_ctrl


def _empty_app():
    impl = MagicMock()
    impl.HeadCtrl = None
    impl.ExistBookMark.return_value = False
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = BookmarkCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(BookmarkCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(BookmarkCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert "missing" not in BookmarkCollection(app)
    assert 123 not in BookmarkCollection(app)


def test_iter_yields_bookmark_names():
    head = chain_ctrls(
        make_ctrl(ctrl_id="bokm", CtrlCh="ch1"),
        make_ctrl(ctrl_id="tbl "),  # non-bookmark, should be skipped
        make_ctrl(ctrl_id="bokm", CtrlCh="ch2"),
    )
    impl = MagicMock()
    impl.HeadCtrl = head
    app = make_app(impl)
    names = [b.name for b in BookmarkCollection(app)]
    assert names == ["ch1", "ch2"]


def test_contains_delegates_to_exist_bookmark():
    impl = MagicMock()
    impl.HeadCtrl = None
    impl.ExistBookMark.return_value = True
    app = make_app(impl)
    assert "anything" in BookmarkCollection(app)
    impl.ExistBookMark.assert_called_with("anything")


def test_bookmark_eq_with_string():
    assert Bookmark(make_app(MagicMock()), "ch1") == "ch1"
