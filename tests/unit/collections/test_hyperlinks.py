"""Unit tests for :class:`hwpapi.collections.hyperlinks.HyperlinkCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.collections import Collection
from hwpapi.collections.hyperlinks import Hyperlink, HyperlinkCollection

from ._helpers import chain_ctrls, is_collection_shaped, make_app, make_ctrl


def _empty_app():
    impl = MagicMock()
    impl.HeadCtrl = None
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = HyperlinkCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(HyperlinkCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(HyperlinkCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert "missing" not in HyperlinkCollection(app)


def test_iter_yields_hyperlink_objects():
    def props_for(text, url):
        props = MagicMock()
        props.Item.side_effect = lambda k: {
            "Text": text, "Command": url, "URL": url, "Location": url
        }.get(k)
        return props

    head = chain_ctrls(
        make_ctrl(ctrl_id="%hlk", UserDesc="anthropic",
                  Properties=props_for("Anthropic", "https://anthropic.com")),
        make_ctrl(ctrl_id="tbl "),
        make_ctrl(ctrl_id="%hlk", UserDesc="python",
                  Properties=props_for("Python", "https://python.org")),
    )
    impl = MagicMock()
    impl.HeadCtrl = head
    app = make_app(impl)
    links = list(HyperlinkCollection(app))
    assert len(links) == 2
    assert all(isinstance(h, Hyperlink) for h in links)
    assert links[0].text == "Anthropic"


def test_hyperlink_equality():
    assert Hyperlink("a", "u") == Hyperlink("a", "u")
    assert Hyperlink("a", "u") != Hyperlink("b", "u")
