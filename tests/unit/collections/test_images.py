"""Unit tests for :class:`hwpapi.collections.images.ImageCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.collections import Collection
from hwpapi.collections.images import Image, ImageCollection

from ._helpers import chain_ctrls, is_collection_shaped, make_app, make_ctrl


def _empty_app():
    impl = MagicMock()
    impl.HeadCtrl = None
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = ImageCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(ImageCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(ImageCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert "missing" not in ImageCollection(app)


def test_iter_yields_image_objects():
    head = chain_ctrls(
        make_ctrl(ctrl_id="gso ", UserDesc="그림 A"),
        make_ctrl(ctrl_id="tbl "),
        make_ctrl(ctrl_id="gso ", UserDesc="image B"),
        make_ctrl(ctrl_id="gso ", UserDesc="표"),  # not a picture → skip
    )
    impl = MagicMock()
    impl.HeadCtrl = head
    app = make_app(impl)
    imgs = list(ImageCollection(app))
    assert len(imgs) == 2
    assert all(isinstance(img, Image) for img in imgs)
    assert imgs[0].name == "그림 A"
    assert imgs[0].index == 0
    assert imgs[1].index == 1
