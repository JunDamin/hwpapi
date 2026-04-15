"""Domain-grouped tests: images."""
from hwpapi.classes.images import Image, Images
from hwpapi.classes.selection import Selection
from hwpapi.presets import Presets
from unittest.mock import MagicMock, call, patch
import pytest


def _mk_ctrl(ctrl_id="gso ", desc="그림"):
    c = MagicMock()
    c.CtrlID = ctrl_id
    c.UserDesc = desc
    c.Width = 20000
    c.Height = 15000
    return c


@pytest.fixture
def sel_app():
    app = MagicMock()
    app.selection = "current selection text"
    app.api.Run.return_value = True
    app.api.GetTextFile.return_value = "fallback text"
    return app


@pytest.fixture
def img_app():
    app = MagicMock()
    app.logger = MagicMock()
    # Linked list: ctrl1 → ctrl2 → None
    ctrl2 = _mk_ctrl(desc="그림")
    ctrl2.Next = None
    ctrl1 = _mk_ctrl(desc="그림")
    ctrl1.Next = ctrl2
    app.api.HeadCtrl = ctrl1
    app.api.Run.return_value = True
    return app


@pytest.fixture
def table_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    # TableLowerCell returns True 3x then False → 3 rows
    app.api.Run.side_effect = [True] * 50 + [False] * 50
    return app


def test_images_iter(img_app):
    imgs = list(Images(img_app))
    assert len(imgs) == 2
    assert all(isinstance(i, Image) for i in imgs)


def test_images_len(img_app):
    assert len(Images(img_app)) == 2


def test_images_non_image_filtered():
    app = MagicMock()
    app.logger = MagicMock()
    ctrl_txt = _mk_ctrl(ctrl_id="tbl ", desc="표")  # 표는 이미지 아님
    ctrl_txt.Next = None
    app.api.HeadCtrl = ctrl_txt
    assert len(Images(app)) == 0


def test_images_getitem(img_app):
    img = Images(img_app)[0]
    assert img.index == 0


def test_images_resize_all_calls_action(img_app):
    with patch("hwpapi.classes.controls.Control") as MockCtrl:
        MockCtrl.return_value.select.return_value = True
        Images(img_app).resize_all(width="100mm")
    # ShapeObjectSize CreateAction should have been attempted
    assert img_app.api.CreateAction.called


