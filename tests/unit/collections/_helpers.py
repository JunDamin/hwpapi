"""Shared test helpers for :mod:`hwpapi.collections` unit tests."""
from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock


def make_app(engine_impl) -> SimpleNamespace:
    """Return an App-shaped object exposing ``.engine.impl``."""
    return SimpleNamespace(engine=SimpleNamespace(impl=engine_impl))


def is_collection_shaped(obj) -> bool:
    """Structural check for the :class:`Collection` Protocol."""
    return all(
        callable(getattr(obj, attr, None)) or hasattr(obj, attr)
        for attr in (
            "__getitem__",
            "__iter__",
            "__len__",
            "__contains__",
            "names",
            "filter",
        )
    )


def make_ctrl(ctrl_id: str = "", user_desc: str = "", **kwargs) -> MagicMock:
    """Build a MagicMock shaped like an HWP CtrlCode."""
    c = MagicMock()
    c.CtrlID = ctrl_id
    c.UserDesc = user_desc
    for k, v in kwargs.items():
        setattr(c, k, v)
    c.Next = None
    return c


def chain_ctrls(*ctrls) -> MagicMock:
    """Link ctrls as a ``HeadCtrl`` → ``Next`` chain, return the head."""
    if not ctrls:
        head = MagicMock()
        head.Next = None
        return head
    for i in range(len(ctrls) - 1):
        ctrls[i].Next = ctrls[i + 1]
    ctrls[-1].Next = None
    return ctrls[0]
