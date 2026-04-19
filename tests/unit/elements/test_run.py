"""Unit tests for :class:`hwpapi.collections.paragraphs.Run`."""
from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from hwpapi.collections.paragraphs import Paragraph, Run


def _app(**pset_by_action):
    actions = SimpleNamespace()
    for name, pset in pset_by_action.items():
        setattr(actions, name, SimpleNamespace(pset=pset))
    return SimpleNamespace(engine=SimpleNamespace(impl=MagicMock()), actions=actions)


def _paragraph(text: str, index: int = 0, app=None) -> Paragraph:
    app = app or _app()
    raw = MagicMock()
    raw.Text = text
    return Paragraph(app, index, raw)


def test_default_run_covers_paragraph():
    p = _paragraph("hello world")
    run = Run(p)
    assert run.start == 0
    assert run.end == -1
    assert run.text == "hello world"


def test_run_slice_semantics():
    p = _paragraph("abcdefghij")
    run = Run(p, start=3, end=7)
    assert run.text == "defg"
    assert len(run) == 4


def test_run_end_sentinel_is_paragraph_length():
    p = _paragraph("xyz")
    run = Run(p, start=1, end=-1)
    assert run.text == "yz"
    assert len(run) == 2


def test_run_len_on_empty_paragraph_is_zero():
    p = _paragraph("")
    run = Run(p)
    assert run.text == ""
    assert len(run) == 0


def test_run_charshape_is_paragraph_charshape():
    sentinel = MagicMock(name="charshape_pset")
    p = _paragraph("foo", app=_app(CharShape=sentinel))
    run = Run(p)
    assert run.charshape is sentinel


def test_run_does_not_hit_com_on_construction():
    # Build Paragraph with a raw that would blow up on access
    raw = MagicMock()
    raw.Text = MagicMock(side_effect=AssertionError("Text should not be read"))
    app = _app()
    p = Paragraph(app, 0, raw)
    # Construction must succeed; accessing properties defers.
    run = Run(p, 0, 3)
    # Access .start/.end without hitting .text — should be safe.
    assert run.start == 0
    assert run.end == 3


def test_run_paragraph_back_reference():
    p = _paragraph("foo", index=4)
    run = Run(p)
    assert run.paragraph is p


def test_run_repr_includes_slice_info():
    p = _paragraph("foo", index=2)
    run = Run(p, 1, 3)
    r = repr(run)
    assert "para=#2" in r
    assert "1:3" in r


def test_run_repr_end_sentinel():
    p = _paragraph("foo")
    run = Run(p)
    r = repr(run)
    assert "0:end" in r
