"""Unit tests for :class:`hwpapi.collections.paragraphs.Paragraph` element API."""
from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from hwpapi.collections.paragraphs import Paragraph, Run


def _app_with_action_psets(**pset_by_action):
    """Return an App-shaped mock exposing ``actions.<Name>.pset``."""
    actions = SimpleNamespace()
    for name, pset in pset_by_action.items():
        setattr(actions, name, SimpleNamespace(pset=pset))
    impl = MagicMock()
    return SimpleNamespace(engine=SimpleNamespace(impl=impl), actions=actions)


def test_text_empty_when_raw_none():
    app = _app_with_action_psets()
    p = Paragraph(app, 0, None)
    assert p.text == ""


def test_text_reads_raw_text_attr():
    app = _app_with_action_psets()
    raw = MagicMock()
    raw.Text = "안녕"
    p = Paragraph(app, 0, raw)
    assert p.text == "안녕"


def test_style_reads_style_name_attribute():
    app = _app_with_action_psets()
    raw = MagicMock()
    raw.StyleName = "바탕글"
    p = Paragraph(app, 0, raw)
    assert p.style == "바탕글"


def test_style_empty_when_raw_none():
    app = _app_with_action_psets()
    p = Paragraph(app, 0, None)
    assert p.style == ""


def test_style_falls_back_to_style_attribute():
    app = _app_with_action_psets()
    raw = MagicMock(spec=["Style"])
    raw.Style = "본문"
    p = Paragraph(app, 0, raw)
    assert p.style == "본문"


def test_charshape_delegates_to_action_pset():
    sentinel = MagicMock(name="charshape_pset")
    app = _app_with_action_psets(CharShape=sentinel)
    raw = MagicMock()
    raw.Text = ""
    p = Paragraph(app, 0, raw)
    assert p.charshape is sentinel


def test_parashape_delegates_to_paragraph_shape_action():
    sentinel = MagicMock(name="parashape_pset")
    app = _app_with_action_psets(ParagraphShape=sentinel)
    raw = MagicMock()
    raw.Text = ""
    p = Paragraph(app, 0, raw)
    assert p.parashape is sentinel


def test_charshape_returns_none_when_actions_missing():
    app = SimpleNamespace(engine=SimpleNamespace(impl=MagicMock()))
    p = Paragraph(app, 0, None)
    assert p.charshape is None
    assert p.parashape is None


def test_runs_returns_single_run_covering_paragraph():
    app = _app_with_action_psets()
    raw = MagicMock()
    raw.Text = "hello world"
    p = Paragraph(app, 0, raw)
    runs = p.runs
    assert isinstance(runs, list)
    assert len(runs) == 1
    run = runs[0]
    assert isinstance(run, Run)
    assert run.text == "hello world"


def test_runs_preserves_paragraph_reference():
    app = _app_with_action_psets()
    raw = MagicMock()
    raw.Text = "foo"
    p = Paragraph(app, 7, raw)
    (run,) = p.runs
    assert run.paragraph is p


def test_repr_shows_index():
    app = _app_with_action_psets()
    p = Paragraph(app, 3, None)
    assert repr(p) == "Paragraph(#3)"
