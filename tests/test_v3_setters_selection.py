"""v3.1 — set_charshape / set_parashape / Selection / Range / find_all (Mock)."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.document import Document
from hwpapi.selection import Selection, Range


def _mock_app():
    raw = MagicMock(name="IXHwpDoc")
    raw.FullName = "C:/x.hwp"
    raw.Modified = False
    xhd = MagicMock(name="XHwpDocuments")
    xhd.Count = 1
    xhd.Item = lambda i: raw
    xhd.Active_XHwpDocument = raw
    api = MagicMock(name="api")
    api.XHwpDocuments = xhd
    api.GetSelectedText = lambda: "선택됨"
    api.GetTextFile = lambda *a, **k: "전체텍스트"
    app = MagicMock(name="app")
    app.api = api
    app.actions = MagicMock(name="actions")
    return app, raw


# ── set_charshape / set_parashape ────────────────────────────────

def test_set_charshape_calls_apply(monkeypatch):
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)

    captured = {}
    def fake_apply(app_, action_name, slot, values):
        captured["action"] = action_name
        captured["values"] = dict(values)

    monkeypatch.setattr("hwpapi.context.scopes._apply", fake_apply)
    doc.set_charshape(bold=True, height=1400)

    assert captured["action"] == "CharShape"
    raw.SetActive_XHwpDocument.assert_called()


def test_set_parashape_normalises_align(monkeypatch):
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)

    captured = {}
    monkeypatch.setattr(
        "hwpapi.context.scopes._apply",
        lambda a, n, s, v: captured.update({"action": n, "values": dict(v)}),
    )
    doc.set_parashape(align="center")
    assert captured["action"] == "ParaShape"
    # Align type translated + normalised
    assert "AlignType" in captured["values"]


# ── Selection ────────────────────────────────────────────────────

def test_selection_text():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    sel = doc.selection
    assert isinstance(sel, Selection)
    assert sel.text == "선택됨"


def test_selection_exists_true():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    assert doc.selection.exists is True


def test_selection_exists_false():
    app, raw = _mock_app()
    app.api.GetSelectedText = lambda: ""
    doc = Document(app, _raw=raw)
    assert doc.selection.exists is False


def test_selection_delete():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    doc.selection.delete()
    runs = [c.args[0] for c in app.api.Run.call_args_list if c.args]
    assert "Delete" in runs


def test_selection_set_charshape_delegates(monkeypatch):
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    captured = {}
    monkeypatch.setattr(
        "hwpapi.context.scopes._apply",
        lambda a, n, s, v: captured.update({"action": n}),
    )
    doc.selection.set_charshape(italic=True)
    assert captured["action"] == "CharShape"


# ── Range ────────────────────────────────────────────────────────

def test_range_construction():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    r = doc.range(2, 5)
    assert isinstance(r, Range)
    assert r.start == 2
    assert r.end == 5


def test_range_default_end():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    r = doc.range(3)
    assert r.start == 3
    assert r.end == 3


def test_range_text_calls_select_then_read():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    r = doc.range(0, 1)
    text = r.text
    assert text == "선택됨"
    # Ensure MovePos was invoked (i.e., selection set up)
    app.api.MovePos.assert_called()


def test_range_set_charshape(monkeypatch):
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    captured = {}
    monkeypatch.setattr(
        "hwpapi.context.scopes._apply",
        lambda a, n, s, v: captured.update({"action": n}),
    )
    doc.range(0, 2).set_charshape(bold=True)
    assert captured["action"] == "CharShape"


# ── find_all / replace_brackets ──────────────────────────────────

def test_replace_brackets_calls_replace_all_per_key():
    app, raw = _mock_app()
    doc = Document(app, _raw=raw)
    # Stub replace_all to count
    calls = []
    doc.replace_all = lambda find, replace: calls.append((find, replace)) or 1
    n = doc.replace_brackets({
        "{name}": "홍길동",
        "{date}": "2026-04-29",
    })
    assert n == 2
    assert ("{name}", "홍길동") in calls
    assert ("{date}", "2026-04-29") in calls
