"""
Tests for v3 Documents collection + Document v3 surface (Mock-based).

Live-HWP smoke 는 별도 — 여기는 Mock 으로 빠르게 구조 검증만.
"""
from __future__ import annotations

from unittest.mock import MagicMock, PropertyMock

import pytest

from hwpapi.collections.documents import DocumentCollection
from hwpapi.document import Document, _DocActions, _DocCursor


# ── helpers ──────────────────────────────────────────────────────

def _mock_app(doc_count: int = 1, doc_names: list[str] | None = None):
    """Build an App-like mock with XHwpDocuments populated."""
    if doc_names is None:
        doc_names = [f"doc{i}.hwp" for i in range(doc_count)]
    raws = []
    for name in doc_names:
        raw = MagicMock(name=f"IXHwpDoc({name})")
        raw.FullName = f"C:/{name}" if name != "(untitled)" else ""
        raw.Modified = False
        raws.append(raw)

    xhd = MagicMock(name="XHwpDocuments")
    xhd.Count = len(raws)
    xhd.Item = lambda i: raws[i]
    xhd.Active_XHwpDocument = raws[0] if raws else None

    api = MagicMock(name="api")
    api.XHwpDocuments = xhd
    api.GetTextFile = lambda *a, **k: "mock-text"

    app = MagicMock(name="app")
    app.api = api
    app.actions = MagicMock(name="actions")
    return app, raws


# ── DocumentCollection ───────────────────────────────────────────

def test_collection_len_and_iter():
    app, _ = _mock_app(doc_count=3)
    docs = DocumentCollection(app)
    assert len(docs) == 3
    items = list(docs)
    assert all(isinstance(d, Document) for d in items)
    assert len(items) == 3


def test_collection_open_returns_document(monkeypatch):
    app, raws = _mock_app(doc_count=1)
    monkeypatch.setattr(
        "hwpapi.functions.get_absolute_path", lambda p: f"C:/{p}",
    )
    docs = DocumentCollection(app)
    doc = docs.open("foo.hwp")
    assert isinstance(doc, Document)
    app.api.Open.assert_called_once()


def test_collection_add_returns_document():
    app, raws = _mock_app(doc_count=1)
    docs = DocumentCollection(app)
    doc = docs.add()
    assert isinstance(doc, Document)
    app.api.Run.assert_called_with("FileNew")


def test_collection_active_returns_document():
    app, raws = _mock_app(doc_count=2, doc_names=["a.hwp", "b.hwp"])
    docs = DocumentCollection(app)
    active = docs.active
    assert isinstance(active, Document)
    # active doc 의 raw 는 Active_XHwpDocument 와 동일
    assert active.raw is raws[0]


def test_collection_indexing_int():
    app, raws = _mock_app(doc_count=3, doc_names=["a", "b", "c"])
    docs = DocumentCollection(app)
    d = docs[1]
    assert isinstance(d, Document)
    assert d.raw is raws[1]


def test_collection_indexing_negative():
    app, raws = _mock_app(doc_count=3)
    docs = DocumentCollection(app)
    assert docs[-1].raw is raws[2]


def test_collection_indexing_out_of_range():
    app, _ = _mock_app(doc_count=2)
    docs = DocumentCollection(app)
    with pytest.raises(IndexError):
        _ = docs[5]


def test_collection_indexing_by_name():
    app, raws = _mock_app(doc_count=2, doc_names=["alpha.hwp", "beta.hwp"])
    docs = DocumentCollection(app)
    d = docs["beta.hwp"]
    assert d.raw is raws[1]


def test_collection_indexing_by_unknown_name():
    app, _ = _mock_app(doc_count=2)
    docs = DocumentCollection(app)
    with pytest.raises(KeyError):
        _ = docs["nope.hwp"]


def test_collection_contains():
    app, _ = _mock_app(doc_count=2, doc_names=["a.hwp", "b.hwp"])
    docs = DocumentCollection(app)
    assert "a.hwp" in docs
    assert "missing.hwp" not in docs
    assert 0 in docs
    assert 99 not in docs


# ── Document meta ────────────────────────────────────────────────

def test_document_path_and_name():
    app, raws = _mock_app(doc_names=["report.hwp"])
    raws[0].FullName = "C:/work/report.hwp"
    doc = Document(app, _raw=raws[0])
    assert doc.path == "C:/work/report.hwp"
    assert doc.name == "report.hwp"


def test_document_path_none_for_unsaved():
    app, raws = _mock_app(doc_names=["(untitled)"])
    raws[0].FullName = ""
    doc = Document(app, _raw=raws[0])
    assert doc.path is None
    assert doc.name == "(untitled)"


def test_document_saved_flag():
    app, raws = _mock_app()
    raws[0].Modified = True
    doc = Document(app, _raw=raws[0])
    assert doc.saved is False
    raws[0].Modified = False
    assert doc.saved is True


# ── Document lifecycle ───────────────────────────────────────────

def test_document_activate_calls_setactive():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    doc.activate()
    raws[0].SetActive_XHwpDocument.assert_called_once()


def test_document_close_invalidates_handle():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    assert doc.close(save=False) is True
    raws[0].Close.assert_called_with(False)
    assert doc.raw is None
    # 두 번 close 면 False
    assert doc.close() is False


def test_document_save_inplace():
    app, raws = _mock_app()
    raws[0].FullName = "C:/foo.hwp"
    doc = Document(app, _raw=raws[0])
    result = doc.save()
    app.api.Save.assert_called_once()
    assert result == "C:/foo.hwp"


def test_document_save_as_with_path(monkeypatch):
    monkeypatch.setattr("hwpapi.functions.get_absolute_path", lambda p: f"C:/{p}")
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    doc.save_as("out.pdf")
    app.api.SaveAs.assert_called_once()
    args, _ = app.api.SaveAs.call_args
    assert args[0] == "C:/out.pdf"
    assert args[1] == "PDF"


# ── Document text I/O ────────────────────────────────────────────

def test_document_insert_text_with_newlines():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    doc.insert_text("line1\nline2\n")
    # InsertText 가 두 번 호출되고 BreakPara 가 두 번
    assert app.actions.InsertText.run.call_count == 2
    runs = [c.args for c in app.api.Run.call_args_list]
    break_count = sum(1 for r in runs if r and r[0] == "BreakPara")
    assert break_count == 2


def test_document_text_property_reads_via_api():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    assert doc.text == "mock-text"


# ── _DocActions proxy ────────────────────────────────────────────

def test_doc_actions_attr_triggers_activate():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    proxy = doc.actions
    assert isinstance(proxy, _DocActions)
    _ = proxy.SomeAction
    raws[0].SetActive_XHwpDocument.assert_called_once()


def test_doc_actions_returns_underlying_action():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    sentinel = object()
    app.actions.MyAction = sentinel
    assert doc.actions.MyAction is sentinel


# ── _DocCursor ───────────────────────────────────────────────────

def test_doc_cursor_in_table_uses_keyindicator():
    app, raws = _mock_app()
    app.api.KeyIndicator = lambda: (0, 0, 0, 0, 0, 0, True, 0, 0)
    doc = Document(app, _raw=raws[0])
    assert doc.cursor.in_table() is True


def test_doc_cursor_in_table_false():
    app, raws = _mock_app()
    app.api.KeyIndicator = lambda: (0, 0, 0, 0, 0, 0, False, 0, 0)
    doc = Document(app, _raw=raws[0])
    assert doc.cursor.in_table() is False


def test_doc_cursor_goto_page_calls_movepos():
    app, raws = _mock_app()
    doc = Document(app, _raw=raws[0])
    doc.cursor.goto_page(3)
    app.api.MovePos.assert_called_with(23, 2, 0)
