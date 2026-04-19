"""Unit tests for :class:`hwpapi.collections.paragraphs.ParagraphCollection`."""
from __future__ import annotations

from unittest.mock import MagicMock

from hwpapi.collections import Collection
from hwpapi.collections.paragraphs import Paragraph, ParagraphCollection

from ._helpers import is_collection_shaped, make_app


def _app_with_paragraphs(count: int):
    impl = MagicMock()
    section = MagicMock()
    section.Paragraphs = count

    def _get_paragraph(i):
        p = MagicMock()
        p.Text = f"para-{i}"
        return p

    section.Paragraph.side_effect = _get_paragraph
    impl.XHwpDocuments.Active_XHwpDocument.Section.return_value = section
    return make_app(impl), impl


def _empty_app():
    impl = MagicMock()
    impl.XHwpDocuments.Active_XHwpDocument.Section.side_effect = Exception("no doc")
    return make_app(impl), impl


def test_protocol_compliance():
    app, _ = _empty_app()
    coll = ParagraphCollection(app)
    assert is_collection_shaped(coll)
    assert isinstance(coll, Collection)


def test_len_empty():
    app, _ = _empty_app()
    assert len(ParagraphCollection(app)) == 0


def test_iter_empty():
    app, _ = _empty_app()
    assert list(ParagraphCollection(app)) == []


def test_contains_false_when_missing():
    app, _ = _empty_app()
    assert 99 not in ParagraphCollection(app)
    assert "missing" not in ParagraphCollection(app)


def test_len_reflects_section_count():
    app, _ = _app_with_paragraphs(3)
    assert len(ParagraphCollection(app)) == 3


def test_getitem_by_ordinal():
    app, _ = _app_with_paragraphs(2)
    coll = ParagraphCollection(app)
    p = coll[0]
    assert isinstance(p, Paragraph)
    assert p.index == 0
    assert p.text == "para-0"
