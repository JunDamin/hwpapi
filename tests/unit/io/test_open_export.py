"""Unit tests for :mod:`hwpapi.io.open` and :mod:`hwpapi.io.export` (v3)."""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from hwpapi.errors import FileIOError
from hwpapi.io.open import open_file, new_document
from hwpapi.io.export import export_pdf, export_image, export_text


# ---------------------------------------------------------------------
# open_file (v3: app.docs.open(path, format=...) returns a Document)
# ---------------------------------------------------------------------

def test_open_file_delegates_to_docs_open():
    """``open_file(app, path)`` calls ``app.docs.open(path, format=None)``."""
    app = MagicMock()
    fake_doc = MagicMock()
    fake_doc.path = "C:/abs/x.hwp"
    app.docs.open.return_value = fake_doc

    result = open_file(app, "x.hwp")

    app.docs.open.assert_called_once_with("x.hwp", format=None)
    assert result == "C:/abs/x.hwp"


def test_open_file_forwards_format_argument():
    """Explicit ``format=`` is passed through to ``app.docs.open``."""
    app = MagicMock()
    fake_doc = MagicMock()
    fake_doc.path = "C:/abs/x.pdf"
    app.docs.open.return_value = fake_doc

    open_file(app, "x.pdf", format="PDF")
    app.docs.open.assert_called_once_with("x.pdf", format="PDF")


class _FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` on non-Windows runners."""


def test_open_file_wraps_com_error(monkeypatch):
    """A COM error from ``app.docs.open`` becomes :class:`FileIOError`."""
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    app = MagicMock()
    app.docs.open.side_effect = _FakeComError("bad path")

    with pytest.raises(FileIOError) as excinfo:
        open_file(app, "missing.hwp")

    assert "missing.hwp" in str(excinfo.value)
    assert isinstance(excinfo.value.__cause__, _FakeComError)


# ---------------------------------------------------------------------
# new_document — uses App.new classmethod (unchanged)
# ---------------------------------------------------------------------

def test_new_document_calls_app_new(monkeypatch):
    sentinel = object()
    from hwpapi.core import app as core_app
    monkeypatch.setattr(
        core_app.App, "new", classmethod(lambda cls, is_visible=True: sentinel)
    )
    assert new_document() is sentinel


def test_new_document_wraps_com_error(monkeypatch):
    from hwpapi import errors as _errors
    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )
    from hwpapi.core import app as core_app

    def _raise(cls, is_visible=True):
        raise _FakeComError("nope")

    monkeypatch.setattr(core_app.App, "new", classmethod(_raise))
    with pytest.raises(FileIOError):
        new_document()


# ---------------------------------------------------------------------
# export_pdf / export_image / export_text
# (v3: ``app.docs.active.save(path, format=...)``)
# ---------------------------------------------------------------------

def _mock_app_with_active_doc(returned_path: str = "C:/out.pdf"):
    app = MagicMock()
    app.docs.active.save.return_value = returned_path
    return app


def test_export_pdf_saves_with_pdf_format():
    app = _mock_app_with_active_doc("C:/out.pdf")
    result = export_pdf(app, "out.pdf")
    app.docs.active.save.assert_called_once_with("out.pdf", format="PDF")
    assert result == "C:/out.pdf"


def test_export_image_png_uses_png_format():
    app = _mock_app_with_active_doc("C:/img.png")
    export_image(app, "img.png")
    app.docs.active.save.assert_called_once_with("img.png", format="PNG")


def test_export_image_defaults_to_bmp():
    app = _mock_app_with_active_doc("C:/img.bmp")
    export_image(app, "img.bmp")
    app.docs.active.save.assert_called_once_with("img.bmp", format="BMP")


def test_export_image_ignores_page_for_now():
    app = _mock_app_with_active_doc("C:/img.png")
    export_image(app, "img.png", page=2)
    app.docs.active.save.assert_called_once()


def test_export_text_saves_with_text_format():
    app = _mock_app_with_active_doc("C:/out.txt")
    export_text(app, "out.txt")
    app.docs.active.save.assert_called_once_with("out.txt", format="TEXT")


def test_export_pdf_wraps_com_error(monkeypatch):
    from hwpapi import errors as _errors
    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )
    app = MagicMock()
    app.docs.active.save.side_effect = _FakeComError("bad")
    with pytest.raises(FileIOError):
        export_pdf(app, "out.pdf")
