"""Unit tests for :mod:`hwpapi.io.open` and :mod:`hwpapi.io.export`."""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from hwpapi.errors import FileIOError
from hwpapi.io.open import open_file, new_document
from hwpapi.io.export import export_pdf, export_image, export_text


# ---------------------------------------------------------------------
# open_file
# ---------------------------------------------------------------------

def test_open_file_delegates_to_app_open():
    """``open_file(app, path)`` calls ``app.open(path, None)`` once."""
    app = MagicMock()
    app.open.return_value = "C:/abs/x.hwp"

    result = open_file(app, "x.hwp")

    app.open.assert_called_once_with("x.hwp", None)
    assert result == "C:/abs/x.hwp"


def test_open_file_forwards_format_argument():
    """Explicit ``format=`` is passed through to ``app.open``."""
    app = MagicMock()
    open_file(app, "x.pdf", format="PDF")
    app.open.assert_called_once_with("x.pdf", "PDF")


class _FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` on non-Windows runners."""


def test_open_file_wraps_com_error(monkeypatch):
    """A COM error from ``app.open`` becomes :class:`FileIOError`."""
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    app = MagicMock()
    app.open.side_effect = _FakeComError("bad path")

    with pytest.raises(FileIOError) as excinfo:
        open_file(app, "missing.hwp")

    # Path and underlying cause surface in the message.
    assert "missing.hwp" in str(excinfo.value)
    assert isinstance(excinfo.value.__cause__, _FakeComError)


# ---------------------------------------------------------------------
# new_document
# ---------------------------------------------------------------------

def test_new_document_calls_app_new(monkeypatch):
    """``new_document`` is a thin alias for :meth:`App.new`."""
    fake_app_cls = MagicMock()
    monkeypatch.setattr("hwpapi.core.app.App", fake_app_cls)

    new_document(is_visible=False)

    fake_app_cls.new.assert_called_once_with(is_visible=False)


def test_new_document_wraps_com_error(monkeypatch):
    """Connection-style failures during ``App.new`` raise FileIOError."""
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    fake_app_cls = MagicMock()
    fake_app_cls.new.side_effect = _FakeComError("dispatch failed")
    monkeypatch.setattr("hwpapi.core.app.App", fake_app_cls)

    with pytest.raises(FileIOError):
        new_document()


# ---------------------------------------------------------------------
# export_pdf / export_image / export_text
# ---------------------------------------------------------------------

def test_export_pdf_saves_with_pdf_format():
    app = MagicMock()
    app.save_as.return_value = "C:/out.pdf"

    result = export_pdf(app, "out.pdf")

    app.save_as.assert_called_once_with("out.pdf", format="PDF")
    assert result == "C:/out.pdf"


def test_export_image_png_uses_png_format():
    """An ``.png`` path gets format="PNG"."""
    app = MagicMock()
    export_image(app, "picture.png")
    app.save_as.assert_called_once_with("picture.png", format="PNG")


def test_export_image_defaults_to_bmp():
    """Non-png extensions fall back to BMP, which HWP always supports."""
    app = MagicMock()
    export_image(app, "picture.bmp")
    app.save_as.assert_called_once_with("picture.bmp", format="BMP")


def test_export_image_ignores_page_for_now():
    """``page=`` is accepted but currently a no-op — see docstring."""
    app = MagicMock()
    export_image(app, "x.png", page=3)
    app.save_as.assert_called_once_with("x.png", format="PNG")


def test_export_text_saves_with_text_format():
    app = MagicMock()
    export_text(app, "out.txt")
    app.save_as.assert_called_once_with("out.txt", format="TEXT")


def test_export_pdf_wraps_com_error(monkeypatch):
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    app = MagicMock()
    app.save_as.side_effect = _FakeComError("save failed")

    with pytest.raises(FileIOError) as excinfo:
        export_pdf(app, "out.pdf")

    assert "out.pdf" in str(excinfo.value)
    assert "export_pdf" in str(excinfo.value)
