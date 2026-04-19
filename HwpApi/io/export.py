"""
:mod:`hwpapi.io.export` — named export shortcuts (PDF / image / text).

Each helper wraps :meth:`hwpapi.App.save_as` with the right HWP format
string and re-raises COM errors as :class:`~hwpapi.errors.FileIOError`
so callers get a uniform, package-specific exception type.
"""
from __future__ import annotations

from typing import Optional, TYPE_CHECKING

from hwpapi import errors as _errors
from hwpapi.errors import FileIOError
from hwpapi.logging import get_logger

if TYPE_CHECKING:  # pragma: no cover
    from hwpapi.core.app import App

__all__ = ["export_pdf", "export_image", "export_text"]

logger = get_logger("io.export")


def _save_as(app: "App", path: str, format: str, label: str) -> str:
    """Shared body — :meth:`App.save_as` wrapped in a COM-error fence."""
    com_types = _errors._iter_com_error_types()
    try:
        result = app.save_as(path, format=format)
    except com_types as exc:
        raise FileIOError(
            f"{label}({path!r}) failed: {exc!r}"
        ) from exc
    return result


def export_pdf(app: "App", path: str) -> str:
    """Export the active document to PDF at ``path``.

    Thin wrapper around ``app.save_as(path, format="PDF")`` that turns
    raw COM errors into :class:`FileIOError`.
    """
    return _save_as(app, path, "PDF", "export_pdf")


def export_image(
    app: "App",
    path: str,
    page: Optional[int] = None,
) -> str:
    """Export the active document (or one page) to an image file.

    Parameters
    ----------
    app : hwpapi.App
    path : str | Path
        Destination path. Extension suggests the target format — the
        ``.png`` path passes ``format="PNG"`` to HWP, everything else
        falls through to ``BMP`` which HWP always supports.
    page : int, optional
        Reserved for future per-page export. Currently HWP's
        ``SaveBlockAction`` path doesn't accept a page index, so we log
        the request and delegate to the whole-document export.

    Raises
    ------
    FileIOError
        On any COM-level failure.
    """
    # Extension sniff so users can call ``export_image(app, "x.png")``.
    lower = str(path).lower()
    fmt = "PNG" if lower.endswith(".png") else "BMP"

    if page is not None:
        logger.debug(
            f"export_image: per-page export not yet wired; "
            f"ignoring page={page!r}"
        )

    return _save_as(app, path, fmt, "export_image")


def export_text(app: "App", path: str) -> str:
    """Export the active document to a plain-text file at ``path``.

    Thin wrapper around ``app.save_as(path, format="TEXT")`` that turns
    raw COM errors into :class:`FileIOError`.
    """
    return _save_as(app, path, "TEXT", "export_text")
