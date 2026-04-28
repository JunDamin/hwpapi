"""
:mod:`hwpapi.io.open` — open/new document helpers.

Thin wrappers around :meth:`hwpapi.App.open` and :meth:`hwpapi.App.new`
that translate raw COM errors into :class:`~hwpapi.errors.FileIOError`
so callers never have to ``except pywintypes.com_error``.
"""
from __future__ import annotations

from typing import Optional, TYPE_CHECKING

from hwpapi import errors as _errors
from hwpapi.errors import FileIOError
from hwpapi.logging import get_logger

if TYPE_CHECKING:  # pragma: no cover
    from hwpapi.core.app import App

__all__ = ["open_file", "new_document"]

logger = get_logger("io.open")


def open_file(app: "App", path: str, format: Optional[str] = None) -> str:
    """Open ``path`` in ``app`` and return the absolute path opened.

    Parameters
    ----------
    app : hwpapi.App
        Live App instance.
    path : str | Path
        File path (relative resolved against cwd).
    format : str, optional
        HWP format string (``"HWP"``, ``"PDF"``, ``"HWPX"`` …). ``None``
        lets HWP auto-detect.

    Raises
    ------
    FileIOError
        Wraps any COM error raised by :meth:`App.open`, with the path
        included in the message.
    """
    com_types = _errors._iter_com_error_types()
    try:
        return app.open(path, format)
    except com_types as exc:
        raise FileIOError(
            f"open_file({path!r}, format={format!r}) failed: {exc!r}"
        ) from exc


def new_document(is_visible: bool = True) -> "App":
    """Create a fresh :class:`~hwpapi.App` backed by a blank document.

    Thin helper around :meth:`hwpapi.App.new` — prefer this when reading
    linearly in io-heavy scripts so the ``hwpapi.io`` namespace exposes
    the full open/new/export lifecycle.

    Raises
    ------
    FileIOError
        If the underlying :meth:`App.new` raises a COM error (rare —
        usually means HWP can't be dispatched at all).
    """
    # Lazy import — avoids a core.app → io.open → core.app cycle.
    from hwpapi.core.app import App

    com_types = _errors._iter_com_error_types()
    try:
        return App.new(is_visible=is_visible)
    except com_types as exc:
        raise FileIOError(f"new_document() failed: {exc!r}") from exc
