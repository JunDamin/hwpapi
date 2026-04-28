"""
:mod:`hwpapi.errors` — lean exception hierarchy for hwpapi v2.

Every public hwpapi call (``App.open``, ``App.save``, ``io.open_file``,
``io.export_pdf`` …) wraps raw ``pywintypes.com_error`` and other COM-level
failures behind one of the subclasses below. End users only ever see
:class:`HwpApiError` descendants.

Hierarchy::

    HwpApiError                    — base for everything
    ├── ConnectionError            — can't dispatch HwpObject / engine dead
    ├── ActionFailedError          — HAction.Run / HAction.Execute returned False
    ├── InvalidArgumentError       — bad path, unknown format, wrong arg shape
    └── FileIOError                — open/save/export failed at the FS level

The :func:`wrap_com_error` context manager is a lightweight helper used by
:mod:`hwpapi.io` and :mod:`hwpapi.context.scopes` to re-raise raw COM
errors as :class:`ActionFailedError` with the offending action's name.

Note
----
``ConnectionError`` intentionally shadows the built-in — inside hwpapi
code always import from this module (``from hwpapi.errors import
ConnectionError``). Outside code that needs the built-in should keep
using the stdlib name.
"""
from __future__ import annotations

from contextlib import contextmanager
from typing import Optional

__all__ = [
    "HwpApiError",
    "ConnectionError",
    "ActionFailedError",
    "InvalidArgumentError",
    "FileIOError",
    "wrap_com_error",
]


class HwpApiError(Exception):
    """Base for every hwpapi-raised exception."""


class ConnectionError(HwpApiError):
    """Raised when the HWP COM object can't be dispatched or has died."""


class ActionFailedError(HwpApiError):
    """Raised when ``HAction.Run`` / ``HAction.Execute`` returns falsy."""


class InvalidArgumentError(HwpApiError):
    """Raised for unknown formats, malformed paths, or bad kwargs."""


class FileIOError(HwpApiError):
    """Raised when an open/save/export fails at the filesystem/COM edge."""


def _iter_com_error_types():
    """Yield COM-error exception types we should wrap, best-effort.

    ``pywintypes`` is only available on Windows in a real HWP environment;
    on non-HWP test runners (Linux/mac) we gracefully degrade to just
    :class:`Exception`. Returning a tuple keeps ``except`` syntax happy.
    """
    types = []
    try:
        import pywintypes  # type: ignore
        types.append(pywintypes.com_error)
    except Exception:
        pass
    return tuple(types) or (Exception,)


@contextmanager
def wrap_com_error(action_name: str):
    """Re-raise COM errors inside the block as :class:`ActionFailedError`.

    Parameters
    ----------
    action_name : str
        Short label for the failing operation — included in the error
        message so callers can tell ``FileOpen`` from ``FileSaveAsPdf``.

    Example
    -------
    >>> with wrap_com_error("FileOpen"):
    ...     app.api.Open(path)
    """
    com_types = _iter_com_error_types()
    try:
        yield
    except HwpApiError:
        # Already wrapped — don't double-wrap.
        raise
    except com_types as exc:
        raise ActionFailedError(
            f"{action_name} failed: {exc!r}"
        ) from exc
