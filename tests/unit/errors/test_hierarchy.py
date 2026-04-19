"""Unit tests for :mod:`hwpapi.errors` — Phase 5 exception hierarchy."""
from __future__ import annotations

import pytest

from hwpapi.errors import (
    HwpApiError,
    ConnectionError,
    ActionFailedError,
    InvalidArgumentError,
    FileIOError,
    wrap_com_error,
)


# ---------------------------------------------------------------------
# Class hierarchy
# ---------------------------------------------------------------------

@pytest.mark.parametrize(
    "cls",
    [ConnectionError, ActionFailedError, InvalidArgumentError, FileIOError],
)
def test_subclass_of_hwpapi_error(cls):
    """Every subclass descends from :class:`HwpApiError`."""
    assert issubclass(cls, HwpApiError)


def test_hwpapi_error_is_exception():
    assert issubclass(HwpApiError, Exception)


def test_subclasses_are_distinct():
    """The four subclasses are not aliases of each other."""
    subs = {ConnectionError, ActionFailedError, InvalidArgumentError, FileIOError}
    assert len(subs) == 4


# ---------------------------------------------------------------------
# wrap_com_error helper
# ---------------------------------------------------------------------

class _FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` on non-Windows test runners."""


def test_wrap_com_error_reraises_as_action_failed(monkeypatch):
    """A COM-shaped exception inside the block becomes ActionFailedError."""
    # Pretend pywintypes.com_error is our local fake so wrap_com_error picks
    # it up inside its ``except`` tuple.
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    with pytest.raises(ActionFailedError) as excinfo:
        with wrap_com_error("SomeAction"):
            raise _FakeComError("kaboom")

    # Message includes the action name for debuggability.
    assert "SomeAction" in str(excinfo.value)
    # Chains the original cause.
    assert isinstance(excinfo.value.__cause__, _FakeComError)


def test_wrap_com_error_passes_hwpapi_error_untouched(monkeypatch):
    """Already-wrapped :class:`HwpApiError` is not re-wrapped."""
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    original = FileIOError("already wrapped")
    with pytest.raises(FileIOError) as excinfo:
        with wrap_com_error("SomeAction"):
            raise original
    assert excinfo.value is original


def test_wrap_com_error_lets_unrelated_exceptions_propagate(monkeypatch):
    """Exceptions outside the COM family propagate unchanged."""
    from hwpapi import errors as _errors

    monkeypatch.setattr(
        _errors, "_iter_com_error_types", lambda: (_FakeComError,)
    )

    with pytest.raises(ValueError):
        with wrap_com_error("SomeAction"):
            raise ValueError("not a COM error")


def test_wrap_com_error_noop_when_no_exception():
    """The block completes cleanly when nothing raises."""
    marker = []
    with wrap_com_error("SomeAction"):
        marker.append("ran")
    assert marker == ["ran"]
