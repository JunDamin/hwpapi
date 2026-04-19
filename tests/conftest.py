"""Shared pytest fixtures for the hwpapi test suite.

The only cross-cutting fixture here is :func:`_autoyes_hwp_dialogs`, which
patches :class:`hwpapi.App` so that every ``App()`` instantiated during a
pytest run auto-answers **Yes / OK / Abort** (the first button) for every
HWP dialog category.

Without this, HWP-integration tests (``smoke_*.py``, ``test_all_*.py``)
would block indefinitely on modal dialogs such as save-prompts or
overwrite-confirmations.

The patch is scoped to the test session only — production ``App`` users
keep their normal interactive dialog behaviour.
"""
from __future__ import annotations

import pytest


SILENCE_ALL_YES = 0x111111
"""HWP ``SetMessageBoxMode`` mask: first button on every dialog category.

Nibble layout (4 bits per dialog type):
    0xF      OK-only dialog            → 1 = OK
    0xF0     OK/Cancel dialog          → 1 = OK
    0xF00    Abort/Retry/Ignore dialog → 1 = Abort
    0xF000   Yes/No/Cancel dialog      → 1 = Yes
    0xF0000  Yes/No dialog             → 1 = Yes
    0xF00000 Retry/Cancel dialog       → 1 = Retry
"""


@pytest.fixture(autouse=True, scope="session")
def _autoyes_hwp_dialogs():
    """Auto-answer Yes on every HWP dialog raised during the test session.

    Patches :meth:`hwpapi.App.__init__` so that, after a real HWP engine
    has been bound, ``self.api.SetMessageBoxMode(SILENCE_ALL_YES)`` runs
    once per ``App`` instance. Non-HWP tests (those using ``App.__new__``
    or ``MagicMock``) never touch ``__init__`` and are therefore
    unaffected.
    """
    from hwpapi.core.app import App

    original_init = App.__init__

    def _init_with_autoyes(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        try:
            self.api.SetMessageBoxMode(SILENCE_ALL_YES)
        except Exception:
            pass

    App.__init__ = _init_with_autoyes
    try:
        yield
    finally:
        App.__init__ = original_init
