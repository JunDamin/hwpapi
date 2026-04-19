"""
:mod:`hwpapi.context` — formatting/state context managers.

Phase 5 of the v2 redesign lifts the v1 ``App.charshape_scope`` /
``parashape_scope`` / ``styled_text`` helpers out of the monolithic App
into this dedicated, module-level package. The App surface stays at
14 + ``actions`` members per :file:`docs/design/app-member-audit.md`.

Re-exported names::

    charshape_scope   # context manager — char formatting block
    parashape_scope   # context manager — paragraph formatting block
    styled_text       # one-shot — insert styled text then restore

See :mod:`hwpapi.context.scopes` for details and the decision tree.
"""
from __future__ import annotations

from .scopes import charshape_scope, parashape_scope, styled_text

__all__ = ["charshape_scope", "parashape_scope", "styled_text"]
