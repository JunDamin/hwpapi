"""
Core API for hwpapi v2: :class:`App` and re-exported engine handles.

- :mod:`hwpapi.low.engine`  — :class:`Engine`, :class:`Engines`, :class:`Apps`
- :mod:`hwpapi.core.app`   — :class:`App` (slim v2 facade, ≤15 public members)

All public names are re-exported at package root:

    from hwpapi import App  # preferred
    from hwpapi.core import App  # also works
"""
from __future__ import annotations

from hwpapi.low.engine import Engine, Engines, Apps
from .app import App

__all__ = ["Engine", "Engines", "Apps", "App"]
