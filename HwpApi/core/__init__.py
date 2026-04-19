"""
Core API for hwpapi: App, Engine, and related.

Split into:
- :mod:`~hwpapi.low.engine`: Engine/Engines/Apps (low-level, relocated in v2 Phase 1)
- :mod:`~hwpapi.core.app`: App class (main user entry)

All public names are re-exported at package root for backward compatibility:

    from hwpapi.core import App  # as before
"""
from __future__ import annotations

from hwpapi.low.engine import Engine, Engines, Apps
from .app import App, Document, Documents, move_to_line

__all__ = ['Engine', 'Engines', 'Apps', 'App',
           'Document', 'Documents', 'move_to_line']
