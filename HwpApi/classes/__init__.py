"""
Accessor and shape classes for hwpapi.

Split into:
- :mod:`~hwpapi.classes.accessors`: Move/Cell/Table/Page accessors
- :mod:`~hwpapi.classes.shapes`: Character/Paragraph/PageShape dataclasses + wrappers

All public names are re-exported at the package root for backward compatibility:

    from hwpapi.classes import MoveAccessor, CharShape, Paragraph, ...
"""
from __future__ import annotations

from .accessors import MoveAccessor, CellAccessor, TableAccessor, PageAccessor
from .shapes import Character, CharShape, Paragraph, ParaShape, PageShape
from .styles import StylesAccessor, Style
from .controls import ControlsAccessor, Control
from .fields import Fields, Field, Bookmarks, Hyperlinks, Hyperlink

__all__ = [
    'MoveAccessor', 'CellAccessor', 'TableAccessor', 'PageAccessor',
    'Character', 'CharShape', 'Paragraph', 'ParaShape', 'PageShape',
    'StylesAccessor', 'Style',
    'ControlsAccessor', 'Control',
    'Fields', 'Field', 'Bookmarks', 'Hyperlinks', 'Hyperlink',
]
