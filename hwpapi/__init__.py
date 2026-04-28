"""
hwpapi — Pythonic automation for HWP (한컴 오피스) via win32com.

Public surface is intentionally tiny. Everything else lives in submodules:

- :class:`hwpapi.App`           — slim application facade
- :class:`hwpapi.Document`      — per-document facade (via ``app.doc``)
- :mod:`hwpapi.collections`     — fields, bookmarks, hyperlinks, images,
                                  paragraphs, styles, tables
- :mod:`hwpapi.context`         — charshape_scope, parashape_scope, styled_text
- :mod:`hwpapi.io`              — open_file, new_document, export_*
- :mod:`hwpapi.errors`          — HwpApiError hierarchy + wrap_com_error
- :mod:`hwpapi.units`           — mm/cm/inch/pt ↔ HWPUNIT helpers
- :mod:`hwpapi.low`             — raw actions / parametersets / engine (escape hatch)

See https://JunDamin.github.io/hwpapi for the full documentation site.
"""
from __future__ import annotations

__version__ = "2.0.0"

from .core.app import App
from .document import Document

__all__ = ["App", "Document", "__version__"]
