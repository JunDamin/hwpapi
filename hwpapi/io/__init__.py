"""
:mod:`hwpapi.io` — thin, error-wrapping file-IO shortcuts.

The canonical open/save APIs still live on :class:`hwpapi.App` (``open``,
``save``, ``save_as``). This package adds two small conveniences:

1. **Error wrapping** — raw ``pywintypes.com_error`` bubbling up from
   HWP becomes :class:`~hwpapi.errors.FileIOError` with a helpful
   message including the path and format.
2. **Export shortcuts** — named helpers for the common export targets
   (PDF / image / text) so callers don't have to remember HWP's
   magic format strings.

Public names::

    from hwpapi.io.open   import open_file, new_document
    from hwpapi.io.export import export_pdf, export_image, export_text
"""
from __future__ import annotations

from .open import open_file, new_document
from .export import export_pdf, export_image, export_text

__all__ = [
    "open_file",
    "new_document",
    "export_pdf",
    "export_image",
    "export_text",
]
