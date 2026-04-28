"""
:mod:`hwpapi.document` — Phase 3 ``Document`` facade.

``Document`` is the **primary v2 surface** for per-document operations.
Obtained via :attr:`hwpapi.App.doc`.

Phase 3 wires each ``@cached_property`` to a real
:mod:`hwpapi.collections` class. The engine is the single source of
truth — :class:`Document` never owns its own COM handle.

Canonical usage::

    app = App()
    doc = app.doc                      # cached Document
    doc.fields['name'] = '홍길동'       # FieldCollection
    for bookmark in doc.bookmarks: ...
"""
from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from hwpapi.core.app import App
    from hwpapi.collections.bookmarks import BookmarkCollection
    from hwpapi.collections.fields import FieldCollection
    from hwpapi.collections.hyperlinks import HyperlinkCollection
    from hwpapi.collections.images import ImageCollection
    from hwpapi.collections.paragraphs import ParagraphCollection
    from hwpapi.collections.styles import StyleCollection
    from hwpapi.collections.tables import TableCollection

__all__ = ["Document"]


class Document:
    """
    Per-document façade for HWP operations.

    Obtained via :attr:`hwpapi.App.doc`. Holds a reference to the
    owning :class:`App`; does **not** own a COM handle.

    Parameters
    ----------
    app : hwpapi.App
        The owning :class:`~hwpapi.core.app.App` instance.

    Examples
    --------
    >>> app = App()
    >>> doc = app.doc              # same instance on every access
    >>> len(doc.fields)            # number of 누름틀 fields
    0
    """

    __slots__ = (
        "_app",
        "__dict__",  # cached_property needs a real __dict__
    )

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Collection properties (Phase 3 — wired to hwpapi.collections.*).
    # ------------------------------------------------------------------

    @cached_property
    def fields(self) -> "FieldCollection":
        """Field collection (누름틀)."""
        from hwpapi.collections.fields import FieldCollection
        return FieldCollection(self._app)

    @cached_property
    def bookmarks(self) -> "BookmarkCollection":
        """Bookmark collection (책갈피)."""
        from hwpapi.collections.bookmarks import BookmarkCollection
        return BookmarkCollection(self._app)

    @cached_property
    def hyperlinks(self) -> "HyperlinkCollection":
        """Hyperlink collection."""
        from hwpapi.collections.hyperlinks import HyperlinkCollection
        return HyperlinkCollection(self._app)

    @cached_property
    def images(self) -> "ImageCollection":
        """Image collection."""
        from hwpapi.collections.images import ImageCollection
        return ImageCollection(self._app)

    @cached_property
    def styles(self) -> "StyleCollection":
        """Paragraph-style collection (minimal, see module docstring)."""
        from hwpapi.collections.styles import StyleCollection
        return StyleCollection(self._app)

    @cached_property
    def paragraphs(self) -> "ParagraphCollection":
        """Paragraph collection (ordinal access)."""
        from hwpapi.collections.paragraphs import ParagraphCollection
        return ParagraphCollection(self._app)

    @cached_property
    def tables(self) -> "TableCollection":
        """Table collection."""
        from hwpapi.collections.tables import TableCollection
        return TableCollection(self._app)

    # ------------------------------------------------------------------
    # Access to the underlying App (escape hatch for collections/elements).
    # ------------------------------------------------------------------

    @property
    def app(self) -> "App":
        """The owning :class:`App` — escape hatch."""
        return self._app

    def __repr__(self) -> str:
        return f"<Document app={self._app!r}>"
