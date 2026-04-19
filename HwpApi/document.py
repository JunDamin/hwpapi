"""
:mod:`hwpapi.document` — Phase 2 ``Document`` facade.

``Document`` is the **primary v2 surface** for per-document operations:
text read/write, selection, clipboard, undo/redo, find-replace, page
navigation, and the seven collection accessors (``paragraphs``,
``tables``, ``fields``, ``bookmarks``, ``hyperlinks``, ``images``,
``styles``).

Canonical usage::

    app = App()
    app.doc.text              # full-document text
    app.doc.fields['name']    # FieldCollection lookup
    app.doc.select_all()
    app.doc.insert_text('…')

Phase 2 intentionally keeps this class a **thin facade**: every method
and collection property delegates to the backing :class:`App` or to
existing domain wrappers (``app.fields``, ``app.bookmarks`` …).
Dedicated :mod:`hwpapi.collections` classes land in Phase 3 (PRD
P3-00x). When that lands, ``Document`` becomes the implementation site
and the corresponding ``App`` methods are removed (PRD P2-002).

Design notes
------------
- ``Document`` holds a **reference** to the owning :class:`App`; it does
  not own its own COM handle. The engine is the single source of truth.
- All seven collection properties are lazy: the wrapper object is built
  on first access and cached via :func:`functools.cached_property`
  (same instance on repeated access).
- Delegated methods forward ``*args, **kwargs`` verbatim so behaviour
  matches v1 during Phase 2. The methods on ``App`` are the current
  implementation; once they are removed, the delegates here will be
  promoted to real implementations.

See also
--------
- ``docs/design/app-member-audit.md`` — classifies 35 ``App`` members
  as ``move_to_Document`` (this file's eventual home).
- ``docs/design/hwpapi_v2_redesign.md`` — Phase 2 redesign plan.
"""
from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Document"]


class _PlaceholderCollection:
    """
    Minimal stand-in for collections that have no v1 wrapper yet.

    Returned by :attr:`Document.paragraphs` and :attr:`Document.tables`
    during Phase 2. Phase 3 (PRD P3-00x) replaces this with a real
    :class:`ParagraphCollection` / :class:`TableCollection`.

    The object is intentionally minimal — it is a usable placeholder,
    not a broken one. It exposes ``len(...)`` == 0 and iterates empty
    so early callers can introspect the shape without crashing.
    """

    __slots__ = ("_name", "_doc")

    def __init__(self, name: str, doc: "Document") -> None:
        self._name = name
        self._doc = doc

    def __repr__(self) -> str:
        return (
            f"<{self._name} collection — full Phase 3 implementation "
            f"pending (placeholder on {self._doc!r})>"
        )

    def __len__(self) -> int:
        return 0

    def __iter__(self):
        return iter(())

    def __bool__(self) -> bool:
        # Truthy placeholder — distinguishes "property returns something"
        # from "property returned None". Real collections will override.
        return True


class Document:
    """
    Per-document facade for HWP operations.

    Obtained via :attr:`App.doc`. Document is a **thin proxy** around
    the owning :class:`App` during Phase 2; see module docstring for
    migration notes.

    Parameters
    ----------
    app : App
        The owning :class:`hwpapi.core.app.App` instance. ``Document``
        keeps a reference but does not own a COM handle.

    Examples
    --------
    >>> app = App()
    >>> doc = app.doc              # same instance on every access
    >>> doc.text                    # full-document text (delegates)
    >>> doc.fields['name'] = '홍길동'
    >>> doc.select_all()
    >>> doc.find_text('foo')
    """

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Collection properties — lazy, cached, same instance every access.
    # ------------------------------------------------------------------

    @cached_property
    def fields(self) -> Any:
        """Fields collection (누름틀). Delegates to ``app.fields``."""
        return self._app.fields

    @cached_property
    def bookmarks(self) -> Any:
        """Bookmark collection (책갈피). Delegates to ``app.bookmarks``."""
        return self._app.bookmarks

    @cached_property
    def hyperlinks(self) -> Any:
        """Hyperlink collection. Delegates to ``app.hyperlinks``."""
        return self._app.hyperlinks

    @cached_property
    def images(self) -> Any:
        """Image collection. Delegates to ``app.images``."""
        return self._app.images

    @cached_property
    def styles(self) -> Any:
        """Paragraph-style collection. Delegates to ``app.styles``."""
        return self._app.styles

    @cached_property
    def paragraphs(self) -> Any:
        """
        Paragraph collection.

        Phase 2 placeholder: real :class:`ParagraphCollection` lands
        in Phase 3. Returns a minimal usable object (see
        :class:`_PlaceholderCollection`).
        """
        return _PlaceholderCollection("paragraphs", self)

    @cached_property
    def tables(self) -> Any:
        """
        Table collection.

        Phase 2 placeholder: real :class:`TableCollection` lands in
        Phase 3. Returns a minimal usable object (see
        :class:`_PlaceholderCollection`).
        """
        return _PlaceholderCollection("tables", self)

    # ------------------------------------------------------------------
    # Document-level property delegates.
    # ------------------------------------------------------------------

    @property
    def text(self) -> str:
        """Full document text. Delegates to ``app.text``."""
        return self._app.text

    @text.setter
    def text(self, value: str) -> None:
        self._app.text = value

    @property
    def page_count(self) -> int:
        """Total page count. Delegates to ``app.page_count``."""
        return self._app.page_count

    @property
    def current_page(self) -> int:
        """Cursor-relative page number. Delegates to ``app.current_page``."""
        return self._app.current_page

    @property
    def selection(self) -> str:
        """Current selection text. Delegates to ``app.selection``."""
        return self._app.selection

    # ------------------------------------------------------------------
    # High-value method delegates.
    #
    # Each forwards verbatim to the App method of the same name.
    # Phase 2: behaviour is identical to v1.
    # Phase 3+: the App methods are removed (PRD P2-002) and these
    # become the implementation site.
    # ------------------------------------------------------------------

    def select_all(self, *args, **kwargs):
        """Select the entire document. Delegates to ``app.select_all``."""
        return self._app.select_all(*args, **kwargs)

    def get_text(self, *args, **kwargs):
        """Read text from the document. Delegates to ``app.get_text``."""
        return self._app.get_text(*args, **kwargs)

    def insert_text(self, *args, **kwargs):
        """Insert text at cursor. Delegates to ``app.insert_text``."""
        return self._app.insert_text(*args, **kwargs)

    def undo(self, *args, **kwargs):
        """Undo. Delegates to ``app.undo``."""
        return self._app.undo(*args, **kwargs)

    def redo(self, *args, **kwargs):
        """Redo. Delegates to ``app.redo``."""
        return self._app.redo(*args, **kwargs)

    def copy(self, *args, **kwargs):
        """Clipboard copy. Delegates to ``app.copy``."""
        return self._app.copy(*args, **kwargs)

    def cut(self, *args, **kwargs):
        """Clipboard cut. Delegates to ``app.cut``."""
        return self._app.cut(*args, **kwargs)

    def paste(self, *args, **kwargs):
        """Clipboard paste. Delegates to ``app.paste``."""
        return self._app.paste(*args, **kwargs)

    def delete(self, *args, **kwargs):
        """Delete current selection. Delegates to ``app.delete``."""
        return self._app.delete(*args, **kwargs)

    def find_text(self, *args, **kwargs):
        """Find text in document. Delegates to ``app.find_text``."""
        return self._app.find_text(*args, **kwargs)

    def replace_all(self, *args, **kwargs):
        """Document-wide find/replace. Delegates to ``app.replace_all``."""
        return self._app.replace_all(*args, **kwargs)

    def insert_line_break(self, *args, **kwargs):
        """Insert line break. Delegates to ``app.insert_line_break``."""
        return self._app.insert_line_break(*args, **kwargs)

    def insert_page_break(self, *args, **kwargs):
        """Insert page break. Delegates to ``app.insert_page_break``."""
        return self._app.insert_page_break(*args, **kwargs)

    # ------------------------------------------------------------------
    # Introspection
    # ------------------------------------------------------------------

    def __repr__(self) -> str:
        return f"<Document app={self._app!r}>"
