"""
:mod:`hwpapi.document` — Phase 2 ``Document`` facade (scaffold).

``Document`` is the **primary v2 surface** for per-document operations.
Obtained via :attr:`hwpapi.App.doc`.

Phase 2 scope
-------------
The v2 slim :class:`hwpapi.App` no longer carries document-level
state (see ``docs/design/app-member-audit.md`` §1.2–§1.4). That
state lands on :class:`Document` and its collection/element helpers
in later phases:

- **Phase 3** (PRD P3-00x) — populate collection properties
  (``fields``, ``bookmarks``, ``hyperlinks``, ``images``, ``styles``,
  ``tables``, ``paragraphs``).
- **Phase 4** (PRD P4-00x) — populate element helpers
  (``cursor``, ``selection``, ``page`` …) and the rich method surface
  listed under ``move_to_Document`` in the audit.

For now :class:`Document` is a **thin scaffold** that holds a
reference to the owning :class:`App` and exposes seven lazy
collection-shaped placeholders so downstream code can introspect the
contract without crashing. The placeholders intentionally iterate
empty and ``len == 0`` so callers discover the "not yet implemented"
state the moment they try to use one.

The engine is the single source of truth — :class:`Document` never
owns its own COM handle.

Canonical usage (Phase 3 onward)::

    app = App()
    doc = app.doc                   # cached Document
    doc.fields['name'] = '홍길동'    # collection lookup (Phase 3)
    doc.select_all()                # document action (Phase 4)
"""
from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Document"]


class _PlaceholderCollection:
    """
    Minimal stand-in for collections that have no v2 implementation yet.

    Returned by every :class:`Document` collection property during
    Phase 2. Phase 3 replaces each placeholder with the real
    :mod:`hwpapi.collections` object.

    The placeholder is a **usable empty collection**, not a broken
    one. It exposes ``len(...) == 0`` and iterates empty so test-suite
    introspection works. Any ``__getitem__``/``__setitem__``/
    ``__delitem__`` call raises :class:`LookupError` with a pointer
    to the Phase 3 story.
    """

    __slots__ = ("_name", "_doc")

    def __init__(self, name: str, doc: "Document") -> None:
        self._name = name
        self._doc = doc

    def __repr__(self) -> str:
        return (
            f"<{self._name} collection — Phase 3 implementation pending "
            f"(placeholder on {self._doc!r})>"
        )

    def __len__(self) -> int:
        return 0

    def __iter__(self):
        return iter(())

    def __bool__(self) -> bool:
        # Truthy so `if doc.fields:` reflects "collection exists", not
        # "collection is non-empty". Matches Python container idiom
        # once the real class arrives.
        return True

    def _not_yet(self):
        raise LookupError(
            f"{self._name}: Phase 3 not implemented yet. "
            f"See PRD story P3-00x and hwpapi/collections/."
        )

    def __getitem__(self, key):
        self._not_yet()

    def __setitem__(self, key, value):
        self._not_yet()

    def __delitem__(self, key):
        self._not_yet()

    def __contains__(self, key) -> bool:
        return False


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
    >>> len(doc.fields)            # empty placeholder until Phase 3
    0
    """

    __slots__ = (
        "_app",
        "__dict__",  # cached_property needs a real __dict__
    )

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Collection properties (Phase 3 target — placeholders for now).
    # ------------------------------------------------------------------

    @cached_property
    def fields(self) -> Any:
        """Field collection (누름틀). Phase 3 placeholder."""
        return _PlaceholderCollection("fields", self)

    @cached_property
    def bookmarks(self) -> Any:
        """Bookmark collection (책갈피). Phase 3 placeholder."""
        return _PlaceholderCollection("bookmarks", self)

    @cached_property
    def hyperlinks(self) -> Any:
        """Hyperlink collection. Phase 3 placeholder."""
        return _PlaceholderCollection("hyperlinks", self)

    @cached_property
    def images(self) -> Any:
        """Image collection. Phase 3 placeholder."""
        return _PlaceholderCollection("images", self)

    @cached_property
    def styles(self) -> Any:
        """Paragraph-style collection. Phase 3 placeholder."""
        return _PlaceholderCollection("styles", self)

    @cached_property
    def paragraphs(self) -> Any:
        """Paragraph collection. Phase 3 placeholder."""
        return _PlaceholderCollection("paragraphs", self)

    @cached_property
    def tables(self) -> Any:
        """Table collection. Phase 3 placeholder."""
        return _PlaceholderCollection("tables", self)

    # ------------------------------------------------------------------
    # Access to the underlying App (escape hatch for Phase 3/4 impls).
    # ------------------------------------------------------------------

    @property
    def app(self) -> "App":
        """The owning :class:`App` — escape hatch."""
        return self._app

    def __repr__(self) -> str:
        return f"<Document app={self._app!r}>"
