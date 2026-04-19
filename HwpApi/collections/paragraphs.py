"""
:mod:`hwpapi.collections.paragraphs` — ParagraphCollection.

HWP exposes paragraphs through ``XHwpDocuments.Active_XHwpDocument.
Section(0).Paragraph(i)``. Length is computed via section paragraph
counts; ordinal subscript returns a lightweight :class:`Paragraph`
value object.

Named enumeration is intentionally not supported — HWP paragraphs have
no natural names. ``names()`` returns index strings (``"0"``, ``"1"``,
…) to satisfy the Protocol.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Paragraph", "ParagraphCollection"]


class Paragraph:
    """Value object for a single paragraph (ordinal only)."""

    __slots__ = ("_app", "index", "_raw")

    def __init__(self, app: "App", index: int, raw=None) -> None:
        self._app = app
        self.index = index
        self._raw = raw

    @property
    def text(self) -> str:
        """Paragraph text. Returns ``""`` when the backend cannot resolve it."""
        raw = self._raw
        if raw is None:
            return ""
        for attr in ("Text", "GetText"):
            try:
                v = getattr(raw, attr, None)
                if callable(v):
                    return str(v() or "")
                if v is not None:
                    return str(v)
            except Exception:
                continue
        return ""

    def __repr__(self) -> str:
        return f"Paragraph(#{self.index})"


class ParagraphCollection:
    """
    ``doc.paragraphs`` — ordinal-access paragraph collection.

    Iteration uses ``Section(0).Paragraph(i)`` on the active HWP
    document. When the underlying handle can't enumerate (e.g. no
    document open), the collection behaves as empty.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Internal COM access
    # ------------------------------------------------------------------

    def _section(self):
        impl = self._app.engine.impl
        try:
            return impl.XHwpDocuments.Active_XHwpDocument.Section(0)
        except Exception:
            return None

    def _count(self) -> int:
        sec = self._section()
        if sec is None:
            return 0
        for attr in ("Paragraphs", "ParagraphCount"):
            try:
                n = getattr(sec, attr, None)
                if callable(n):
                    return int(n() or 0)
                if n is not None:
                    return int(n)
            except Exception:
                continue
        return 0

    def _raw_paragraph(self, i: int):
        sec = self._section()
        if sec is None:
            return None
        try:
            return sec.Paragraph(i)
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        """Ordinal strings — paragraphs have no intrinsic names."""
        return [str(i) for i in range(self._count())]

    def __iter__(self) -> Iterator[Paragraph]:
        for i in range(self._count()):
            yield Paragraph(self._app, i, self._raw_paragraph(i))

    def __len__(self) -> int:
        return self._count()

    def __contains__(self, key) -> bool:
        if isinstance(key, Paragraph):
            return 0 <= key.index < self._count()
        if isinstance(key, int):
            return 0 <= key < self._count()
        if isinstance(key, str):
            return key in self.names()
        return False

    def __getitem__(self, key) -> Paragraph:
        if isinstance(key, int):
            n = self._count()
            if key < 0:
                key += n
            if not (0 <= key < n):
                raise IndexError(f"Paragraph index {key} out of range")
            return Paragraph(self._app, key, self._raw_paragraph(key))
        raise TypeError(
            f"Paragraph key must be int, got {type(key).__name__}"
        )

    def filter(
        self, predicate: Callable[[Paragraph], bool]
    ) -> List[Paragraph]:
        return [p for p in self if predicate(p)]

    def __repr__(self) -> str:
        try:
            n = len(self)
        except Exception:
            n = "?"
        return f"ParagraphCollection(count={n})"
