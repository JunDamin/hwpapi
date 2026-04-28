"""
:mod:`hwpapi.collections.paragraphs` — ParagraphCollection.

HWP exposes paragraphs through ``XHwpDocuments.Active_XHwpDocument.
Section(0).Paragraph(i)``. Length is computed via section paragraph
counts; ordinal subscript returns a lightweight :class:`Paragraph`
value object.

Named enumeration is intentionally not supported — HWP paragraphs have
no natural names. ``names()`` returns index strings (``"0"``, ``"1"``,
…) to satisfy the Protocol.

Phase 4 enriches :class:`Paragraph` with ``.style``, ``.charshape``,
``.parashape`` and ``.runs``. Shape properties delegate to
:mod:`hwpapi.low.parametersets` via the ``CharShape`` and
``ParagraphShape`` actions on the owning :class:`App`. :class:`Run`
is a lightweight slice of a paragraph's text + charshape.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Paragraph", "ParagraphCollection", "Run"]


class Run:
    """
    A run — contiguous slice of a :class:`Paragraph` sharing a
    :class:`~hwpapi.low.parametersets.CharShape`.

    Phase 4 ships a single-run stub covering the whole paragraph; a
    future phase can split on CharShape boundaries by walking the
    characters.

    Parameters
    ----------
    paragraph : Paragraph
        Owning paragraph value object.
    start, end : int
        Character offsets into :attr:`Paragraph.text` (Python slice
        semantics). ``end == -1`` means "to end of paragraph".

    Attributes are evaluated lazily; constructing a :class:`Run` does
    not touch COM.
    """

    __slots__ = ("_paragraph", "start", "end")

    def __init__(self, paragraph: "Paragraph", start: int = 0, end: int = -1) -> None:
        self._paragraph = paragraph
        self.start = int(start)
        self.end = int(end)

    @property
    def paragraph(self) -> "Paragraph":
        """The :class:`Paragraph` this run belongs to."""
        return self._paragraph

    @property
    def text(self) -> str:
        """Substring of the paragraph text covered by this run."""
        text = self._paragraph.text
        if self.end == -1:
            return text[self.start:]
        return text[self.start:self.end]

    @property
    def charshape(self):
        """Best-effort :class:`CharShape` — currently the paragraph's."""
        return self._paragraph.charshape

    def __len__(self) -> int:
        if self.end == -1:
            return max(0, len(self._paragraph.text) - self.start)
        return max(0, self.end - self.start)

    def __repr__(self) -> str:
        end = "end" if self.end == -1 else str(self.end)
        return f"Run(para=#{self._paragraph.index}, {self.start}:{end})"


class Paragraph:
    """Value object for a single paragraph (ordinal only)."""

    __slots__ = ("_app", "index", "_raw")

    def __init__(self, app: "App", index: int, raw=None) -> None:
        self._app = app
        self.index = index
        self._raw = raw

    # ------------------------------------------------------------------
    # Text / style
    # ------------------------------------------------------------------

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

    @property
    def style(self) -> str:
        """Paragraph style name (e.g. "바탕글"). Empty string on failure."""
        raw = self._raw
        if raw is None:
            return ""
        for attr in ("StyleName", "Style"):
            try:
                v = getattr(raw, attr, None)
                if callable(v):
                    v = v()
                if v is None:
                    continue
                return str(v)
            except Exception:
                continue
        return ""

    # ------------------------------------------------------------------
    # Shape delegation to hwpapi.low.parametersets
    # ------------------------------------------------------------------

    def _action_pset(self, action_name: str):
        """Return ``app.actions.<action_name>.pset`` or ``None``."""
        try:
            actions = getattr(self._app, "actions", None)
            if actions is None:
                return None
            action = getattr(actions, action_name, None)
            if action is None:
                return None
            return getattr(action, "pset", None)
        except Exception:
            return None

    @property
    def charshape(self):
        """
        :class:`~hwpapi.low.parametersets.CharShape` snapshot for this
        paragraph. Returns the pset wrapped by
        ``app.actions.CharShape`` — shared with the current caret
        position. ``None`` when the action is unavailable.
        """
        return self._action_pset("CharShape")

    @property
    def parashape(self):
        """
        :class:`~hwpapi.low.parametersets.ParaShape` snapshot for this
        paragraph. Delegates to ``app.actions.ParagraphShape.pset``.
        """
        return self._action_pset("ParagraphShape")

    @property
    def runs(self) -> List[Run]:
        """
        Runs covering this paragraph. Phase 4 returns a single-element
        list spanning the full paragraph; later phases may split on
        CharShape boundaries.
        """
        # Phase 4+: refine by walking characters and emitting a Run per
        # CharShape-contiguous region. Single run is correct for the
        # common case of uniformly-formatted paragraphs.
        return [Run(self, 0, -1)]

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
