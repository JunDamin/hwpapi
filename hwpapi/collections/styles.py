"""
:mod:`hwpapi.collections.styles` — StyleCollection (minimal).

HWP's COM surface does **not** expose a first-class paragraph style
enumeration; the v1 :mod:`hwpapi.classes.styles` accessor resorts to
HWPML export + XML regex parsing for its names list. That strategy
needs a live HWP runtime and a selection, so the Phase 3 collection
provides a **Protocol-compliant stub**:

* ``names()`` returns ``[]`` by default.
* Iteration yields nothing, ``len(...)`` is 0.
* ``__getitem__`` by string returns the :class:`Style` wrapper anyway —
  HWP's ``Style`` action accepts a style name and validates lazily, so
  callers can still ``doc.styles["제목 1"].apply()`` without enumeration.

Port the XML-parsing enumeration from ``hwpapi/classes/styles.py``
in a later phase when an HWP runtime is guaranteed.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Style", "StyleCollection"]


class Style:
    """Value object for a single paragraph style."""

    __slots__ = ("_app", "name", "index")

    def __init__(self, app: "App", name: str, index: Optional[int] = None) -> None:
        self._app = app
        self.name = name
        self.index = index

    def apply(self) -> bool:
        """Apply this style to the current cursor paragraph."""
        impl = self._app.engine.impl
        try:
            pset = impl.HParameterSet.HStyle
            impl.HAction.GetDefault("Style", pset.HSet)
            if self.index is not None:
                pset.Apply = int(self.index)
            else:
                pset.HSet.SetItem("Name", self.name)
            return bool(impl.HAction.Execute("Style", pset.HSet))
        except Exception:
            return False

    def __repr__(self) -> str:
        if self.index is not None:
            return f"Style({self.name!r}, index={self.index})"
        return f"Style({self.name!r})"


class StyleCollection:
    """
    ``doc.styles`` — paragraph-style collection (minimal Phase 3 impl).

    Enumeration requires a live HWP runtime because HWP COM doesn't
    expose a direct style list. This stub is Protocol-compliant; full
    enumeration will land after Phase 4 when the HWPML-export path
    is ported over from ``hwpapi/classes/styles.py``.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        """Empty until HWPML-based enumeration is ported."""
        return []

    def __iter__(self) -> Iterator[Style]:
        return iter(())

    def __len__(self) -> int:
        return 0

    def __contains__(self, key) -> bool:
        if isinstance(key, Style):
            return False
        if isinstance(key, str):
            return key in self.names()
        return False

    def __getitem__(self, key) -> Style:
        if isinstance(key, str):
            # HWP validates the name on apply(); the wrapper is cheap.
            return Style(self._app, key)
        if isinstance(key, int):
            names = self.names()
            if 0 <= key < len(names):
                return Style(self._app, names[key], index=key)
            raise IndexError(f"Style index {key} out of range")
        raise TypeError(
            f"Style key must be str or int, got {type(key).__name__}"
        )

    def filter(self, predicate: Callable[[Style], bool]) -> List[Style]:
        return [s for s in self if predicate(s)]

    @property
    def current(self) -> Optional[Style]:
        """Paragraph style currently applied at the cursor, or ``None``."""
        impl = self._app.engine.impl
        try:
            pset = impl.HParameterSet.HStyle
            impl.HAction.GetDefault("Style", pset.HSet)
            idx = int(pset.Apply)
            return Style(self._app, f"style_{idx}", index=idx)
        except Exception:
            return None

    def __repr__(self) -> str:
        return "StyleCollection(<enumeration requires HWP runtime>)"
