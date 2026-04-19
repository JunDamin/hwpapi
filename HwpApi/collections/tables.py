"""
:mod:`hwpapi.collections.tables` — TableCollection.

Walks the ``HeadCtrl`` → ``Next`` chain for ``CtrlID == "tbl "`` entries.
Named access matches on the table's caption (``UserDesc``); ordinal
access returns the nth table in document order.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Table", "TableCollection"]


_TABLE_CTRL_ID = "tbl "


class Table:
    """Value object for a single table control."""

    __slots__ = ("_app", "_ctrl", "index")

    def __init__(self, app: "App", ctrl, index: int) -> None:
        self._app = app
        self._ctrl = ctrl
        self.index = index

    @property
    def caption(self) -> str:
        """Caption text, or an empty string if the table has none."""
        try:
            return str(getattr(self._ctrl, "UserDesc", "") or "")
        except Exception:
            return ""

    @property
    def name(self) -> str:
        """Best-effort display name — caption when present, else ordinal."""
        cap = self.caption
        return cap if cap else f"table_{self.index}"

    def select(self) -> bool:
        try:
            self._app.engine.impl.SelectCtrl(self._ctrl)
            return True
        except Exception:
            return False

    def __repr__(self) -> str:
        cap = self.caption
        return (
            f"Table(#{self.index}, caption={cap!r})" if cap
            else f"Table(#{self.index})"
        )


class TableCollection:
    """``doc.tables`` — tables in document order."""

    __slots__ = ("_app",)

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Internal iteration
    # ------------------------------------------------------------------

    def _iter_ctrls(self) -> Iterator:
        impl = self._app.engine.impl
        try:
            ctrl = impl.HeadCtrl
        except Exception:
            return
        seen: set[int] = set()
        while ctrl is not None:
            key = id(ctrl)
            if key in seen:
                break
            seen.add(key)
            yield ctrl
            try:
                ctrl = ctrl.Next
            except Exception:
                break

    def _raw(self) -> List[Table]:
        out: List[Table] = []
        idx = 0
        for ctrl in self._iter_ctrls():
            try:
                cid = str(getattr(ctrl, "CtrlID", "") or "")
            except Exception:
                continue
            if cid != _TABLE_CTRL_ID:
                continue
            out.append(Table(self._app, ctrl, idx))
            idx += 1
        return out

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        return [t.name for t in self._raw()]

    def __iter__(self) -> Iterator[Table]:
        return iter(self._raw())

    def __len__(self) -> int:
        return len(self._raw())

    def __contains__(self, key) -> bool:
        if isinstance(key, Table):
            return any(t._ctrl is key._ctrl for t in self._raw())
        if isinstance(key, str):
            return any(t.caption == key for t in self._raw())
        if isinstance(key, int):
            return 0 <= key < len(self._raw())
        return False

    def __getitem__(self, key) -> Optional[Table]:
        items = self._raw()
        if isinstance(key, int):
            return items[key]
        if isinstance(key, str):
            for t in items:
                if t.caption == key:
                    return t
            return None
        raise TypeError(
            f"Table key must be str or int, got {type(key).__name__}"
        )

    def filter(self, predicate: Callable[[Table], bool]) -> List[Table]:
        return [t for t in self if predicate(t)]

    def __repr__(self) -> str:
        try:
            n = len(self)
        except Exception:
            n = "?"
        return f"TableCollection(count={n})"
