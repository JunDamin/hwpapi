"""
:mod:`hwpapi.collections.tables` — TableCollection.

Walks the ``HeadCtrl`` → ``Next`` chain for ``CtrlID == "tbl "`` entries.
Named access matches on the table's caption (``UserDesc``); ordinal
access returns the nth table in document order.

Phase 4 enriches :class:`Table` with ``.rows``, ``.cols`` and
``.cell(row, col)`` returning a :class:`Cell` value object. COM
interaction routes through the table control's ParameterSet and the
low-level engine.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Cell", "Table", "TableCollection"]


_TABLE_CTRL_ID = "tbl "


def _col_letter(col: int) -> str:
    """0-indexed column → Excel column letters (A, B, …, Z, AA, AB, …)."""
    if col < 0:
        return "?"
    letters = ""
    n = col
    while True:
        letters = chr(ord("A") + (n % 26)) + letters
        n = n // 26 - 1
        if n < 0:
            break
    return letters


class Cell:
    """
    Value object for a single table cell.

    Phase 4 surface: ``.text``, ``.address`` (A1-style), ``.select()``.
    Does not own a COM object — addresses are passed to the HWP engine
    on demand.
    """

    __slots__ = ("_app", "_table", "row", "col")

    def __init__(self, app: "App", table: "Table", row: int, col: int) -> None:
        self._app = app
        self._table = table
        self.row = int(row)
        self.col = int(col)

    @property
    def address(self) -> str:
        """A1-style address (e.g. ``"B3"``)."""
        return f"{_col_letter(self.col)}{self.row + 1}"

    @property
    def text(self) -> str:
        """
        Best-effort cell text. Selects the cell, reads the saved-block
        text through HWP's ``GetTextFile`` API, then cancels the
        selection. Returns ``""`` on failure.
        """
        if not self.select():
            return ""
        impl = self._app.engine.impl
        text = ""
        try:
            impl.Run("TableCellBlock")
            raw = impl.GetTextFile("TEXT", "saveblock")
            text = str(raw or "")
        except Exception:
            text = ""
        finally:
            try:
                impl.Run("Cancel")
            except Exception:
                pass
        return text.rstrip("\r\n")

    def select(self) -> bool:
        """
        Navigate the caret to this cell. Selects the owning table first,
        then uses ``SetCellAddr`` / ``TableCellBlock`` navigation. Returns
        ``True`` on success, ``False`` on any COM error.
        """
        # Select the parent table control so subsequent navigation
        # lands inside it.
        if not self._table.select():
            return False
        impl = self._app.engine.impl
        try:
            # Prefer SetCellAddr when the runtime exposes it.
            setter = getattr(impl, "SetCellAddr", None)
            if callable(setter):
                setter(self.address)
                return True
        except Exception:
            pass
        # Fallback: walk using Table*Cell navigation actions.
        try:
            impl.Run("TableColBegin")
            for _ in range(self.row):
                impl.Run("TableLowerCell")
            for _ in range(self.col):
                impl.Run("TableRightCell")
            return True
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"Cell({self.address}, table=#{self._table.index})"


class Table:
    """Value object for a single table control."""

    __slots__ = ("_app", "_ctrl", "index")

    def __init__(self, app: "App", ctrl, index: int) -> None:
        self._app = app
        self._ctrl = ctrl
        self.index = index

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Shape (rows/cols)
    # ------------------------------------------------------------------

    def _table_pset(self):
        """Fetch the table control's ParameterSet (``Properties``)."""
        ctrl = self._ctrl
        for attr in ("Properties", "Param"):
            try:
                pset = getattr(ctrl, attr, None)
                if pset is not None:
                    return pset
            except Exception:
                continue
        return None

    def _pset_int(self, names: tuple) -> int:
        """Probe a table pset for the first attribute in ``names``."""
        pset = self._table_pset()
        if pset is None:
            return 0
        # Native attribute access first — mocks and HParameterSet both
        # expose attributes directly.
        for n in names:
            try:
                v = getattr(pset, n, None)
                if callable(v):
                    v = v()
                if v is not None:
                    return int(v)
            except Exception:
                continue
        # Fall back to Item(name) — HParameterSet flavour.
        item = getattr(pset, "Item", None)
        if callable(item):
            for n in names:
                try:
                    v = item(n)
                    if v is not None:
                        return int(v)
                except Exception:
                    continue
        return 0

    @property
    def rows(self) -> int:
        """Number of rows in the table (0 if unknown)."""
        return self._pset_int(("Rows", "RowCount"))

    @property
    def cols(self) -> int:
        """Number of columns in the table (0 if unknown)."""
        return self._pset_int(("Cols", "ColCount", "Columns"))

    # ------------------------------------------------------------------
    # Cell access
    # ------------------------------------------------------------------

    def cell(self, row: int, col: int) -> Cell:
        """Return a :class:`Cell` for ``(row, col)`` (0-indexed)."""
        return Cell(self._app, self, row, col)

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
