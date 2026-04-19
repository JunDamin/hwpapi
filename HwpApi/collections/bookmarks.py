"""
:mod:`hwpapi.collections.bookmarks` — BookmarkCollection (책갈피).

HWP exposes bookmarks as a ``bmkt`` linked-list hanging off ``HeadCtrl``.
Each control's ``CtrlCh`` carries the bookmark name. Operations use the
engine's ``ExistBookMark`` / ``SelectBookMark`` / ``DeleteBookMark`` COM
methods and ``Run("InsertBookMark")`` via parameter set.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Bookmark", "BookmarkCollection"]


_BOOKMARK_CTRL_ID = "bokm"


class Bookmark:
    """Value object for a single HWP bookmark."""

    __slots__ = ("_app", "name")

    def __init__(self, app: "App", name: str):
        self._app = app
        self.name = name

    def goto(self) -> bool:
        """Move the cursor to this bookmark. Returns success."""
        try:
            return bool(self._app.engine.impl.SelectBookMark(self.name))
        except Exception:
            return False

    def remove(self) -> bool:
        """Delete this bookmark. Returns success."""
        try:
            return bool(self._app.engine.impl.DeleteBookMark(self.name))
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"Bookmark({self.name!r})"

    def __eq__(self, other) -> bool:
        if isinstance(other, Bookmark):
            return self.name == other.name
        if isinstance(other, str):
            return self.name == other
        return NotImplemented

    def __hash__(self) -> int:
        return hash(("Bookmark", self.name))


class BookmarkCollection:
    """
    ``doc.bookmarks`` — collection of HWP bookmarks (책갈피).

    Enumerates the document's ``HeadCtrl`` → ``Next`` chain, yielding a
    :class:`Bookmark` for each control whose ``CtrlID`` is ``"bokm"``.
    """

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

    def _raw_names(self) -> List[str]:
        out: List[str] = []
        for ctrl in self._iter_ctrls():
            try:
                cid = str(getattr(ctrl, "CtrlID", "") or "")
            except Exception:
                continue
            if cid.strip() != _BOOKMARK_CTRL_ID.strip():
                continue
            # HWP bookmark control carries the name on CtrlCh or UserDesc.
            name = None
            for attr in ("CtrlCh", "UserDesc"):
                try:
                    val = getattr(ctrl, attr, None)
                except Exception:
                    val = None
                if val:
                    name = str(val)
                    break
            if name:
                out.append(name)
        return out

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        return self._raw_names()

    def __iter__(self) -> Iterator[Bookmark]:
        for n in self._raw_names():
            yield Bookmark(self._app, n)

    def __len__(self) -> int:
        return len(self._raw_names())

    def __contains__(self, key) -> bool:
        if not isinstance(key, str):
            return False
        try:
            return bool(self._app.engine.impl.ExistBookMark(key))
        except Exception:
            return key in self._raw_names()

    def __getitem__(self, key) -> Bookmark:
        names = self._raw_names()
        if isinstance(key, int):
            return Bookmark(self._app, names[key])
        if isinstance(key, str):
            if key not in names and not self._exists(key):
                raise KeyError(f"Bookmark {key!r} not found")
            return Bookmark(self._app, key)
        raise TypeError(
            f"Bookmark key must be str or int, got {type(key).__name__}"
        )

    def _exists(self, name: str) -> bool:
        try:
            return bool(self._app.engine.impl.ExistBookMark(name))
        except Exception:
            return False

    def filter(self, predicate: Callable[[Bookmark], bool]) -> List[Bookmark]:
        return [b for b in self if predicate(b)]

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def add(self, name: str) -> bool:
        """Insert a bookmark at the current cursor position."""
        impl = self._app.engine.impl
        try:
            pset = impl.HParameterSet.HBookMark
            impl.HAction.GetDefault("InsertBookMark", pset.HSet)
            pset.Bookmark = str(name)
            return bool(impl.HAction.Execute("InsertBookMark", pset.HSet))
        except Exception:
            return False

    def remove(self, name: str) -> bool:
        try:
            return bool(self._app.engine.impl.DeleteBookMark(name))
        except Exception:
            return False

    def goto(self, name: str) -> bool:
        try:
            return bool(self._app.engine.impl.SelectBookMark(name))
        except Exception:
            return False

    def __repr__(self) -> str:
        try:
            n = len(self)
        except Exception:
            n = "?"
        return f"BookmarkCollection(count={n})"
