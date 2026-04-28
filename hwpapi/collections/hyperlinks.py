"""
:mod:`hwpapi.collections.hyperlinks` — HyperlinkCollection.

Walks the document's ``HeadCtrl`` chain for controls whose ``CtrlID``
matches HWP's hyperlink identifier (``"%hlk"``). The URL/Text payload
lives on the control's ``Properties`` parameter set.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Hyperlink", "HyperlinkCollection"]


# HWP hyperlink control IDs seen in the wild: "%hlk" (4 chars) or
# "%hyp" depending on version. We accept both.
_HYPERLINK_CTRL_IDS = {"%hlk", "%hyp"}


class Hyperlink:
    """Value object for a single hyperlink."""

    __slots__ = ("text", "url")

    def __init__(self, text: str, url: str):
        self.text = text
        self.url = url

    def __repr__(self) -> str:
        return f"Hyperlink(text={self.text!r}, url={self.url!r})"

    def __eq__(self, other) -> bool:
        if isinstance(other, Hyperlink):
            return self.text == other.text and self.url == other.url
        return NotImplemented

    def __hash__(self) -> int:
        return hash(("Hyperlink", self.text, self.url))


class HyperlinkCollection:
    """
    ``doc.hyperlinks`` — collection of hyperlink controls.

    Iteration yields :class:`Hyperlink` value objects extracted from the
    ``HeadCtrl`` chain. Named subscript matches on link ``text``; ordinal
    subscript returns the nth link in document order.
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

    def _extract(self, ctrl) -> Hyperlink | None:
        try:
            cid = str(getattr(ctrl, "CtrlID", "") or "")
        except Exception:
            return None
        if cid not in _HYPERLINK_CTRL_IDS:
            return None
        text = ""
        url = ""
        props = getattr(ctrl, "Properties", None)
        if props is not None:
            for t_attr in ("Text", "Command"):
                try:
                    v = props.Item(t_attr)
                    if v:
                        text = str(v)
                        break
                except Exception:
                    continue
            for u_attr in ("Command", "URL", "Location"):
                try:
                    v = props.Item(u_attr)
                    if v:
                        url = str(v)
                        break
                except Exception:
                    continue
        if not text:
            text = str(getattr(ctrl, "UserDesc", "") or "")
        return Hyperlink(text, url)

    def _raw(self) -> List[Hyperlink]:
        out: List[Hyperlink] = []
        for ctrl in self._iter_ctrls():
            h = self._extract(ctrl)
            if h is not None:
                out.append(h)
        return out

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        """Link-display-text values, document order."""
        return [h.text for h in self._raw()]

    def __iter__(self) -> Iterator[Hyperlink]:
        return iter(self._raw())

    def __len__(self) -> int:
        return len(self._raw())

    def __contains__(self, key) -> bool:
        if isinstance(key, Hyperlink):
            return key in self._raw()
        if isinstance(key, str):
            return any(h.text == key or h.url == key for h in self._raw())
        return False

    def __getitem__(self, key) -> Hyperlink:
        items = self._raw()
        if isinstance(key, int):
            return items[key]
        if isinstance(key, str):
            for h in items:
                if h.text == key or h.url == key:
                    return h
            raise KeyError(f"Hyperlink {key!r} not found")
        raise TypeError(
            f"Hyperlink key must be str or int, got {type(key).__name__}"
        )

    def filter(
        self, predicate: Callable[[Hyperlink], bool]
    ) -> List[Hyperlink]:
        return [h for h in self if predicate(h)]

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def add(self, text: str, url: str) -> Hyperlink:
        """Insert a hyperlink at the current cursor position."""
        impl = self._app.engine.impl
        try:
            pset = impl.HParameterSet.HHyperLink
            impl.HAction.GetDefault("Hyperlink", pset.HSet)
            pset.Text = str(text)
            pset.Command = str(url)
            impl.HAction.Execute("Hyperlink", pset.HSet)
        except Exception:
            pass
        return Hyperlink(text, url)

    def __repr__(self) -> str:
        try:
            n = len(self)
        except Exception:
            n = "?"
        return f"HyperlinkCollection(count={n})"
