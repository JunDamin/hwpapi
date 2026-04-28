"""
:mod:`hwpapi.collections.images` — ImageCollection.

Enumerates ``HeadCtrl`` → ``Next`` for ``CtrlID == "gso "`` (generic
shape object) entries whose ``UserDesc`` indicates a picture. The
``gso`` space also covers non-image shapes, so we filter on description
keywords to narrow to images.

All COM access goes through ``self._app.engine.impl``.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Image", "ImageCollection"]


_IMAGE_CTRL_IDS = {"gso ", "pic "}
_IMAGE_DESC_KEYWORDS = ("그림", "picture", "image", "photo", "사진")


class Image:
    """Value object for a single image control."""

    __slots__ = ("_app", "_ctrl", "index")

    def __init__(self, app: "App", ctrl, index: int) -> None:
        self._app = app
        self._ctrl = ctrl
        self.index = index

    @property
    def name(self) -> str:
        try:
            desc = getattr(self._ctrl, "UserDesc", "") or ""
            if desc:
                return str(desc)
        except Exception:
            pass
        return f"image_{self.index}"

    def select(self) -> bool:
        try:
            self._app.engine.impl.SelectCtrl(self._ctrl)
            return True
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"Image(#{self.index}, {self.name!r})"


class ImageCollection:
    """``doc.images`` — image controls in document order."""

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

    def _is_image(self, ctrl) -> bool:
        try:
            cid = str(getattr(ctrl, "CtrlID", "") or "")
        except Exception:
            return False
        if cid not in _IMAGE_CTRL_IDS:
            return False
        try:
            desc = str(getattr(ctrl, "UserDesc", "") or "").lower()
        except Exception:
            desc = ""
        if not desc:
            # bare gso: treat as image by default — picture is the common case
            return cid.strip() == "gso" or cid.strip() == "pic"
        return any(kw in desc for kw in _IMAGE_DESC_KEYWORDS)

    def _raw(self) -> List[Image]:
        out: List[Image] = []
        idx = 0
        for ctrl in self._iter_ctrls():
            if self._is_image(ctrl):
                out.append(Image(self._app, ctrl, idx))
                idx += 1
        return out

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    def names(self) -> List[str]:
        return [img.name for img in self._raw()]

    def __iter__(self) -> Iterator[Image]:
        return iter(self._raw())

    def __len__(self) -> int:
        return len(self._raw())

    def __contains__(self, key) -> bool:
        if isinstance(key, Image):
            return any(img._ctrl is key._ctrl for img in self._raw())
        if isinstance(key, str):
            return key in self.names()
        if isinstance(key, int):
            return 0 <= key < len(self._raw())
        return False

    def __getitem__(self, key) -> Image:
        items = self._raw()
        if isinstance(key, int):
            return items[key]
        if isinstance(key, str):
            for img in items:
                if img.name == key:
                    return img
            raise KeyError(f"Image {key!r} not found")
        raise TypeError(
            f"Image key must be str or int, got {type(key).__name__}"
        )

    def filter(self, predicate: Callable[[Image], bool]) -> List[Image]:
        return [img for img in self if predicate(img)]

    def __repr__(self) -> str:
        try:
            n = len(self)
        except Exception:
            n = "?"
        return f"ImageCollection(count={n})"
