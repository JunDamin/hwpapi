"""
Images collection accessor — ``app.images``.

문서 내 모든 이미지(그림 control) 를 list-like 로 다루고, 일괄 작업을 제공.
승승아빠 매크로의 ``모든_이미지_크기조절`` / ``이미지_회색처리_선택된거만`` 대응.

Usage::

    app.images                          # 이미지 컬렉션
    app.images.resize_all(width="100mm")
    app.images.grayscale_all()
    for img in app.images:
        img.select()                    # 커서를 이미지 위치로
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


class Image:
    """단일 이미지 control 값 객체."""

    __slots__ = ("_app", "_ctrl", "index")

    def __init__(self, app: "App", ctrl, index: int):
        self._app = app
        self._ctrl = ctrl
        self.index = index

    @property
    def name(self) -> str:
        try:
            return str(self._ctrl.UserDesc or self._ctrl.CtrlID or f"image_{self.index}")
        except Exception:
            return f"image_{self.index}"

    def select(self) -> bool:
        """커서를 이 이미지 위치로 이동."""
        try:
            # Use AnchorPosCtrl-like approach — naive fallback via SetPosBySet
            from hwpapi.classes.controls import Control
            c = Control(self._app, self._ctrl)
            return c.select()
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"Image(#{self.index}, {self.name!r})"


class Images:
    """
    문서 내 이미지(그림) 컬렉션 — ``app.images``.

    HeadCtrl → Next linked list 를 순회해서 ``CtrlID`` 가 이미지 계열인
    control 만 필터링.

    이미지 계열 CtrlID (HWP 상수):
      - "gso " (general shape object — 그림, 도형)
      - "pic " (picture — 이전 버전 호환)
    """

    __slots__ = ("_app",)

    IMAGE_CTRL_IDS = {"gso ", "pic "}

    def __init__(self, app: "App"):
        self._app = app

    def _iter_controls(self) -> Iterator:
        """HeadCtrl → Next 순회."""
        try:
            ctrl = self._app.api.HeadCtrl
        except Exception:
            return
        while ctrl is not None:
            yield ctrl
            try:
                ctrl = ctrl.Next
            except Exception:
                break

    def _is_image(self, ctrl) -> bool:
        """CtrlID 가 이미지 계열인지 판별."""
        try:
            cid = str(ctrl.CtrlID or "")
        except Exception:
            return False
        # gso 는 도형·그림·OLE 등 포함 — UserDesc 로 그림만 필터 시도
        if cid in self.IMAGE_CTRL_IDS:
            try:
                desc = str(ctrl.UserDesc or "").lower()
                # "그림", "picture", "image" 포함이면 이미지
                return any(kw in desc for kw in ("그림", "picture", "image"))
            except Exception:
                # UserDesc 없으면 gso 는 일단 이미지로 간주
                return cid == "gso "
        return False

    def __iter__(self) -> Iterator[Image]:
        idx = 0
        for ctrl in self._iter_controls():
            if self._is_image(ctrl):
                yield Image(self._app, ctrl, idx)
                idx += 1

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __getitem__(self, i: int) -> Image:
        if isinstance(i, slice):
            return list(self)[i]
        items = list(self)
        return items[i]

    def __repr__(self) -> str:
        n = len(self)
        return f"Images(count={n})"

    # ═════════════════════════════════════════════════════════════
    # Batch operations
    # ═════════════════════════════════════════════════════════════

    def resize_all(
        self,
        width: str = "100mm",
        keep_ratio: bool = True,
    ) -> "App":
        """
        문서 내 모든 이미지의 **너비** 를 지정값으로 통일.

        Parameters
        ----------
        width : str
            대상 너비 (예: ``"100mm"``, ``"15cm"``, ``"4in"``). 또는 HWPUNIT 숫자.
        keep_ratio : bool
            True (기본) 이면 가로세로 비율 유지. False 이면 높이는 변경 없음.

        Returns
        -------
        App

        Examples
        --------
        >>> app.images.resize_all(width="100mm")
        >>> app.images.resize_all(width="15cm", keep_ratio=True)

        Notes
        -----
        구현: 각 이미지 control 을 선택한 뒤 ``ShapeObjectSize`` 액션의 pset
        에 Width 주입. 이미지가 표 안에 있어도 작동.
        """
        from hwpapi.functions import to_hwpunit

        # Parse width
        if isinstance(width, str):
            target_w = to_hwpunit(width)
        else:
            target_w = int(width)

        app = self._app
        count = 0
        for img in self:
            try:
                if not img.select():
                    continue
                # 원본 크기 읽기 → 비율 계산
                ctrl = img._ctrl
                cur_w = getattr(ctrl, "Width", None) or 1
                cur_h = getattr(ctrl, "Height", None) or 1

                new_h = int(cur_h * target_w / cur_w) if keep_ratio else cur_h

                # Apply via ShapeObjectSize action (if exists) or ShapeObject pset
                try:
                    act = app.api.CreateAction("ShapeObjectSize")
                    pset = act.CreateSet()
                    act.GetDefault(pset)
                    pset.SetItem("Width", target_w)
                    pset.SetItem("Height", new_h)
                    act.Execute(pset)
                except Exception:
                    # Fallback — ShapeObjectDialog
                    try:
                        ctrl.Width = target_w
                        ctrl.Height = new_h
                    except Exception:
                        pass
                count += 1
                # Deselect
                try:
                    app.api.Run("Cancel")
                except Exception:
                    pass
            except Exception as e:
                app.logger.debug(f"resize_all: image {img.index} failed: {e}")

        app.logger.info(f"resize_all: {count} images resized")
        return app

    def grayscale_all(self) -> "App":
        """
        모든 이미지를 **회색조** 로 변환 (효과 적용). 승승아빠 매크로의
        ``이미지_회색처리`` 전체 버전.

        Examples
        --------
        >>> app.images.grayscale_all()   # 흑백 인쇄 준비
        """
        app = self._app
        count = 0
        for img in self:
            try:
                if not img.select():
                    continue
                # PictureEffect1 / PictureEffectGrayScale 중 환경에 존재하는
                # 액션을 시도
                for action_name in ("PictureEffectGrayScale",
                                    "PictureEffect1",
                                    "ImageEffectGrayscale"):
                    try:
                        app.api.Run(action_name)
                        count += 1
                        break
                    except Exception:
                        continue
                try:
                    app.api.Run("Cancel")
                except Exception:
                    pass
            except Exception as e:
                app.logger.debug(f"grayscale_all: image {img.index} failed: {e}")
        app.logger.info(f"grayscale_all: {count} images processed")
        return app
