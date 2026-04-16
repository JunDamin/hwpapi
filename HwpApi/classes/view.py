"""
View accessor — ``app.view``.

뷰포트 / 창 / 확대율 제어.

Usage::

    app.view.zoom(150)              # 확대 150%
    app.view.zoom_fit_page()        # 페이지 맞춤
    app.view.zoom_fit_width()       # 너비 맞춤
    app.view.zoom_actual()          # 100%
    app.view.scroll_to_cursor()
    app.view.full_screen()
    app.view.exit_full_screen()
    app.view.page_mode()            # 편집 화면 ↔ 미리보기
    app.view.draft_mode()
    app.view.toggle_rulers()
    app.view.zoom_current            # property (float) — 현재 확대율
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from hwpapi.core.app import App


class View:
    """뷰포트 제어 accessor."""

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    # ═════════════════════════════════════════════════════════════
    # Zoom
    # ═════════════════════════════════════════════════════════════

    @property
    def zoom_current(self) -> float:
        """현재 확대율 (%) — read-only."""
        try:
            return float(self._app.api.XHwpDocuments.Active_XHwpDocument
                                 .XHwpDocumentInfo.ZoomRate)
        except Exception:
            return 100.0

    def zoom(self, percent: int) -> "App":
        """
        확대율을 지정 퍼센트로 설정.

        Parameters
        ----------
        percent : int
            대략 10 ~ 500 범위. 10 미만은 10 으로, 500 초과는 500 으로 클램프.

        Examples
        --------
        >>> app.view.zoom(150)
        >>> app.view.zoom(75)
        """
        app = self._app
        percent = max(10, min(500, int(percent)))

        # v0.0.24+: 이전엔 "PictureScale" 액션 (그림 크기 조절용 — 잘못된
        # 액션) 을 호출했음. 실제 zoom 은 ZoomRate property 직접 설정.
        try:
            doc_info = app.api.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo
            doc_info.ZoomRate = percent
            app.logger.info(f"view.zoom: {percent}%")
        except Exception as e:
            app.logger.debug(
                f"view.zoom XHwpDocumentInfo.ZoomRate: {type(e).__name__}: {e}",
                exc_info=True,
            )
            # Fallback — try HAction "ScreenZoom" (HWP 일부 버전)
            try:
                hpset = app.api.HParameterSet.HZoom
                app.api.HAction.GetDefault("ScreenZoom", hpset.HSet)
                hpset.HSet.SetItem("ZoomFactor", percent)
                app.api.HAction.Execute("ScreenZoom", hpset.HSet)
                app.logger.info(f"view.zoom (ScreenZoom fallback): {percent}%")
            except Exception as e2:
                app.logger.warning(
                    f"view.zoom failed: {type(e2).__name__}: {e2}",
                    exc_info=True,
                )
        return app

    def zoom_fit_page(self) -> "App":
        """페이지 전체가 화면에 맞도록 확대율 조정."""
        try:
            self._app.api.Run("ViewZoomFitPage")
        except Exception:
            try:
                self._app.api.Run("MoveViewFitPage")
            except Exception as e:
                self._app.logger.debug(f"zoom_fit_page: {e}")
        return self._app

    def zoom_fit_width(self) -> "App":
        """페이지 너비 맞춤."""
        try:
            self._app.api.Run("ViewZoomFitWidth")
        except Exception as e:
            self._app.logger.debug(f"zoom_fit_width: {e}")
        return self._app

    def zoom_actual(self) -> "App":
        """원본 크기 (100%)."""
        return self.zoom(100)

    # ═════════════════════════════════════════════════════════════
    # Scroll / position
    # ═════════════════════════════════════════════════════════════

    def scroll_to_cursor(self) -> "App":
        """현재 커서 위치가 화면에 보이도록 스크롤."""
        try:
            self._app.api.Run("MoveScrollCurPos")
        except Exception:
            try:
                # Fallback — tiny cursor nudge
                self._app.api.Run("MoveRight")
                self._app.api.Run("MoveLeft")
            except Exception as e:
                self._app.logger.debug(f"scroll_to_cursor: {e}")
        return self._app

    # ═════════════════════════════════════════════════════════════
    # Window modes
    # ═════════════════════════════════════════════════════════════

    def full_screen(self) -> "App":
        """전체 화면 모드."""
        try:
            self._app.api.Run("FullScreen")
        except Exception as e:
            self._app.logger.debug(f"full_screen: {e}")
        return self._app

    def exit_full_screen(self) -> "App":
        """전체 화면 해제 (토글)."""
        return self.full_screen()    # HWP 는 토글 방식

    def page_mode(self) -> "App":
        """편집 화면 모드 (기본)."""
        try:
            self._app.api.Run("ViewPage")
        except Exception as e:
            self._app.logger.debug(f"page_mode: {e}")
        return self._app

    def draft_mode(self) -> "App":
        """초안 모드 (빠른 렌더링)."""
        try:
            self._app.api.Run("ViewDraft")
        except Exception as e:
            self._app.logger.debug(f"draft_mode: {e}")
        return self._app

    def toggle_rulers(self) -> "App":
        """가로/세로 자 표시 토글."""
        try:
            self._app.api.Run("ViewRuler")
        except Exception as e:
            self._app.logger.debug(f"toggle_rulers: {e}")
        return self._app

    def __repr__(self) -> str:
        try:
            z = self.zoom_current
            return f"View(zoom={z:.0f}%)"
        except Exception:
            return "View(<app>)"
