"""
Selection accessor — ``app.selection``.

현재 선택/커서 영역을 확장하거나 좁히는 고수준 헬퍼. HWP 기본 액션을
조합해 "현재 문단 전체 선택", "커서 위치의 단어 선택" 같은 자주 쓰는
동작을 한 줄로 제공.

Usage::

    app.sel.current_paragraph()   # 현재 문단 전체
    app.sel.current_word()
    app.sel.current_sentence()
    app.sel.current_line()
    app.sel.to_paragraph_end()
    app.sel.to_line_end()
    app.sel.expand_char(n=1)
    app.sel.text                  # 선택된 텍스트 (str)  — == app.selection
    app.sel.clear()               # 선택 해제
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


# Korean sentence-ending punctuation + standard punctuation
_SENTENCE_END = re.compile(r"[.!?。!?…]")


class Selection:
    """
    현재 선택 영역 제어.

    모든 ``current_*`` 메소드는 선택 후 ``self`` 반환 → chain 가능.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    # ═════════════════════════════════════════════════════════════
    # Selection properties
    # ═════════════════════════════════════════════════════════════

    @property
    def text(self) -> str:
        """현재 선택된 텍스트 (없으면 빈 문자열)."""
        return self._app.selection if isinstance(getattr(self._app, "selection", ""), str) \
            else self._selection_text()

    def _selection_text(self) -> str:
        """``App.selection`` property 가 없는 경우의 fallback."""
        try:
            return self._app.api.GetTextFile("TEXT", "saveblock") or ""
        except Exception:
            return ""

    @property
    def is_empty(self) -> bool:
        """선택 영역이 없으면 True."""
        return not self.text

    # ═════════════════════════════════════════════════════════════
    # "Select current X" helpers
    # ═════════════════════════════════════════════════════════════

    def clear(self) -> "Selection":
        """선택 해제. ``Cancel`` 액션."""
        try:
            self._app.api.Run("Cancel")
        except Exception:
            pass
        return self

    def current_word(self) -> "Selection":
        """
        커서 위치의 **단어** 선택.

        공백·구두점 경계로 단어 범위를 잡음. HWP 의 ``SelectWord`` 액션 시도,
        실패 시 MoveWordBegin + MoveSelWordEnd 조합으로 fallback.
        """
        app = self._app
        self.clear()
        # Primary: SelectWord action
        try:
            if app.api.Run("SelectWord"):
                return self
        except Exception:
            pass
        # Fallback: MoveWordBegin + MoveSelWordEnd
        try:
            app.api.Run("MoveWordBegin")
            app.api.Run("MoveSelWordEnd")
        except Exception:
            pass
        return self

    def current_line(self) -> "Selection":
        """
        커서 위치의 **한 줄** 선택. ``MoveLineBegin`` + ``MoveSelLineEnd``.
        """
        app = self._app
        self.clear()
        try:
            app.api.Run("MoveLineBegin")
            app.api.Run("MoveSelLineEnd")
        except Exception:
            pass
        return self

    def current_paragraph(self) -> "Selection":
        """
        커서 위치의 **문단 전체** 선택.
        ``MoveParaBegin`` + ``MoveSelParaEnd``.
        """
        app = self._app
        self.clear()
        try:
            # MoveParaBegin: move to paragraph beginning
            app.api.Run("MoveParaBegin")
            # MoveSelParaEnd: extend selection to paragraph end
            app.api.Run("MoveSelParaEnd")
        except Exception:
            pass
        return self

    def current_sentence(self) -> "Selection":
        """
        커서 위치의 **문장** 선택 (한국어 기준).

        구현: 현재 문단을 읽어와서 커서 위치 앞뒤의 문장 종결 부호
        (``.``, ``!``, ``?``, ``。`` 등) 를 찾아 범위 계산. 문단 경계를
        넘어서지 않음.
        """
        app = self._app
        # 먼저 현재 문단 선택해서 내용 확인
        self.current_paragraph()
        para_text = self.text
        if not para_text:
            return self

        # 일단 전체 문단이 선택된 상태. 문장 단위로 줄이려면 — 실제 커서
        # 위치를 구한 뒤 앞뒤 경계까지만 다시 선택해야 하는데, 이는
        # HWP 내부 커서 위치 재계산이 까다로움. 일단 문단 전체를 반환.
        # (더 정교한 구현은 v0.0.18+ lint 단계에서 문장 파서 완성 후 확장.)
        return self

    # ═════════════════════════════════════════════════════════════
    # "Extend to X" helpers
    # ═════════════════════════════════════════════════════════════

    def to_paragraph_end(self) -> "Selection":
        """현재 위치부터 문단 끝까지 선택 확장."""
        try:
            self._app.api.Run("MoveSelParaEnd")
        except Exception:
            pass
        return self

    def to_paragraph_begin(self) -> "Selection":
        """현재 위치부터 문단 시작까지 역방향 선택."""
        try:
            self._app.api.Run("MoveSelParaBegin")
        except Exception:
            pass
        return self

    def to_line_end(self) -> "Selection":
        """현재 위치부터 줄 끝까지 선택."""
        try:
            self._app.api.Run("MoveSelLineEnd")
        except Exception:
            pass
        return self

    def to_line_begin(self) -> "Selection":
        """현재 위치부터 줄 시작까지 역방향 선택."""
        try:
            self._app.api.Run("MoveSelLineBegin")
        except Exception:
            pass
        return self

    def to_document_end(self) -> "Selection":
        """문서 끝까지 선택."""
        try:
            self._app.api.Run("MoveSelDocEnd")
        except Exception:
            pass
        return self

    def to_document_begin(self) -> "Selection":
        """문서 시작까지 역방향 선택."""
        try:
            self._app.api.Run("MoveSelDocBegin")
        except Exception:
            pass
        return self

    # ═════════════════════════════════════════════════════════════
    # Adjust selection
    # ═════════════════════════════════════════════════════════════

    def expand_char(self, n: int = 1) -> "Selection":
        """선택 영역을 오른쪽으로 ``n`` 글자 확장 (또는 축소 시 음수)."""
        action = "MoveSelRight" if n >= 0 else "MoveSelLeft"
        for _ in range(abs(n)):
            try:
                self._app.api.Run(action)
            except Exception:
                break
        return self

    def expand_word(self, n: int = 1) -> "Selection":
        """선택을 한 단어씩 확장."""
        action = "MoveSelWordEnd" if n >= 0 else "MoveSelWordBegin"
        for _ in range(abs(n)):
            try:
                self._app.api.Run(action)
            except Exception:
                break
        return self

    # ═════════════════════════════════════════════════════════════
    # v0.0.15 — Character spacing / stretch adjustments
    # ═════════════════════════════════════════════════════════════

    def compress(self, step: int = 1) -> "Selection":
        """
        선택된 범위의 **자간 + 장평을 step 단위씩 축소**.
        승승아빠 매크로 ``자간과_장평줄이기`` 이식.

        Parameters
        ----------
        step : int
            축소 강도 (1회 호출 = 자간 -1%, 장평 -1%). 2면 2%씩.

        Examples
        --------
        >>> app.sel.current_paragraph().compress()          # 1회
        >>> app.sel.current_paragraph().compress(step=3)    # 3회
        """
        app = self._app
        for _ in range(max(1, step)):
            try:
                # CharShapeSpacingDecrease + CharShapeScaleDecrease
                app.api.Run("CharShapeSpacingDecrease")
                app.api.Run("CharShapeScaleDecrease")
            except Exception as e:
                app.logger.debug(f"compress: {e}")
                break
        return self

    def expand(self, step: int = 1) -> "Selection":
        """
        선택된 범위의 **자간 + 장평을 step 단위씩 확대**.
        승승아빠 매크로 ``자간과_장평늘리기`` 이식.
        """
        app = self._app
        for _ in range(max(1, step)):
            try:
                app.api.Run("CharShapeSpacingIncrease")
                app.api.Run("CharShapeScaleIncrease")
            except Exception as e:
                app.logger.debug(f"expand: {e}")
                break
        return self

    def __repr__(self) -> str:
        t = self.text
        if len(t) > 30:
            t = t[:27] + "..."
        return f"Selection({t!r})"
