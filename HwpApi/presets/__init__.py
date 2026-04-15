"""
Document presets — "한 줄로 보기 좋은 결과".

승승아빠 JScript 매크로(ref/승승아빠-한글매크로*.zip) 를 Python 으로 이식한
30 여 개의 문서 꾸미기 프리셋 모음. 계획서: docs/DOCUMENT_PRESETS_PLAN.md

Usage::

    app = App()
    app.preset.striped_rows()                       # 현재 표 줄무늬
    app.preset.table_header(color="sky")            # 제목행 하늘색
    app.preset.title_box(text="보고서 제목")
    app.preset.toc(with_bookmarks=True)

각 preset 은 kwargs 로 override 가능. 자세한 설명은 각 메소드의 docstring.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


class Presets:
    """
    문서 프리셋 accessor. ``app.preset`` 으로 접근.

    v0.0.14 기준: 줄무늬 표 (striped_rows) 만 구현.
    이후 릴리즈에서 계속 확장 (DOCUMENT_PRESETS_PLAN.md 참고).
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    # ═════════════════════════════════════════════════════════════
    # Table presets
    # ═════════════════════════════════════════════════════════════

    def striped_rows(
        self,
        colors: Optional[List[str]] = None,
        header_color: Optional[str] = None,
        skip_header: bool = True,
    ) -> "App":
        """
        현재 커서가 들어있는 표를 **줄무늬** 로 꾸밉니다 (zebra stripe).

        Parameters
        ----------
        colors : list[str], optional
            줄무늬에 쓸 색상 hex 리스트. 기본 ``["#FFFFFF", "#F5F5F5"]``
            (흰색 / 연회색). 3개 이상 지정하면 순환.
        header_color : str, optional
            첫 행(헤더) 전용 색상. ``None`` 이면 skip_header 에 따라 동작.
        skip_header : bool
            True 이면 첫 행은 색상 적용에서 제외 (기본).

        Returns
        -------
        App
            chainable — ``self._app``.

        Examples
        --------
        >>> app.preset.striped_rows()                       # 기본 흰/회
        >>> app.preset.striped_rows(colors=["#E8F4F8", "#FFFFFF"])
        >>> app.preset.striped_rows(header_color="#003366")  # 진파랑 헤더

        Notes
        -----
        구현: 표 내부에서 ``TableCellBlockRow`` + ``CellFill`` 을 행 단위로
        반복. 호출 전에 커서가 표 안에 있어야 합니다.
        """
        app = self._app
        colors = colors or ["#FFFFFF", "#F5F5F5"]

        # 표 안이 아니면 아무것도 하지 않음
        if not app.in_table():
            app.logger.warning("striped_rows: 커서가 표 안에 있지 않습니다.")
            return app

        # 표 맨 위 셀로 이동
        try:
            app.api.Run("TableColBegin")   # 현재 행 맨 앞
            # 맨 위로 올라가기
            while True:
                prev_ok = app.api.Run("TableUpperCell")
                if not prev_ok:
                    break
        except Exception:
            pass

        row_index = 0
        while True:
            # 현재 행 전체 선택
            app.api.Run("TableCellBlock")
            app.api.Run("TableCellBlockRow")

            # 색상 결정
            if row_index == 0 and (header_color or skip_header):
                if header_color:
                    _apply_cell_bg(app, header_color)
                # skip_header=True + header_color=None → 첫 행 그대로 두기
            else:
                color_idx = (row_index - 1 if skip_header else row_index) % len(colors)
                _apply_cell_bg(app, colors[color_idx])

            # 선택 해제 후 다음 행으로
            app.api.Run("Cancel")
            row_index += 1

            # 다음 행으로 이동 — 실패 시 표 끝
            moved = app.api.Run("TableLowerCell")
            if not moved:
                break
            # TableLowerCell 이 열까지 이동시킬 수 있으므로 항상 0열로
            try:
                app.api.Run("TableColBegin")
            except Exception:
                pass

        return app


def _apply_cell_bg(app, hex_color: str) -> None:
    """현재 선택된 셀(들)의 배경색을 hex 색상으로 지정."""
    try:
        from hwpapi.parametersets.properties import convert_to_hwp_color
        color = convert_to_hwp_color(hex_color)
    except Exception:
        # Fallback — parse hex manually (BBGGRR order for HWP)
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        color = b * 0x10000 + g * 0x100 + r

    try:
        # Try modern CellFill pset approach
        act = app.api.CreateAction("CellFill")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("FillColor", color)
        try:
            pset.SetItem("HasFill", 1)
        except Exception:
            pass
        act.Execute(pset)
    except Exception as e:
        app.logger.debug(f"_apply_cell_bg failed: {e}")
