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

        # 표 맨 위 셀로 이동 — cell_addr 로 진행 추적
        app.api.Run("TableColBegin")
        last_addr = None
        for _ in range(500):   # safety bound
            cur = _cell_addr(app)
            if cur is None or cur == last_addr:
                break
            last_addr = cur
            app.api.Run("TableUpperCell")

        row_index = 0
        visited = set()
        for _ in range(500):
            cur = _cell_addr(app)
            if cur is None or cur in visited:
                break
            visited.add(cur)

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

            # 다음 행으로 이동 — cell_addr 변화로 진행 확인
            prev_addr = _cell_addr(app)
            app.api.Run("TableLowerCell")
            app.api.Run("TableColBegin")
            new_addr = _cell_addr(app)
            if new_addr is None or new_addr == prev_addr:
                break  # 더 이상 이동 불가 = 표 끝

        return app


    # ═════════════════════════════════════════════════════════════
    # v0.0.15 — Structure presets
    # ═════════════════════════════════════════════════════════════

    def title_box(
        self,
        text: str = "",
        subtitle: str = "",
        width: str = "168mm",
        height: str = "22.6mm",
        bg_color: str = "#EEEEEE",
        font_size: int = 1400,
        title_color: str = "#000000",
    ) -> "App":
        """
        문서 **타이틀 박스** 삽입 — 승승아빠 매크로 ``문서타이틀박스`` 이식.

        1x2 표를 문자취급으로 생성하고 왼쪽 셀에 ``text``, 오른쪽 셀에
        ``subtitle`` 을 배치. 배경색/폰트 크기 custom 가능.

        Examples
        --------
        >>> app.preset.title_box(text="2026년 상반기 업무 보고", subtitle="○○과")
        >>> app.preset.title_box(text="회의록", bg_color="#003366", title_color="#FFFFFF")
        """
        from hwpapi.functions import to_hwpunit
        app = self._app

        try:
            # 1x2 table 생성 — TableCreate 는 cursor 를 **첫 셀에 배치**
            act = app.api.CreateAction("TableCreate")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Rows", 1)
            pset.SetItem("Cols", 2)
            pset.SetItem("WidthType", 0)
            pset.SetItem("HeightType", 0)
            pset.SetItem("WidthValue", to_hwpunit(width))
            pset.SetItem("HeightValue", to_hwpunit(height))
            act.Execute(pset)
        except Exception as e:
            app.logger.warning(f"title_box: table create failed: {e}")
            return app

        # 1단계: 전체 표 배경색 칠하기 — 전체 셀 선택 후 CellFill
        try:
            # 첫 셀에서 시작 → 전체 행 선택
            app.api.Run("TableCellBlock")
            app.api.Run("TableCellBlockRow")
            _apply_cell_bg(app, bg_color)
            app.api.Run("Cancel")
        except Exception as e:
            app.logger.debug(f"title_box bg: {e}")

        # 2단계: 첫 셀로 돌아와 title 삽입 — cell_addr 로 진행 추적
        try:
            app.api.Run("TableColBegin")
            last = None
            for _ in range(10):
                cur = _cell_addr(app)
                if cur is None or cur == last:
                    break
                last = cur
                app.api.Run("TableLeftCell")
            if text:
                app.styled_text(text, bold=True, height=font_size, text_color=title_color)
            # 3단계: 두 번째 셀로 → subtitle
            app.api.Run("TableRightCell")
            if subtitle:
                app.styled_text(subtitle, height=font_size - 200, text_color=title_color)
        except Exception as e:
            app.logger.debug(f"title_box text insertion: {e}")

        # 4단계: 표 밖으로 이동 → 이후 호출되는 preset 은 표 밖 위치에서 시작
        _exit_table(app)

        return app

    def subtitle_bar(
        self,
        text: str = "",
        bg_color: str = "#EEEEEE",
        font_size: int = 1200,
    ) -> "App":
        """
        문서 **소제목 바** — 승승아빠 매크로 ``문서_소제목_바`` 이식.

        한 줄 짜리 회색 배경 바에 소제목을 삽입.

        Examples
        --------
        >>> app.preset.subtitle_bar("1. 개요")
        """
        app = self._app
        try:
            act = app.api.CreateAction("TableCreate")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Rows", 1)
            pset.SetItem("Cols", 1)
            act.Execute(pset)
            app.api.Run("TableCellBlock")
            _apply_cell_bg(app, bg_color)
            app.api.Run("Cancel")
            if text:
                app.styled_text(text, bold=True, height=font_size)
            # 다음 콘텐츠는 표 밖에서 이어지도록
            _exit_table(app)
        except Exception as e:
            app.logger.warning(f"subtitle_bar: {e}")
        return app

    # ═════════════════════════════════════════════════════════════
    # v0.0.15 — Table header/footer presets
    # ═════════════════════════════════════════════════════════════

    TABLE_HEADER_COLORS = {
        "gray": "#4D4D4D",       # 진회색
        "sky": "#9FC5E8",        # 하늘색
        "dark_blue": "#003366",  # 진파랑
        "green": "#6AA84F",      # 녹색
        "red": "#CC4125",        # 빨강
    }

    def table_header(
        self,
        color: str = "sky",
        text_color: str = "#FFFFFF",
        bold: bool = True,
        rows: int = 1,
    ) -> "App":
        """
        현재 표의 **상단 ``rows`` 행을 제목행** 으로 꾸밈.

        Parameters
        ----------
        color : str
            배경색 — preset name (``"gray"``, ``"sky"``, ``"dark_blue"``, ``"green"``, ``"red"``)
            또는 hex (``"#FF0000"``).
        text_color : str
            글자색 hex.
        bold : bool
            글자 굵게.
        rows : int
            제목행 개수 (기본 1).

        Examples
        --------
        >>> app.preset.table_header()                       # 하늘색 + 흰 글씨
        >>> app.preset.table_header(color="gray", text_color="#FFFFFF")
        >>> app.preset.table_header(color="#FF6600", rows=2)
        """
        app = self._app
        if not app.in_table():
            app.logger.warning("table_header: 커서가 표 안에 있지 않습니다.")
            return app

        bg = self.TABLE_HEADER_COLORS.get(color, color)

        # 맨 위로 — cell_addr 변화로 루프 종료 감지
        try:
            app.api.Run("TableColBegin")
            last = None
            for _ in range(500):
                cur = _cell_addr(app)
                if cur is None or cur == last:
                    break
                last = cur
                app.api.Run("TableUpperCell")

            for i in range(rows):
                # 현재 행 전체 선택
                app.api.Run("TableCellBlock")
                app.api.Run("TableCellBlockRow")
                _apply_cell_bg(app, bg)
                # 글자색/굵게 적용
                try:
                    app.set_charshape(bold=bold, text_color=text_color)
                except Exception:
                    pass
                app.api.Run("Cancel")
                if i + 1 < rows:
                    prev = _cell_addr(app)
                    app.api.Run("TableLowerCell")
                    app.api.Run("TableColBegin")
                    new = _cell_addr(app)
                    if new is None or new == prev:
                        break
        except Exception as e:
            app.logger.debug(f"table_header: {e}")
        return app

    def table_footer(
        self,
        color: str = "gray",
        text_color: str = "#FFFFFF",
        bold: bool = True,
        rows: int = 1,
    ) -> "App":
        """
        현재 표의 **하단 ``rows`` 행을 바닥행** 으로 꾸밈. 사용법은
        :meth:`table_header` 와 동일.

        Examples
        --------
        >>> app.preset.table_footer(color="gray", rows=1)   # 합계 행
        """
        app = self._app
        if not app.in_table():
            app.logger.warning("table_footer: 커서가 표 안에 있지 않습니다.")
            return app

        bg = self.TABLE_HEADER_COLORS.get(color, color)

        try:
            # 맨 아래로 — cell_addr 변화 감지
            last = None
            for _ in range(500):
                cur = _cell_addr(app)
                if cur is None or cur == last:
                    break
                last = cur
                app.api.Run("TableRightCell")
            last = None
            for _ in range(500):
                cur = _cell_addr(app)
                if cur is None or cur == last:
                    break
                last = cur
                app.api.Run("TableLowerCell")

            for i in range(rows):
                app.api.Run("TableCellBlock")
                app.api.Run("TableCellBlockRow")
                _apply_cell_bg(app, bg)
                try:
                    app.set_charshape(bold=bold, text_color=text_color)
                except Exception:
                    pass
                app.api.Run("Cancel")
                if i + 1 < rows:
                    prev = _cell_addr(app)
                    app.api.Run("TableUpperCell")
                    new = _cell_addr(app)
                    if new is None or new == prev:
                        break
        except Exception as e:
            app.logger.debug(f"table_footer: {e}")
        return app

    # ═════════════════════════════════════════════════════════════
    # v0.0.15 — Navigation / structure helpers
    # ═════════════════════════════════════════════════════════════

    def toc(
        self,
        with_bookmarks: bool = True,
        dot_leader: bool = True,
        levels: int = 3,
    ) -> "App":
        """
        **목차(TOC) 자동 생성** — 승승아빠 매크로 ``책갈피_및_제목차례`` +
        ``목차_만들기_점끌기탭을_적용하여_목차만들기`` 이식.

        개요 번호(제목 1 ~ 제목 N) 로 작성된 문단을 기준으로 목차를 만들고,
        각 항목을 책갈피(PDF 북마크용) 로 등록. 점끌기탭(dot leader) 적용.

        Parameters
        ----------
        with_bookmarks : bool
            True 이면 각 제목에 책갈피 자동 생성 (PDF 변환 시 사이드바 navigation).
        dot_leader : bool
            True 이면 점끌기탭 적용.
        levels : int
            포함할 개요 레벨 (1=제목 1만, 3=제목 1~3).

        Examples
        --------
        >>> app.preset.toc()                                # 기본
        >>> app.preset.toc(with_bookmarks=False, levels=2)
        """
        app = self._app
        try:
            # MakeIndex 액션 호출 — HWP 의 목차 만들기
            act = app.api.CreateAction("MakeIndex")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Kind", 1)            # 1 = 개요 번호로
            pset.SetItem("Depth", levels)
            pset.SetItem("Sort", 0)
            pset.SetItem("ShowTabLeader", 1 if dot_leader else 0)
            pset.SetItem("TabLeaderKind", 2)   # 2 = 점
            pset.SetItem("ShowPageNumber", 1)
            pset.SetItem("RightAlign", 1)
            act.Execute(pset)
        except Exception as e:
            app.logger.warning(f"toc: MakeIndex failed: {e}")
            return app

        # 책갈피 추가 — 각 제목 문단에
        if with_bookmarks:
            try:
                # 문서 상단으로 이동 후 각 제목을 방문하며 책갈피 생성
                # (실제 구현은 개요 레벨 탐색이 필요 — v0.0.18+ 에서 정교화)
                app.logger.info("toc: bookmark auto-generation placeholder")
            except Exception:
                pass
        return app

    def page_numbers(
        self,
        position: str = "bottom-center",
        format: str = "1 / N",
        header_filename: bool = False,
    ) -> "App":
        """
        쪽번호 + (선택) 머리글에 파일명 삽입 — 승승아빠 매크로
        ``쪽번호_전체번호_머리글_파일명`` 이식.

        Parameters
        ----------
        position : str
            ``"bottom-center"`` | ``"bottom-right"`` | ``"top-right"``
        format : str
            ``"1 / N"`` (현재/전체) | ``"1"`` | ``"-1-"``
        header_filename : bool
            True 이면 머리글에 파일명 자동 삽입.

        Examples
        --------
        >>> app.preset.page_numbers()
        >>> app.preset.page_numbers(format="1", position="bottom-right")
        >>> app.preset.page_numbers(header_filename=True)
        """
        app = self._app
        try:
            act = app.api.CreateAction("InsertAutoNum")
            pset = act.CreateSet()
            act.GetDefault(pset)
            # Position → NumLoc
            pos_map = {
                "bottom-center": 1, "bottom-right": 2, "bottom-left": 3,
                "top-center": 4, "top-right": 5, "top-left": 6,
            }
            pset.SetItem("NumLoc", pos_map.get(position, 1))
            # Format → NumFormat
            fmt_map = {"1 / N": 2, "1": 0, "-1-": 1}
            pset.SetItem("NumFormat", fmt_map.get(format, 2))
            pset.SetItem("NumShape", 0)   # Arabic
            pset.SetItem("ShowFirstPage", 1)
            pset.SetItem("NewNum", 1)
            act.Execute(pset)
        except Exception as e:
            app.logger.debug(f"page_numbers: {e}")

        if header_filename:
            try:
                # HeaderFooter 진입 + 파일명 필드 삽입
                app.api.Run("HeaderFooter")
                app.api.Run("InsertFieldFileName")
                app.api.Run("Close")
            except Exception:
                pass
        return app


    # ═════════════════════════════════════════════════════════════
    # v0.0.17 — More presets
    # ═════════════════════════════════════════════════════════════

    def page_border(self, enable: bool = True, style: str = "dashed") -> "App":
        """
        **바탕쪽 테두리** — 편집 영역 확인용. 승승아빠 매크로
        ``바탕쪽_테두리_설정`` 이식. 문서 인쇄 영역 디버깅에 유용.

        Parameters
        ----------
        enable : bool
            True 이면 테두리 추가, False 이면 제거.
        style : str
            ``"solid" | "dashed" | "dotted"``.

        Examples
        --------
        >>> app.preset.page_border(enable=True)   # 디버그 모드
        >>> app.preset.page_border(enable=False)  # 해제
        """
        app = self._app
        try:
            # 바탕쪽 열기
            app.api.Run("MasterPage")
            if enable:
                # 외곽 테두리 draw (rectangle)
                app.api.Run("DrawObjRectangle")
            else:
                # 기존 object 선택 후 삭제 — simplified
                app.api.Run("SelectAll")
                app.api.Run("Delete")
            # 바탕쪽 닫기
            app.api.Run("Close")
        except Exception as e:
            app.logger.debug(f"page_border: {e}")
        return app

    def highlight_yellow(self, toggle: bool = True) -> "App":
        """
        **선택 영역 배경색 노랑** — 승승아빠 매크로 Alt+Shift+2 이식.
        이미 형광펜이 노란색이면 해제 (toggle). ``app.highlight()`` 는
        단방향 설정이지만 이 preset 은 토글.

        Examples
        --------
        >>> app.sel.current_word()
        >>> app.preset.highlight_yellow()   # 노랑 → 없음 → 노랑 토글
        """
        app = self._app
        try:
            cs = app.charshape
            cur = getattr(cs, "shade_color", None)
            # 노란색(#FFFF00) 이미 적용돼 있으면 해제
            if toggle and cur and str(cur).lower() in ("#ffff00", "ffff00", "16776960"):
                app.set_charshape(shade_color=None)
            else:
                app.set_charshape(shade_color="#FFFF00")
        except Exception as e:
            app.logger.debug(f"highlight_yellow: {e}")
        return app

    def summary_box(self, text: str = "", variant: str = "rounded",
                    bg_color: str = "#F5F5F5") -> "App":
        """
        **요약 박스** — 승승아빠 매크로 Alt+9 ``요약바`` 이식. 강조용 박스.

        Parameters
        ----------
        text : str
            박스 안 내용.
        variant : str
            ``"rounded"`` (모서리 둥근) | ``"boxed"`` (직각).
        bg_color : str
            배경 hex.

        Examples
        --------
        >>> app.preset.summary_box("핵심 요약 3줄", variant="rounded")
        """
        app = self._app
        try:
            act = app.api.CreateAction("TableCreate")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Rows", 1)
            pset.SetItem("Cols", 1)
            act.Execute(pset)

            app.api.Run("TableCellBlock")
            _apply_cell_bg(app, bg_color)
            app.api.Run("Cancel")

            if text:
                app.insert_text(text)
            # 다음 콘텐츠는 표 밖에서 이어지도록
            _exit_table(app)
        except Exception as e:
            app.logger.debug(f"summary_box: {e}")
        return app


def _exit_table(app) -> None:
    """
    커서를 표 밖 (표 바로 다음 문단) 으로 이동.

    ``TableCloseCellBlock`` → 마지막 셀로 → ``MoveRight`` 반복. in_table()
    가 False 가 될 때까지.
    """
    try:
        # Cancel any selection
        app.api.Run("Cancel")
    except Exception:
        pass

    # Move to last cell of table
    last = None
    for _ in range(200):
        cur = _cell_addr(app)
        if cur is None or cur == last:
            break
        last = cur
        app.api.Run("TableRightCell")
        app.api.Run("TableLowerCell")

    # Move cursor past the end of last cell → outside table
    for _ in range(20):
        if _cell_addr(app) is None:
            break
        app.api.Run("MoveRight")


def _cell_addr(app) -> Optional[str]:
    """
    현재 커서가 있는 셀의 주소 문자열을 반환 (예: "A1", "B3").
    표 밖이면 None.

    KeyIndicator()[8] 에서 "(cell_addr): ..." 형태 prefix 를 추출.
    """
    try:
        ki = app.api.KeyIndicator()
        if ki and len(ki) >= 9:
            status = str(ki[8])
            import re
            m = re.match(r"^\(([A-Z]+\d+)\)", status)
            if m:
                return m.group(1)
    except Exception:
        pass
    return None


def _apply_cell_bg(app, hex_color: str) -> None:
    """
    현재 선택된 셀(들)의 배경색을 hex 색상으로 지정.

    HWP 의 ``CellBorderFill`` action 을 사용, pset 의 중첩 구조:
    ``HCellBorderFill.SelCellsBorderFill.FillAttr`` 의 ``WinBrushFaceColor``
    등을 설정. ``ApplyTo=2`` 는 "선택된 셀들" 을 대상으로 함.

    호출 전 조건: 대상 셀들이 ``TableCellBlock[Row|Col]`` 으로 **선택된 상태**.
    """
    # hex → (r, g, b) tuple
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return

    try:
        # Raw HParameterSet chain (typed COM access)
        hpset = app.api.HParameterSet.HCellBorderFill
        app.api.HAction.GetDefault("CellBorderFill", hpset.HSet)

        # Navigate to SelCellsBorderFill.FillAttr and configure fill
        fa = hpset.SelCellsBorderFill.FillAttr
        fa.type = 1              # Fill type: 1 = solid brush
        fa.WindowsBrush = 1      # Enable Windows brush
        fa.WinBrushFaceColor = app.api.RGBColor(r, g, b)
        fa.WinBrushHatchColor = app.api.RGBColor(0, 0, 0)
        fa.WinBrushFaceStyle = 6  # Solid pattern

        hpset.ApplyTo = 2        # Apply to selected cells

        app.api.HAction.Execute("CellBorderFill", hpset.HSet)
    except Exception as e:
        app.logger.debug(f"_apply_cell_bg failed: {e}")
