"""
The :class:`App` class — main user-facing API for HWP automation.

Wraps an :class:`~hwpapi.core.engine.Engine` and provides high-level
methods for all common operations:

- **File I/O**: :meth:`open`, :meth:`save`, :meth:`save_block`,
  :meth:`insert_file`, :meth:`close`, :meth:`quit`
- **Text**: :meth:`insert_text`, :meth:`get_text`, :meth:`find_text`,
  :meth:`replace_all`, :meth:`get_selected_text`, :meth:`select_text`, :meth:`scan`
- **Formatting**: :meth:`charshape`, :meth:`get_charshape`,
  :meth:`set_charshape`, :meth:`get_parashape`, :meth:`set_parashape`
- **Images/Tables**: :meth:`insert_picture`, :meth:`set_cell_border`,
  :meth:`set_cell_color`
- **Page**: :meth:`setup_page`, :meth:`get_filepath`, :meth:`get_hwnd`
- **Window**: :meth:`set_visible`, :meth:`reload`
- **Actions/Fonts**: :meth:`create_action`, :meth:`create_parameterset`,
  :meth:`get_font_list`

Accessors attached as attributes (see :mod:`hwpapi.classes.accessors`):
- ``app.actions`` — dynamic action dispatcher
- ``app.move`` — caret navigation
- ``app.cell`` — table cell ops
- ``app.table`` — table structure ops
- ``app.page`` — page ops
"""
__all__ = ['Engine', 'Engines', 'Apps', 'App', 'Document', 'Documents',
           'move_to_line']

from contextlib import contextmanager
from pathlib import Path
import numbers
import hwpapi.constants as const
import sys
import warnings
from hwpapi.logging import get_logger

from hwpapi.actions import _Action, _Actions
from hwpapi.parametersets import ParaShape
import hwpapi.parametersets as parametersets
from hwpapi.classes import (
    MoveAccessor, CellAccessor, TableAccessor, PageAccessor,
    StylesAccessor, ControlsAccessor,
)
from hwpapi.classes.fields import Fields, Bookmarks, Hyperlinks
from hwpapi.classes.images import Images
from hwpapi.classes.selection import Selection
from hwpapi.classes.debug import Debug
from hwpapi.classes.convert import Convert
from hwpapi.classes.view import View
from hwpapi.classes.lint import Linter, Template, Config
from hwpapi.presets import Presets
from .engine import Engine, Engines, Apps
from hwpapi.functions import (
    check_dll,
    get_hwp_objects,
    dispatch,
    get_absolute_path,
    get_rgb_tuple,
    set_pset,
)



from .document import Document, Documents


class App:
    """
    `App` 클래스는 한컴오피스의 한/글 프로그램과 상호작용하기 위한 인터페이스를 제공합니다.

    이 클래스는 한/글 프로그램의 COM 객체와의 연동을 관리하고, 해당 객체에 대한 작업을 수행하는 메서드들을 포함합니다.

    Attributes
    ----------
    engine : Engine, optional
        사용할 엔진 객체입니다.
    is_visible : bool, optional
        한/글 프로그램의 가시성을 설정합니다.
    dll_path : str, optional
        사용할 DLL 파일의 경로입니다.

    Methods
    -------
    __init__(engine=None, is_visible=True, dll_path=None)
        `App` 클래스의 인스턴스를 초기화합니다.
    api()
        현재 사용 중인 엔진의 구현체를 반환합니다.
    __str__()
        `App` 인스턴스의 문자열 표현을 반환합니다.
    """

    def __init__(
        self,
        new_app=False,
        is_visible=True,
        dll_path=None,
        engine=None,
    ):
        """
        `App` 클래스의 인스턴스를 초기화합니다.

        Parameters
        ----------
        new_app: bool, optional
            별도 앱을 열지를 결정합니다.
        is_visible : bool, optional
            한/글 프로그램의 가시성 설정. 기본값은 True입니다.
        dll_path : str, optional
            사용할 DLL 파일의 경로. 기본값은 None입니다.
        engine : Engine, optional
            한/글 프로그램과의 상호작용을 위한 엔진. 기본값이 없을 경우, 새로운 엔진 인스턴스가 생성됩니다.
        Notes
        -----
        engine이 제공되지 않은 경우, `Engines`를 통해 엔진을 생성하거나 선택합니다.
        """
        self.logger = get_logger('core')
        self.logger.debug(f"Initializing App with new_app={new_app}, is_visible={is_visible}, dll_path={dll_path}")

        self._load(new_app=new_app, dll_path=dll_path, engine=engine)
        self.actions = _Actions(self)
        self.parameters = self.api.HParameterSet
        self.set_visible(is_visible)
        self.logger.info(f"App window visibility set to: {is_visible}")

        self.move = MoveAccessor(self)
        self.cell = CellAccessor(self)
        self.table = TableAccessor(self)
        self.page = PageAccessor(self)
        self.documents = Documents(self)
        self.styles = StylesAccessor(self)
        self.controls = ControlsAccessor(self)
        # v0.0.12+ collection accessors (v1.0 일관성 청사진 Phase 1)
        self.bookmarks = Bookmarks(self)
        self.hyperlinks = Hyperlinks(self)
        # v0.0.14+ preset / images / selection accessor
        # NOTE: `app.selection` 은 str property (선택된 텍스트) 유지.
        # 선택 동작 accessor 는 `app.sel` — `app.move` 와 대칭.
        self.images = Images(self)
        self.sel = Selection(self)
        self.preset = Presets(self)
        # v0.0.16+ debug accessor
        self.debug = Debug(self)
        # v0.0.17+ convert / view accessors
        self.convert = Convert(self)
        self.view = View(self)
        # v0.0.18+ lint / template / config accessors
        self.lint = Linter(self)
        self.template = Template(self)
        self.config = Config(self)
        self.logger.info("App initialized successfully with all accessors")

    def _load(self, new_app=False, engine=None, dll_path=None):
        self.logger.debug(f"Loading App with new_app={new_app}, engine={engine}, dll_path={dll_path}")

        if new_app:
            engine = Engine()
            self.logger.debug("Created new engine for new_app")
        if not engine:
            engines = Engines()
            engine = engines[-1] if len(engines) > 0 else Engine()
            self.logger.debug(f"Selected engine from engines collection")

        self.engine = engine
        self.logger.info(f"Engine loaded successfully")

        # Get DLL path (uses stable appdata location)
        if dll_path is None:
            from hwpapi.functions import get_hwp_dll_path
            dll_path = get_hwp_dll_path()

        # Register DLL if found
        if dll_path is not None:
            self.logger.info(f"Registering DLL: {dll_path}")
            check_dll(dll_path)
        else:
            self.logger.warning("No DLL file found - some functionality may be limited")

        try:
            self.api.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            self.logger.debug("Successfully registered FilePathCheckDLL module")
        except Exception as e:
            self.logger.warning(f"Failed to register FilePathCheckDLL module: {e}")
            warnings.warn(f"Failed to register FilePathCheckDLL module: {e}")

    @property
    def api(self):
        """
        현재 사용 중인 엔진의 구현체를 반환합니다.

        Returns
        -------
        object
            엔진의 구현체.
        """
        return self.engine.impl

    def __str__(self):
        """
        `App` 인스턴스의 문자열 표현을 반환합니다.

        Returns
        -------
        str
            `App` 인스턴스를 설명하는 문자열.
        """
        return f"<Hwp App: {self.get_filepath()}>"

    __repr__ = __str__

    def set_visible(self, is_visible=True, window_i=0):
        """
        한컴오피스 Hwp 프로그램 창의 가시성을 설정합니다.

        이 함수는 `is_visible` 매개변수에 따라 한컴오피스 Hwp 프로그램의 창을 표시하거나 숨기는 데 사용됩니다.
        `window_i` 매개변수는 표시하거나 숨길 창의 인덱스를 지정합니다.

        매개변수
        ----------
        is_visible : bool, optional
            프로그램 창의 가시성을 나타내는 불린 값.
            True이면 창이 표시되고, False이면 숨겨집니다. 기본값은 True입니다.
        window_i : int, optional
            수정할 창의 인덱스. 기본값은 0입니다.

        사용 예시
        --------
        >>> app = App()
        >>> app.set_visible(is_visible=True, window_i=0)
        >>> app.set_visible(is_visible=False, window_i=1)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_visible")
        self.api.XHwpWindows.Item(window_i).Visible = is_visible

    def create_parameterset(self, key: str):
        logger = get_logger('core')
        logger.debug(f"Calling create_parameterset")
        return getattr(self.api.HParameterSet, f"H{key}")

    # ═════════════════════════════════════════════════════════════
    # Standard properties — xlwings-style convenience
    # ═════════════════════════════════════════════════════════════

    @property
    def text(self) -> str:
        """
        현재 활성 문서의 **전체 텍스트** (읽기/쓰기).

        ``doc.text`` 로도 접근 가능 — 해당 문서가 자동 활성화된 뒤 값을 반환.

        Setting 하면 기존 내용을 전부 지우고 새 텍스트로 교체합니다.

        Examples
        --------
        >>> text = app.text                      # 전체 문서 읽기
        >>> app.text = "완전히 새 내용\\n두 번째 줄"  # 전체 교체
        """
        txt = self.get_text(
            spos=const.ScanStartPosition.Document,
            epos=const.ScanEndPosition.Document,
        )
        if isinstance(txt, tuple):
            return str(txt[0]) if txt else ""
        return str(txt) if txt else ""

    @text.setter
    def text(self, value: str):
        self.select_all()
        try:
            self.api.Run("Delete")
        except Exception:
            self.api.Run("DeleteBack")
        if value:
            self.insert_text(value)

    @property
    def visible(self) -> bool:
        """
        HWP 메인 창의 가시성 (property form of :meth:`set_visible`).

        - Read: ``XHwpWindows.Item(0).Visible`` 조회.
        - Write: ``self.set_visible(bool(value))`` 위임.

        단일 창 제어에는 이 property 를 권장합니다. 여러 창 중 특정
        ``window_i`` 를 지정하려면 :meth:`set_visible` 메소드를 사용.

        See Also
        --------
        set_visible : 다중 창 지원 메소드 버전.
        """
        try:
            return bool(self.api.XHwpWindows.Item(0).Visible)
        except Exception:
            return False

    @visible.setter
    def visible(self, value: bool):
        self.set_visible(bool(value))

    @property
    def version(self) -> str:
        """HWP 어플리케이션 버전 문자열 (예: ``"12, 0, 1, 3335"``)."""
        try:
            return str(self.api.Version)
        except Exception:
            return ""

    @property
    def page_count(self) -> int:
        """현재 문서의 페이지 수."""
        try:
            return int(self.api.PageCount)
        except Exception:
            return 0

    @property
    def current_page(self) -> int:
        """현재 커서가 위치한 페이지 번호 (1부터)."""
        try:
            # KeyIndicator tuple index 5 is the actual current page
            return int(self.api.KeyIndicator()[5])
        except Exception:
            return 0

    @property
    def selection(self) -> str:
        """
        현재 선택된 텍스트 (property alias for :meth:`get_selected_text`).

        두 API 는 동일 결과를 반환합니다. 짧은 property 형태는 Pythonic
        값 접근에 적합 (``app.selection`` vs ``app.get_selected_text()``).

        See Also
        --------
        get_selected_text : 동등한 메소드 형태.
        """
        return self.get_selected_text() or ""

    # ═════════════════════════════════════════════════════════════
    # Standard action shortcuts
    # ═════════════════════════════════════════════════════════════

    def select_all(self):
        """문서 전체 선택. ``SelectAll`` 액션 alias."""
        return self.api.Run("SelectAll")

    def clear(self):
        """현재 활성 문서의 모든 내용을 삭제."""
        self.select_all()
        try:
            return self.api.Run("Delete")
        except Exception:
            return self.api.Run("DeleteBack")

    def undo(self):
        """마지막 작업 되돌리기. ``Undo`` 액션 alias."""
        return self.api.Run("Undo")

    def redo(self):
        """되돌린 작업 다시 실행. ``Redo`` 액션 alias."""
        return self.api.Run("Redo")

    def copy(self):
        """선택 영역을 클립보드로 복사."""
        return self.api.Run("Copy")

    def paste(self):
        """클립보드 내용을 커서 위치에 붙여넣기."""
        return self.api.Run("Paste")

    def cut(self):
        """선택 영역을 잘라내기 (클립보드로 이동)."""
        return self.api.Run("Cut")

    def delete(self):
        """선택 영역 또는 커서 위치 글자 삭제."""
        try:
            return self.api.Run("Delete")
        except Exception:
            return self.api.Run("DeleteBack")

    def insert_page_break(self):
        """페이지 나누기 삽입. Fluent — chain 가능 (반환: ``self``)."""
        self.api.Run("BreakPage")
        return self

    def insert_line_break(self):
        """줄 나누기 삽입 (문단 유지, 강제 줄바꿈). Fluent — chain 가능."""
        self.api.Run("BreakLine")
        return self

    def insert_paragraph_break(self):
        """문단 나누기 삽입. Fluent — chain 가능 (반환: ``self``)."""
        self.api.Run("BreakPara")
        return self

    def insert_tab(self):
        """탭 문자 삽입. Fluent — chain 가능 (반환: ``self``)."""
        self.api.Run("InsertTab")
        return self

    # ═════════════════════════════════════════════════════════════
    # High-level insertion helpers
    # ═════════════════════════════════════════════════════════════

    def insert_heading(self, text: str, level: int = 1):
        """
        제목(heading) 삽입 — 굵게 + 큰 글씨로 포맷 후 줄바꿈.

        Parameters
        ----------
        text : str
            제목 텍스트.
        level : int
            제목 레벨 (1-4). 크기: 1=20pt, 2=16pt, 3=14pt, 4=12pt.

        Examples
        --------
        >>> app.insert_heading("1장 도입", level=1)
        >>> app.insert_heading("1.1 배경", level=2)
        """
        sizes = {1: 2000, 2: 1600, 3: 1400, 4: 1200}
        self.styled_text(text, bold=True, height=sizes.get(level, 1000))
        self.insert_paragraph_break()
        return self

    def insert_table(self, rows: int = None, cols: int = None,
                     data=None, headers=None):
        """
        표 생성 — 빈 표, 데이터로 채워진 표, 또는 pandas DataFrame 에서.

        Parameters
        ----------
        rows, cols : int
            행/열 수 (``data`` 없을 때).
        data : list[list] | pandas.DataFrame | None
            2차원 리스트 또는 DataFrame. 지정 시 rows/cols 자동 계산.
            DataFrame 이면 columns 가 자동으로 ``headers`` 로 사용됨.
        headers : list | None
            있으면 첫 행에 삽입되고 굵게 표시.

        Examples
        --------
        >>> app.insert_table(rows=3, cols=4)         # 3x4 빈 표
        >>> app.insert_table(
        ...     data=[[1, 2, 3], [4, 5, 6]],
        ...     headers=["A", "B", "C"],
        ... )                                        # 3x3 표 자동 채움

        DataFrame 직접 삽입:

        >>> import pandas as pd
        >>> df = pd.DataFrame({"지역": ["서울", "경기"], "매출": [1850, 1320]})
        >>> app.insert_table(data=df)                # columns 자동 헤더
        """
        # Detect DataFrame — extract columns as headers and values as data
        try:
            import pandas as pd  # optional dependency
            if isinstance(data, pd.DataFrame):
                if headers is None:
                    headers = data.columns.tolist()
                data = data.values.tolist()
        except ImportError:
            pass

        if data is not None:
            all_rows = ([list(headers)] if headers else []) + list(data)
            rows = len(all_rows)
            cols = max(len(r) for r in all_rows)
        else:
            if not rows or not cols:
                raise ValueError("rows+cols 또는 data 중 하나는 필요합니다")
            all_rows = None

        act = self.api.CreateAction("TableCreate")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Rows", rows)
        pset.SetItem("Cols", cols)
        act.Execute(pset)

        if all_rows:
            for ri, row in enumerate(all_rows):
                for ci in range(cols):
                    cell = row[ci] if ci < len(row) else ""
                    if ri == 0 and headers is not None:
                        self.styled_text(str(cell), bold=True)
                    else:
                        self.insert_text(str(cell))
                    self.api.Run("TableRightCell")
        return self

    def insert_hyperlink(self, text: str, url: str):
        """
        하이퍼링크 삽입.

        Examples
        --------
        >>> app.insert_hyperlink("GitHub", "https://github.com/JunDamin/hwpapi")
        """
        act = self.api.CreateAction("InsertHyperlink")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Text", text)
        pset.SetItem("Command", url)
        act.Execute(pset)
        return self

    def insert_bookmark(self, name: str):
        """
        현재 커서 위치에 책갈피 삽입.

        Examples
        --------
        >>> app.insert_bookmark("ch1")
        """
        act = self.api.CreateAction("Bookmark")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Name", name)
        act.Execute(pset)
        return self

    def new_document(self, is_tab: bool = True) -> "Document":
        """
        새 빈 문서 추가 — :meth:`Documents.add` 의 편의 alias.

        같은 결과: ``app.documents.add(is_tab=True)``.

        Examples
        --------
        >>> doc = app.new_document()
        >>> doc.insert_text("새 문서")

        See Also
        --------
        Documents.add : 정식 경로 (``app.documents.add()``).
        """
        return self.documents.add(is_tab=is_tab)

    # ═════════════════════════════════════════════════════════════
    # Phase C — 편의 헬퍼 (페이지 탐색, 형광펜, 상태, 이미지 export)
    # ═════════════════════════════════════════════════════════════

    def goto_page(self, page: int) -> bool:
        """
        지정 페이지로 커서 이동 (1-indexed).

        HWP의 ``Goto`` 액션을 ``GotoE`` pset 으로 호출. pset 의 주요 키:

        - ``SetSelectionIndex = 1`` → 페이지 번호로 이동
        - ``DialogResult = page`` → 페이지 번호 (1-indexed)

        Parameters
        ----------
        page : int
            이동할 페이지 번호 (1 = 첫 페이지).

        Examples
        --------
        >>> app.goto_page(5)     # 5페이지로
        >>> app.goto_page(1)     # 첫 페이지로
        """
        # Clamp to valid range
        try:
            total = self.page_count
            n = max(1, min(int(page), total if total > 0 else int(page)))
        except Exception:
            n = int(page)

        # Preferred: HAction.Execute("Goto", HGotoE.HSet) with
        # DialogResult = n, SetSelectionIndex = 1
        try:
            pset = self.api.HParameterSet.HGotoE
            self.api.HAction.GetDefault("Goto", pset.HSet)
            pset.HSet.SetItem("DialogResult", int(n))
            pset.SetSelectionIndex = 1
            self.api.HAction.Execute("Goto", pset.HSet)

            # Refine: HAction typically lands somewhere near the page;
            # nudge to exact page with MovePageUp/Down.
            cur = self.current_page
            if cur and n != cur:
                if n < cur:
                    for _ in range(cur - n):
                        self.api.Run("MovePageUp")
                else:
                    for _ in range(n - cur):
                        self.api.Run("MovePageDown")
            return True
        except Exception:
            pass

        # Fallback: MoveDocBegin + N-1 x MovePageDown
        try:
            self.api.Run("MoveDocBegin")
            for _ in range(max(0, n - 1)):
                self.api.Run("MovePageDown")
            return True
        except Exception:
            return False

    def highlight(self, color="#FFFF00") -> bool:
        """
        현재 선택 영역에 **형광펜** 효과 적용 (``shade_color`` 와 별개).

        Parameters
        ----------
        color : str | int | Color | tuple
            형광펜 색. 문자열 (``"#FFFF00"``) / Color / (r, g, b) 튜플 /
            정수 HWP 색상 모두 지원.

        Examples
        --------
        >>> app.find_text("중요")
        >>> app.highlight("#FFFF00")            # 노란색
        >>> app.highlight((255, 150, 150))       # 분홍 RGB
        """
        # HAction 경로로 호출 (pset 접근이 action.CreateSet 보다 정확)
        try:
            from hwpapi.parametersets import Color
            if isinstance(color, tuple) and len(color) == 3:
                col = Color.from_rgb(*color)
            elif isinstance(color, Color):
                col = color
            else:
                col = Color(color)

            pset = self.api.HParameterSet.HMarkpenShape
            self.api.HAction.GetDefault("MarkPenShape", pset.HSet)
            pset.Color = col.raw if col.raw is not None else 0
            return bool(self.api.HAction.Execute("MarkPenShape", pset.HSet))
        except Exception:
            return False

    @property
    def status(self) -> dict:
        """
        상태바 정보 — ``KeyIndicator()`` 반환값을 dict 로.

        KeyIndicator 는 튜플을 반환: ``(col_left, row_left, sect, page,
        total, overtype, char_count, insertmode, extend, lockmode, ...)``.

        Examples
        --------
        >>> app.status
        {'page': 3, 'total_pages': 12, 'section': 1, 'column': 1,
         'overtype': 0, 'char_count': 4521, 'insert_mode': 1}
        """
        try:
            raw = self.api.KeyIndicator()
        except Exception:
            return {}
        if not raw or not isinstance(raw, (tuple, list)):
            return {}
        # HWP KeyIndicator tuple format:
        # (0)has_caret, (1)section, (2)column, (3)line, (4)print_page,
        # (5)page, (6)overtype, (7)char_index, (8)insert_mode_str
        keys = [
            "has_caret", "section", "column", "line",
            "print_page", "page",
            "overtype", "char_index", "insert_mode",
        ]
        out = {}
        for i, key in enumerate(keys):
            if i < len(raw):
                out[key] = raw[i]
        # Fill in total_pages separately (KeyIndicator doesn't provide it)
        try:
            out["total_pages"] = int(self.api.PageCount)
        except Exception:
            pass
        return out

    def save_page_image(self, page: int, path: str,
                         dpi: int = 100, depth: int = 8,
                         format: int = -1) -> bool:
        """
        지정 페이지를 이미지 파일로 저장.

        HWP의 ``CreatePageImage(Path, pgno, resolution, depth, Format)``
        를 감쌉니다.

        Parameters
        ----------
        page : int
            페이지 번호 (1-indexed).
        path : str
            저장 경로. 확장자로 포맷 결정 (``.bmp``/``.png``/``.jpg``/``.gif``).
            확장자 기반 자동 인식이 안되는 경우 ``format`` 파라미터로 명시.
        dpi : int
            해상도 (기본 100). HWP 는 값이 너무 낮거나 높으면 무시하거나
            실패할 수 있습니다.
        depth : int
            색 깊이 비트수 (8 / 24 / 32). 기본 8.
        format : int
            ``-1`` (자동, 확장자 기반) / ``0`` (BMP) / ``1`` (GIF) /
            ``2`` (JPEG) / ``3`` (PNG).

        Returns
        -------
        bool
            성공 여부 (파일 실제 생성 확인 포함).

        Examples
        --------
        >>> app.save_page_image(1, "page1.bmp")              # BMP 추천
        >>> app.save_page_image(2, "page2.png", format=3)    # PNG
        """
        import os
        try:
            ok = bool(self.api.CreatePageImage(
                str(path), int(page) - 1,
                int(dpi), int(depth), int(format),
            ))
            # CreatePageImage returns True even when the file is empty.
            # Verify by checking file size.
            if ok and os.path.isfile(path) and os.path.getsize(path) > 0:
                return True
        except Exception:
            pass
        # Fallback: try BMP (known working)
        try:
            bmp_path = os.path.splitext(path)[0] + ".bmp"
            ok = bool(self.api.CreatePageImage(
                bmp_path, int(page) - 1, 100, 8, 0
            ))
            return ok and os.path.isfile(bmp_path) and os.path.getsize(bmp_path) > 0
        except Exception:
            return False

    def save_all_page_images(self, out_dir: str, prefix: str = "page",
                               format: str = "png", dpi: int = 300) -> list:
        """
        모든 페이지를 이미지로 저장 — 생성된 파일 경로 리스트 반환.

        Examples
        --------
        >>> paths = app.save_all_page_images("out/")
        >>> len(paths)
        12
        """
        import os
        os.makedirs(out_dir, exist_ok=True)
        out_paths = []
        total = self.page_count
        for n in range(1, total + 1):
            p = os.path.join(out_dir, f"{prefix}{n:04d}.{format}")
            if self.save_page_image(n, p, dpi=dpi):
                out_paths.append(p)
        return out_paths

    # ── 단위 변환 헬퍼 (App instance methods) ─────────────────

    # ── 단위 변환 ────────────────────────────────────────────
    # App instance 메소드는 HWP COM (MiliToHwpUnit / PointToHwpUnit)
    # 을 호출. HWP 실행 없이 순수 파이썬으로 변환만 하고 싶으면:
    #   ``from hwpapi.functions import to_hwpunit, from_hwpunit``
    # 를 사용 (unit string 파싱 + 단위 전환 모두 지원).

    def mm_to_hwpunit(self, mm: float) -> int:
        """밀리미터 → HWPUNIT. HWP 엔진의 ``MiliToHwpUnit`` 호출."""
        try:
            return int(self.api.MiliToHwpUnit(float(mm)))
        except Exception:
            return int(mm * 283)

    def point_to_hwpunit(self, pt: float) -> int:
        """포인트 → HWPUNIT. HWP 엔진의 ``PointToHwpUnit`` 호출."""
        try:
            return int(self.api.PointToHwpUnit(float(pt)))
        except Exception:
            return int(pt * 100)

    def hwpunit_to_mm(self, hu: int) -> float:
        """HWPUNIT → 밀리미터. (1 mm ≈ 283 HWPUNIT)"""
        return float(hu) / 283.0

    def hwpunit_to_point(self, hu: int) -> float:
        """HWPUNIT → 포인트. (1 pt = 100 HWPUNIT)"""
        return float(hu) / 100.0

    def rgb_color(self, r: int, g: int, b: int) -> int:
        """
        RGB 값을 HWP 내부 색상 정수 (BBGGRR) 로 변환.

        **반환**: 정수 (HWP COM API 에 직접 전달할 때 유용).

        타입 안전성 있는 ``Color`` 객체가 필요하면
        :meth:`~hwpapi.parametersets.Color.from_rgb` 를 사용하세요.

        Examples
        --------
        >>> app.rgb_color(255, 0, 0)    # 빨강 → int 255
        255
        >>> app.rgb_color(0, 255, 0)    # 초록 → int 65280
        65280

        >>> # Or type-safe Color wrapper
        >>> from hwpapi.parametersets import Color
        >>> Color.from_rgb(255, 0, 0)   # Color('#ff0000')

        See Also
        --------
        hwpapi.parametersets.Color.from_rgb : Color 객체 생성자.
        """
        try:
            return int(self.api.RGBColor(int(r), int(g), int(b)))
        except Exception:
            # HWP BBGGRR encoding
            return (int(b) << 16) | (int(g) << 8) | int(r)

    # ═════════════════════════════════════════════════════════════
    # Dialog suppression — set_message_box_mode / silenced()
    # ═════════════════════════════════════════════════════════════

    def register_security_module(
        self,
        module_name: str = "FilePathCheckerModuleExample",
        dll_path: str = None,
    ) -> bool:
        """
        보안 모듈 (FilePathChecker) 명시적 등록.

        HWP 는 외부 프로그램이 파일을 읽거나 쓸 때 보안 경고 dialog 를
        띄웁니다. ``RegisterModule`` 로 ``FilePathCheckerModule`` 을
        등록하면 이 경고를 사전에 차단합니다.

        Notes
        -----
        ``App.__init__`` 에서 자동으로 한 번 호출됩니다 (DLL 경로
        자동 탐색). 따라서 일반적으로 직접 호출할 필요는 없습니다.
        다른 DLL 을 쓰거나 명시적 재등록이 필요할 때만 사용.

        Parameters
        ----------
        module_name : str
            HWP 에 등록할 모듈 식별자.
        dll_path : str | None
            DLL 파일 경로. ``None`` 이면 ``hwpapi`` 패키지에 포함된
            기본 DLL 사용.

        Returns
        -------
        bool
            등록 성공 여부.

        Examples
        --------
        >>> # 보통은 자동 등록되지만, 명시적 재등록:
        >>> app.register_security_module()

        >>> # 커스텀 DLL:
        >>> app.register_security_module(
        ...     module_name="MyCustomChecker",
        ...     dll_path=r"C:\\my\\custom.dll",
        ... )
        """
        if dll_path is None:
            try:
                from hwpapi.functions import get_hwp_dll_path
                dll_path = get_hwp_dll_path("hwpapi", "FilePathCheckerModuleExample.dll")
            except Exception:
                dll_path = None
        try:
            self.api.RegisterModule(module_name, dll_path or "")
            return True
        except Exception as e:
            try:
                self.logger.warning(f"register_security_module failed: {e}")
            except Exception:
                pass
            return False

    def get_message_box_mode(self) -> int:
        """현재 다이얼로그 모드 반환 (HWP `GetMessageBoxMode`)."""
        try:
            return int(self.api.GetMessageBoxMode())
        except Exception:
            return 0

    # ── HWP MessageBox bitfield 상수 ─────────────────────────
    #
    # SetMessageBoxMode 인자는 6개 nibble (4 bit) 로 6가지 dialog 타입을
    # 각각 어떻게 자동 응답할지 지정합니다.
    #
    #   nibble 0 (0xF)        확인 dialog (OK only)
    #   nibble 1 (0xF0)       확인/취소 dialog (OK/Cancel)
    #   nibble 2 (0xF00)      종료/재시도/무시 dialog (Abort/Retry/Ignore)
    #   nibble 3 (0xF000)     예/아니오/취소 dialog (Yes/No/Cancel)
    #   nibble 4 (0xF0000)    예/아니오 dialog (Yes/No)
    #   nibble 5 (0xF00000)   재시도/취소 dialog (Retry/Cancel)
    #
    # 각 nibble 안에서 1=첫 버튼, 2=둘째, 4=셋째, F=해제(수동).

    # 미리 정의된 preset (App 클래스 상수)
    SILENCE_ALL_YES = 0x111111   # 모든 dialog 첫 버튼 (OK/Yes/Abort/Yes/Yes/Retry)
    SILENCE_ALL_NO  = 0x222222   # 모든 dialog 둘째 버튼 (—/Cancel/Retry/No/No/Cancel)
    SILENCE_RESET   = 0xFFFFFF   # 모든 자동 응답 해제 (수동 모드)

    # 단일 카테고리 preset — hwp-mcp 와의 호환·간편 사용
    SILENCE_OK_AUTO       = 0x00000001  # 확인 dialog 자동 OK
    SILENCE_SAVE_YES      = 0x00010000  # 예/아니오 → YES (저장)
    SILENCE_SAVE_NO       = 0x00020000  # 예/아니오 → NO (저장 안함)
    SILENCE_OKCANCEL_OK   = 0x00000010  # 확인/취소 → OK
    SILENCE_OKCANCEL_NO   = 0x00000020  # 확인/취소 → Cancel

    # hwp-mcp 호환 alias (jkf87/hwp-mcp 의 set_message_box_mode 도큐먼테이션)
    #   - 0x00010000 → "확인 버튼 자동 클릭" (= SILENCE_SAVE_YES)
    #   - 0x00020000 → "취소 버튼 자동 클릭" (= SILENCE_SAVE_NO)
    #   - 0x00100000 → "저장 안함 선택"
    SILENCE_NO_SAVE       = 0x00100000  # 저장 안함 선택 (hwp-mcp)

    def set_message_box_mode(self, mode: int) -> int:
        """
        HWP 다이얼로그 처리 모드 지정 — 이전 모드값을 반환.

        6 개 dialog 카테고리에 대해 각각 자동 응답을 설정합니다.
        클래스 상수 :attr:`SILENCE_ALL_YES`, :attr:`SILENCE_ALL_NO`,
        :attr:`SILENCE_RESET` 사용 권장.

        ====== ===========================================================
        Bits   Dialog Type                            첫/둘째/셋째 버튼
        ====== ===========================================================
        0xF    확인                                    1=OK, F=수동
        0xF0   확인/취소                                1=OK, 2=Cancel
        0xF00  종료/재시도/무시                         1=Abort, 2=Retry, 4=Ignore
        0xF000 예/아니오/취소                            1=Yes, 2=No, 4=Cancel
        0xF0000 예/아니오                                1=Yes, 2=No
        0xF00000 재시도/취소                              1=Retry, 2=Cancel
        ====== ===========================================================

        Examples
        --------
        >>> # 모든 dialog 자동 첫 버튼 (YES/OK)
        >>> app.set_message_box_mode(App.SILENCE_ALL_YES)
        >>> # 또는 직접: 0x111111

        >>> # 예/아니오 만 No 자동, 나머지 수동
        >>> app.set_message_box_mode(0x020000)

        대부분 :meth:`silenced` context manager 사용 권장.
        """
        try:
            return int(self.api.SetMessageBoxMode(int(mode)))
        except Exception:
            return 0

    @contextmanager
    def silenced(self, mode=SILENCE_ALL_YES):
        """
        다이얼로그 자동응답 **context manager** (scoped MessageBox mode).

        ``with app.silenced(...):`` 블록 내부의 모든 HWP 확인/경고 dialog
        를 자동으로 처리하고, 블록 **종료 시 이전 mode 로 복원**됩니다.

        영구 적용이 필요하면 :meth:`set_message_box_mode` 를 직접 호출하세요.
        한 번만 적용 후 자동 복원하려면 항상 :meth:`silenced` 를 쓰세요.

        Parameters
        ----------
        mode : int | str
            응답 방식.

            **str preset** (권장):

            - ``"yes"`` (기본) — 모든 dialog 첫 버튼 (OK/Yes/Abort/Yes/Yes/Retry)
            - ``"no"`` — 모든 dialog 둘째 버튼 (Cancel/Retry/No/No/Cancel)
            - ``"reset"`` — 자동 응답 해제 (사용자 수동 입력)

            **int** — 직접 비트필드 지정 (:meth:`set_message_box_mode` 참고).

        Examples
        --------
        파일 일괄 처리 시 "저장하시겠습니까?" 자동 YES:

        >>> with app.silenced():               # = silenced("yes")
        ...     for path in paths:
        ...         app.open(path)
        ...         app.replace_all("2025", "2026")
        ...         app.save(path)

        저장하지 않고 모두 닫기:

        >>> with app.silenced("no"):
        ...     for doc in app.documents:
        ...         doc.close()

        세밀한 제어 — 예/아니오만 NO, 확인은 자동 OK:

        >>> with app.silenced(0x020001):
        ...     ...

        Notes
        -----
        ``XHwpMessageBox`` 으로 띄우는 일반 정보 박스는 별도. 대부분은
        ``SetMessageBoxMode`` 로 충분히 처리됩니다. 그래도 dialog 가
        나타나면 :class:`~smoke_scenarios.dismiss_dialog_hwnd` 같은
        외부 dismisser 를 병행하세요.
        """
        # Preset 처리
        if isinstance(mode, str):
            preset = mode.lower()
            mode_int = {
                # 전 카테고리
                "yes": self.SILENCE_ALL_YES,
                "no": self.SILENCE_ALL_NO,
                "reset": self.SILENCE_RESET,
                # hwp-mcp 호환 (단일 카테고리)
                "ok": self.SILENCE_OK_AUTO,             # 0x00000001
                "save": self.SILENCE_SAVE_YES,           # 0x00010000
                "save_yes": self.SILENCE_SAVE_YES,       # 0x00010000
                "save_no": self.SILENCE_SAVE_NO,         # 0x00020000
                "no_save": self.SILENCE_NO_SAVE,         # 0x00100000
                "okcancel_ok": self.SILENCE_OKCANCEL_OK,  # 0x00000010
                "okcancel_no": self.SILENCE_OKCANCEL_NO,  # 0x00000020
            }.get(preset)
            if mode_int is None:
                raise ValueError(
                    f"Unknown silenced preset {mode!r}. "
                    f"Valid: 'yes', 'no', 'reset', 'ok', 'save', 'save_yes', "
                    f"'save_no', 'no_save', 'okcancel_ok', 'okcancel_no', "
                    f"or an int bitfield."
                )
            mode = mode_int

        prev = self.get_message_box_mode()
        self.set_message_box_mode(int(mode))
        try:
            yield
        finally:
            self.set_message_box_mode(prev)

    @contextmanager
    def suppress_errors(self):
        """
        에러/경고 dialog 를 모두 자동 ABORT 처리 + 예외 swallowing.

        :meth:`silenced` 와 다르게 **에러 dialog 도** 자동 닫고, 블록
        내부에서 발생하는 Python 예외도 ``logger.warning`` 으로 로그만
        남기고 무시합니다. 대량 자동화에서 한두 파일이 실패해도 전체
        루프를 계속 진행하고 싶을 때 사용.

        Examples
        --------
        >>> for path in many_paths:
        ...     with app.suppress_errors():
        ...         app.open(path)        # 일부 깨진 파일이어도 계속
        ...         app.save(path + ".out")
        """
        prev = self.get_message_box_mode()
        # 에러/abort dialog 는 'abort' (첫 버튼 = 0x100), 나머지는 yes
        # 종료/재시도/무시 카테고리에서 1=Abort
        self.set_message_box_mode(0x111111)
        try:
            yield
        except Exception as exc:
            try:
                self.logger.warning(f"suppress_errors caught: {exc}")
            except Exception:
                pass
        finally:
            self.set_message_box_mode(prev)

    @contextmanager
    def batch_mode(self, hide: bool = True):
        """
        **대량 처리 성능 부스터** — 화면 갱신/창 표시 억제. v0.0.16+.

        수백 개 문서를 열고 편집/저장할 때 HWP 가 매 작업마다 화면을
        다시 그리는 비용이 누적됩니다. 이 context 는 진입 시 창을 숨기고
        ``SetMessageBoxMode`` 를 SILENCE_ALL_YES 로 맞추고, 종료 시
        원상복구.

        Parameters
        ----------
        hide : bool
            True (기본) 이면 블록 내부에서 창 숨김. False 이면 visible 유지
            (dialog 억제만).

        Examples
        --------
        >>> with app.batch_mode():
        ...     for path in paths:
        ...         app.open(path)
        ...         app.fields.update(merge_data)
        ...         app.save(f"out/{path.name}")
        # 일반 대비 5~10배 빠름, 블록 종료 시 창/모드 자동 복원

        Notes
        -----
        내부적으로 :meth:`silenced` + ``set_visible(False)`` + (가능하면)
        ``ScrollUnfollow`` 같은 HWP redraw suppression 을 조합.
        """
        prev_mode = self.get_message_box_mode()
        prev_visible = None
        try:
            prev_visible = bool(self.api.XHwpWindows.Active_XHwpWindow.Visible)
        except Exception:
            pass

        # Enter batch mode
        self.set_message_box_mode(self.SILENCE_ALL_YES)
        if hide and prev_visible:
            try:
                self.set_visible(False)
            except Exception:
                pass

        # ScrollFollowTo off (if supported)
        try:
            self.api.Run("FollowActiveWindowOff")
        except Exception:
            pass

        try:
            yield self
        finally:
            # Restore
            self.set_message_box_mode(prev_mode)
            if hide and prev_visible:
                try:
                    self.set_visible(True)
                except Exception:
                    pass
            try:
                self.api.Run("FollowActiveWindowOn")
            except Exception:
                pass

    @contextmanager
    def undo_group(self, description: str = ""):
        """
        **Undo 경계** — 블록 내부의 여러 편집을 사용자가 단일 Ctrl+Z 로
        되돌릴 수 있게 묶음. v0.0.16+.

        Parameters
        ----------
        description : str
            undo history 에 표시될 설명 (선택).

        Examples
        --------
        >>> with app.undo_group("대량 포맷팅"):
        ...     for para in range(100):
        ...         app.set_charshape(bold=True)
        ...         app.move.line.next()
        # 사용자가 Ctrl+Z 한 번으로 100 번 작업이 모두 rollback

        Notes
        -----
        HWP 는 공식 undo group API 가 제한적. 현재 구현은 best-effort:
        진입 시점을 기록하고 종료 시까지의 모든 작업을 묶으려 시도합니다.
        """
        # HWP 의 공식 undo boundary 액션 시도 (없을 수 있음)
        try:
            self.api.Run("SetUndoBegin")
        except Exception:
            pass
        try:
            yield self
        finally:
            try:
                self.api.Run("SetUndoEnd")
            except Exception:
                pass

    # ═════════════════════════════════════════════════════════════
    # Field API (Mail Merge / Forms) — Phase A
    # ═════════════════════════════════════════════════════════════
    #
    # HWP 필드는 문서 내에 이름표가 붙은 **값 주입 지점** 입니다. 템플릿
    # 문서에서 `{{name}}` 같은 자리를 필드로 만든 뒤, 스크립트가 행별로
    # 데이터를 채워넣어 여러 문서를 생성하는 mail merge 가 핵심 용도.

    def create_field(self, name: str, memo: str = "", direction: str = ""):
        """
        현재 커서 위치에 **누름틀(필드)** 을 생성.

        Parameters
        ----------
        name : str
            필드 이름 (나중에 ``set_field(name, value)`` 로 값 주입).
        memo : str
            필드 설명 메모.
        direction : str
            필드에 표시될 안내 문구 (예: ``"이름 입력"``).

        Examples
        --------
        >>> app.create_field("customer_name", direction="고객명")
        >>> app.set_field("customer_name", "홍길동")
        """
        return self.api.CreateField(direction, memo, name)

    def set_field(self, name: str, value) -> bool:
        """
        이름으로 지정한 필드에 값 주입.

        Parameters
        ----------
        name : str
            필드 이름. 동일 이름의 필드가 여러 개면 전부에 같은 값 주입.
        value : Any
            ``str()`` 로 변환되어 삽입.
        """
        return bool(self.api.PutFieldText(name, str(value)))

    def get_field(self, name: str) -> str:
        """이름으로 지정한 필드의 현재 값 반환."""
        return self.api.GetFieldText(name) or ""

    @property
    def field_names(self) -> list:
        """
        현재 문서의 모든 필드 이름 리스트 (legacy).

        v0.0.12+ 권장: ``list(app.fields)`` 또는 ``for n in app.fields:``

        Examples
        --------
        >>> app.field_names
        ['customer_name', 'order_date', 'total']
        """
        raw = self.api.GetFieldList(0, 0) or ""
        # HWP separates field names by control char \x02 (STX)
        # Some versions use \t or \n
        for sep in ("\x02", "\t", "\n"):
            if sep in raw:
                names = raw.split(sep)
                break
        else:
            names = [raw] if raw else []
        # Deduplicate while preserving order
        seen = set()
        out = []
        for n in names:
            n = n.strip()
            if n and n not in seen:
                seen.add(n)
                out.append(n)
        return out

    @property
    def fields(self) -> "Fields":
        """
        누름틀(필드) 컬렉션 — list-like + dict-like collection accessor.

        v0.0.12+ : ``app.fields`` 가 ``Fields`` 컬렉션을 반환. iteration /
        ``in`` / ``len()`` 은 기존과 동일하게 필드 이름으로 동작
        (하위 호환). 추가로 dict-style + collection 메소드 지원.

        Examples
        --------
        Legacy 사용 (그대로 작동):

        >>> for name in app.fields:        # 필드 이름 iteration
        ...     print(name)
        >>> "customer" in app.fields       # 존재 확인
        >>> len(app.fields)                 # 개수
        >>> list(app.fields)                # ['customer', 'order_date', ...]

        v1.0 패턴 (dict-style):

        >>> app.fields["customer"] = "홍길동"        # 값 주입 (필요시 자동 생성)
        >>> app.fields["customer"].value             # 현재 값
        >>> app.fields.add("date", direction="날짜") # 명시적 생성
        >>> app.fields.update({"a": "1", "b": "2"})  # 일괄 주입
        >>> app.fields.remove("old")                 # 삭제
        >>> app.fields.from_brackets()                # {{tag}} → 필드 변환
        """
        return Fields(self)

    @property
    def fields_dict(self) -> dict:
        """
        ``{필드이름: 값}`` 딕셔너리.

        Examples
        --------
        >>> app.fields_dict
        {'customer_name': '홍길동', 'order_date': '2026-04-15', 'total': '1,200,000원'}
        """
        return {name: self.get_field(name) for name in self.fields}

    def field_exists(self, name: str) -> bool:
        """해당 이름의 필드가 존재하는지."""
        try:
            return bool(self.api.FieldExist(name))
        except Exception:
            return name in self.fields

    def move_to_field(self, name: str,
                      text: bool = True, front: bool = True,
                      select: bool = False) -> bool:
        """
        지정 필드로 커서 이동.

        Parameters
        ----------
        name : str
            필드 이름.
        text : bool
            필드 안의 텍스트 위치로 이동할지 (기본 True).
        front : bool
            필드 앞쪽으로 이동할지 (False 면 뒤쪽).
        select : bool
            필드 범위를 선택할지.
        """
        try:
            return bool(self.api.MoveToField(name, text, front, select))
        except Exception:
            return False

    def delete_field(self, name: str) -> bool:
        """이름으로 지정한 필드 삭제."""
        if not self.move_to_field(name, select=True):
            return False
        try:
            return bool(self.api.Run("DeleteField"))
        except Exception:
            return bool(self.api.Run("Delete"))

    def delete_all_fields(self) -> int:
        """모든 필드 삭제. 삭제한 개수 반환."""
        count = 0
        for name in self.fields:
            if self.delete_field(name):
                count += 1
        return count

    def rename_field(self, old: str, new: str) -> bool:
        """필드 이름 변경."""
        try:
            return bool(self.api.RenameField(old, new))
        except Exception:
            return False

    def replace_brackets_with_fields(
        self,
        pattern: str = r"\{\{(\w+)\}\}",
        memo: str = "",
    ) -> list:
        """
        ``{{name}}`` 형태의 브래킷 표기를 모두 HWP 필드로 변환.

        템플릿 문서를 mail merge 대상으로 준비하는 가장 간편한 방법.

        Parameters
        ----------
        pattern : str
            브래킷 감지 정규식. ``(\\w+)`` 캡처 그룹이 필드명이 됨.
        memo : str
            각 필드에 붙을 메모.

        Returns
        -------
        list[str]
            변환된 고유 필드 이름 목록.

        Examples
        --------
        문서에 ``안녕하세요 {{name}}님, 오늘은 {{date}} 입니다.`` 가 있으면:

        >>> names = app.replace_brackets_with_fields()
        >>> names
        ['name', 'date']
        >>> app.set_field("name", "홍길동")
        >>> app.set_field("date", "2026-04-15")
        """
        import re
        pat = re.compile(pattern)
        converted = []

        # Iterate: find next match, replace it with a field.
        # We restart from top_of_file each iteration because inserting a
        # field shifts positions.
        while True:
            text = self.text
            m = pat.search(text)
            if not m:
                break
            name = m.group(1)
            bracket = m.group(0)
            self.move.top_of_file()
            if not self.find_text(bracket):
                break
            # Selection is on the bracket — delete + create field
            self.api.Run("Delete")
            self.create_field(name, memo=memo)
            if name not in converted:
                converted.append(name)
        return converted

    # ═════════════════════════════════════════════════════════════
    # Pandas integration — Phase B
    # ═════════════════════════════════════════════════════════════

    def read_table(self, to: str = "df",
                   max_rows: int = 500, max_cols: int = 50):
        """
        커서가 들어있는 표를 파이썬 자료 구조로 추출.

        Parameters
        ----------
        to : str
            ``"df"`` (기본, pandas DataFrame) · ``"list"`` (2D list) ·
            ``"csv"`` (CSV 문자열).
        max_rows, max_cols : int
            안전 상한 — 표가 이보다 크면 해당 값까지만 읽고 중단.

        Returns
        -------
        pandas.DataFrame | list[list[str]] | str

        Examples
        --------
        >>> df = app.read_table()
        >>> df.to_csv("out.csv")
        """
        # A1 (표 좌상단) 셀로 이동. 실패하면 빈 결과 반환.
        try:
            self.api.Run("TableColBegin")
            self.api.Run("TableColPageUp")
        except Exception:
            if to == "list":
                return []
            if to == "csv":
                return ""
            try:
                import pandas as pd
                return pd.DataFrame()
            except ImportError:
                return []

        # 전체 셀을 `TableRightCell` 로 순회하면서 GetCellAddr 로 (row, col)
        # 측정. 돌아가거나 이전 셀과 같으면 종료.
        import re
        def _cell_addr():
            try:
                addr = self.api.GetCellAddr() or ""
                m = re.match(r"([A-Z]+)(\d+)", str(addr))
                if m:
                    col_letters, row_num = m.groups()
                    col = 0
                    for ch in col_letters:
                        col = col * 26 + (ord(ch) - ord("A") + 1)
                    return int(row_num) - 1, col - 1
            except Exception:
                pass
            return None, None

        cell_dict = {}  # {(row, col): text}
        visited = set()
        max_steps = max_rows * max_cols
        for _ in range(max_steps):
            r, c = _cell_addr()
            if r is None or (r, c) in visited:
                break
            visited.add((r, c))
            try:
                self.api.Run("TableCellBlock")
                txt = self.get_selected_text() or ""
                self.api.Run("Cancel")
            except Exception:
                txt = ""
            cell_dict[(r, c)] = str(txt).strip()
            prev = self.api.GetPos()
            try:
                self.api.Run("TableRightCell")
            except Exception:
                break
            if self.api.GetPos() == prev:
                break

        # cell_dict → rows (list of lists in row-major)
        if not cell_dict:
            rows = []
        else:
            max_r = max(r for r, c in cell_dict)
            max_c = max(c for r, c in cell_dict)
            rows = []
            for r in range(max_r + 1):
                row = []
                for c in range(max_c + 1):
                    row.append(cell_dict.get((r, c), ""))
                rows.append(row)

        if to == "list":
            return rows
        if to == "csv":
            import csv, io
            buf = io.StringIO()
            csv.writer(buf).writerows(rows)
            return buf.getvalue()
        # default "df"
        try:
            import pandas as pd
            if not rows:
                return pd.DataFrame()
            headers = rows[0]
            data = rows[1:] if len(rows) > 1 else []
            return pd.DataFrame(data, columns=headers)
        except ImportError:
            return rows

    def reload(self, new_app=False, dll_path=None):
        """
        새로운 HWPFrame.HwpObject로 `App` 인스턴스를 다시 로드하고 가시성과 DLL 경로를 재설정합니다.

        이 함수는 `app` 객체의 API를 새로운 HWPFrame.HwpObject 인스턴스로 다시 초기화하도록 설계되었습니다.
        또한 제공된 경우 지정된 DLL 파일을 확인하고 등록합니다.

        매개변수
        ----------
        dll_path : str, optional
            확인하고 등록할 DLL 파일의 경로. None인 경우 DLL 등록이 수행되지 않습니다.

        주의사항
        -----
        `reload` 함수는 HWPFrame.HwpObject의 상태를 재설정해야 하거나
        사용 중인 DLL을 변경할 때 유용합니다.

        사용 예시
        --------
        >>> app = App()
        >>> app.reload(dll_path="path/to/dll")
        """
        logger = get_logger('core')
        logger.debug(f"Calling reload")
        self._load(new_app=new_app, dll_path=dll_path)

    def get_filepath(self):
        """
        현재 열려있는 한컴오피스 Hwp 문서의 파일 경로를 검색합니다.

        이 함수는 `App` 인스턴스와 연결된 한컴오피스 Hwp 프로그램의 활성 문서에 접근하여
        전체 파일 경로를 반환합니다.

        반환값
        -------
        str
            현재 활성화된 한컴오피스 Hwp 문서의 전체 파일 경로.

        사용 예시
        --------
        >>> app = App()
        >>> filepath = app.get_filepath()
        >>> print(filepath)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_filepath")
        doc = self.api.XHwpDocuments.Active_XHwpDocument
        return doc.FullName

    def create_action(self, action_key: str):
        """
        `_Action` 클래스의 인스턴스를 생성하고 반환합니다.

        이 함수는 주어진 `action_key`와 연결된 새로운 `_Action` 인스턴스를 생성합니다.
        `action_key`는 애플리케이션에 대해 생성할 특정 액션의 유형을 지정합니다.

        매개변수
        ----------
        action_key : str
            생성할 특정 액션을 나타내는 키.

        반환값
        -------
        _Action
            제공된 애플리케이션 객체와 액션 키로 초기화된 `_Action` 클래스의 인스턴스.

        사용 예시
        --------
        >>> app = App()
        >>> action = app.create_action('some_action_key')
        >>> print(action)
        """
        logger = get_logger('core')
        logger.debug(f"Calling create_action")
        return _Action(self, action_key)

    def open(self, path: str):
        """
        제공된 파일 경로를 사용하여 한컴오피스 Hwp 프로그램에서 파일을 엽니다.

        이 함수는 먼저 제공된 파일 경로를 `get_absolute_path` 함수를 사용하여 절대 경로로 변환합니다.
        그런 다음 `api.Open` 메서드를 사용하여 한컴오피스 Hwp 프로그램에서 파일을 열고 열린 파일의 절대 경로를 반환합니다.

        매개변수
        ----------
        path : str
            열릴 문서의 파일 경로.

        반환값
        -------
        str
            열린 파일의 절대 경로.

        사용 예시
        --------
        >>> app = App()
        >>> opened_file_path = app.open('path/to/document.hwp')
        >>> print(opened_file_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling open")
        name = get_absolute_path(path)
        self.api.Open(name)
        return name

    def get_hwnd(self):
        """
        Retrieves the window handle (HWND) of the active window in the Hancom Office Hwp program.

        This function accesses the active window in the Hancom Office Hwp program linked to the `App` instance
        and returns its window handle. The window handle can be used in scenarios where direct manipulation
        or interaction with the window at the OS level is required.

        Returns
        -------
        int
            The window handle (HWND) of the active Hancom Office Hwp window.

        Examples
        --------
        >>> app = App()
        >>> hwnd = app.get_hwnd()
        >>> print(hwnd)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_hwnd")
        return self.api.XHwpWindows.Active_XHwpWindow.WindowHandle

    def save(self, path=None):
        """
        Saves the current document in the Hancom Office Hwp program.

        If a path is provided, the document is saved to that location in the appropriate format based on the file extension.
        Supported formats include HWP, PDF, HWPML2X (HWPX), and PNG. If no path is provided, the document is saved
        in its current location.

        Parameters
        ----------
        path : str, optional
            The file path where the document will be saved. If None, the document is saved in its current location.

        Returns
        -------
        str
            The file path where the document has been saved.

        Examples
        --------
        >>> app = App()
        >>> saved_file_path = app.save('path/to/document.hwp')
        >>> print(saved_file_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling save")
        if not path:
            self.api.Save()
            return self.get_filepath()

        name = get_absolute_path(path)
        extension = Path(name).suffix
        format_ = {
            ".hwp": "HWP",
            ".pdf": "PDF",
            ".hwpx": "HWPX",
            ".hml": "HWPML2X",
            ".png": "PNG",
            ".txt": "TEXT",
            ".docx": "MSWORD",
        }.get(extension)

        self.api.SaveAs(name, format_)
        return name

    def save_block(self, path: Path):
        """
        Saves a block of content in the Hancom Office Hwp program and returns the file path.

        This function saves a specified block from the Hancom Office Hwp program to a file in a given format.
        The format is determined by the file extension. Supported formats include HWP, PDF, HWPML2X (HWPX), and PNG.
        It returns the file path if the save operation is successful or None if it fails.

        Parameters
        ----------
        path : Path
            The file path where the block will be saved.

        Returns
        -------
        Path or None
            The path to the saved file if the save operation is successful, or None if it fails.

        Examples
        --------
        >>> app = App()
        >>> saved_path = app.save_block(Path('path/to/save/block.hwp'))
        >>> print(saved_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling save_block")

        name = get_absolute_path(path)
        extension = Path(name).suffix
        format_ = {
            ".hwp": "HWP",
            ".pdf": "PDF",
            ".hwpx": "HWPML2X",
            ".png": "PNG",
        }.get(extension)

        action = self.actions.SaveBlockAction
        p = action.pset

        p.filename = name
        p.Format = format_
        action.run(p)
        return name if Path(name).exists() else None

    def close(self):
        """
        Closes the currently open document in the Hancom Office Hwp program.

        This function triggers the 'FileClose' command within the Hancom Office Hwp program to close the current document.
        It's useful for programmatically managing documents within the application.

        Examples
        --------
        >>> app = App()
        >>> # Open and manipulate the document
        >>> app.close()
        """
        logger = get_logger('core')
        logger.debug(f"Calling close")
        self.api.Run("FileClose")

    def quit(self):
        """
        Terminates the Hancom Office Hwp program instance associated with the `App`.

        This function invokes the 'FileQuit' command within the Hancom Office Hwp program to close the application.
        It's useful for programmatically controlling the lifecycle of the application.

        Examples
        --------
        >>> app = App()
        >>> # Perform actions with the app
        >>> app.quit()
        """
        logger = get_logger('core')
        logger.debug(f"Calling quit")
        self.api.Run("FileQuit")

    def get_font_list(self):
        """
        Retrieves the list of fonts that used in the documnents from the current application.

        This method accesses the font list in the Hancom Office Hwp program linked to the `App` instance.
        It extracts fonts that used and return in list for further use or analysis.

        Returns
        -------
        List
            return font list that used in the hwp document.

        Examples
        --------
        >>> app = App()
        >>> font_list = app.get_font_list()
        >>> print(font_list)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_font_list")
        self.api.ScanFont()
        result = self.api.GetFontList()
        output = list(map(lambda txt: txt.split(",")[0],
            result.split("\x02")
            ))
        return output

    def charshape(self,
        facename=None,
        fonttype=None,
        size=None,
        ratio=None,
        spacing=None,
        offset=None,
        bold=None,
        italic=None,
        small_caps=None,
        emboss=None,
        engrave=None,
        superscript=None,
        subscript=None,
        underline_type=None,
        underline_shape=None,
        outline_type=None,
        shadow_type=None,
        text_color=None,
        shade_color=None,
        underline_color=None,
        shadow_color=None,
        shadow_offset_x=None,
        shadow_offset_y=None,
        strikeout_color=None,
        strikeout_type=None,
        strikeout_shape=None,
        diac_sym_mark=None,
        use_font_space=None,
        use_kerning=None,
        height=None,
        border_fill=None):
        """
        .. deprecated:: 0.0.7
           Use :meth:`set_charshape` directly — it accepts all the same
           keyword arguments and applies the result immediately. For
           scoped formatting that auto-reverts on exit, use
           :meth:`styled_text` or :meth:`charshape_scope`.

           This method returns an **unbound** ``CharShape`` pset that
           the caller must pass back to ``set_charshape`` — an extra
           indirection that adds no value.

        Examples
        --------
        >>> # ❌ Old (deprecated)
        >>> cs = app.charshape(bold=True, italic=True)
        >>> app.set_charshape(cs)

        >>> # ✅ New
        >>> app.set_charshape(bold=True, italic=True)
        """
        import warnings
        warnings.warn(
            "App.charshape() is deprecated and will be removed in v0.1.0. "
            "Use App.set_charshape(**kwargs) directly — same arguments, "
            "applied immediately. For scoped formatting, use styled_text() "
            "or charshape_scope().",
            DeprecationWarning,
            stacklevel=2,
        )

        logger = get_logger('core')
        logger.debug(f"Calling charshape")

        charshape_pset = self.create_parameterset("CharShape")
        value_set = {
            "facename": facename,
            "fonttype": fonttype,
            "size": size,
            "ratio": ratio,
            "spacing": spacing,
            "offset": offset,
            "bold": bold,
            "italic": italic,
            "small_caps": small_caps,
            "emboss": emboss,
            "engrave": engrave,
            "superscript": superscript,
            "subscript": subscript,
            "underline_type": underline_type,
            "underline_shape": underline_shape,
            "outline_type": outline_type,
            "shadow_type": shadow_type,
            "text_color": text_color,
            "shade_color": shade_color,
            "underline_color": underline_color,
            "shadow_color": shadow_color,
            "shadow_offset_x": shadow_offset_x,
            "shadow_offset_y": shadow_offset_y,
            "strikeout_color": strikeout_color,
            "strikeout_type": strikeout_type,
            "strikeout_shape": strikeout_shape,
            "diac_sym_mark": diac_sym_mark,
            "use_font_space": use_font_space,
            "use_kerning": use_kerning,
            "height": height,
            "border_fil": border_fill}
        for key, value in value_set.items():
            if not key:
                continue
            setattr(charshape_pset, key, value)

        return charshape_pset

    @property
    def charshape(self):
        """
        현재 커서 위치의 문자 모양 (CharShape) — read/write property.

        v0.0.12+ 권장 패턴 (v1.0 청사진). 기존 :meth:`get_charshape` /
        :meth:`set_charshape` 와 공존.

        Examples
        --------
        Read (snapshot):

        >>> cs = app.charshape
        >>> cs.bold, cs.fontsize
        (True, 1100)

        Write (전체 교체):

        >>> new_cs = parametersets.CharShape()
        >>> new_cs.bold = True
        >>> new_cs.fontsize = 1500
        >>> app.charshape = new_cs

        Partial update (kwargs) — :meth:`set_charshape` 가 더 간결:

        >>> app.set_charshape(bold=True, fontsize=1500)
        """
        return self.get_charshape()

    @charshape.setter
    def charshape(self, value) -> None:
        if value is None:
            return
        if isinstance(value, dict):
            self.set_charshape(**value)
        else:
            self.set_charshape(charshape=value)

    @property
    def parashape(self):
        """
        현재 커서 위치의 문단 모양 (ParaShape) — read/write property.

        v0.0.12+ 권장 패턴 (v1.0 청사진). 기존 :meth:`get_parashape` /
        :meth:`set_parashape` 와 공존.

        Examples
        --------
        >>> ps = app.parashape          # 현재 스냅샷
        >>> app.parashape = new_ps      # 전체 교체
        >>> app.set_parashape(align='Left', indent=10)  # partial
        """
        return self.get_parashape()

    @parashape.setter
    def parashape(self, value) -> None:
        if value is None:
            return
        if isinstance(value, dict):
            self.set_parashape(**value)
        else:
            self.set_parashape(parashape=value)

    def get_charshape(self):
        """
        Retrieves the character shape settings from the current application.

        This method accesses the character shape settings in the Hancom Office Hwp program linked to the `App` instance.
        It extracts these settings and encapsulates them in a `CharShape` object for further use or analysis.

        Returns
        -------
        CharShape
            An object representing the character shape settings of the application.

        Examples
        --------
        >>> app = App()
        >>> char_shape = app.get_charshape()
        >>> print(char_shape)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_charshape")
        return self.actions.CharShape.pset

    def set_charshape(self, charshape: parametersets.CharShape=None, **kwargs):
        """
        Sets the character shape in the current paragraph of the Hancom Office Hwp application using the provided `CharShape`.

        If `charshape` is None, a default instance of `CharShape` is created. The function can also accept additional keyword
        arguments (`kwargs`), which are assigned as properties of `charshape`. This modified `charshape` is then used to change
        the current paragraph shape in the Hancom Office Hwp document.

        Parameters
        ----------
        charshape : CharShape, optional
            The `CharShape` object to be used for setting character shapes in Hancom Office Hwp. The default is None.
        **kwargs
            Additional keyword arguments that are assigned as properties of `charshape`.

        Returns
        -------
        bool
            A boolean value indicating the success of the `set_charshape` operation.

        Examples
        --------
        >>> app = App()
        >>> char_shape = app.get_charshape()
        >>> success = app.set_charshape(charshape=char_shape, fontName='Arial', fontSize=10)
        >>> print(success)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_charshape")
        action = self.actions.CharShape
        pset = action.pset
        # charshape를 전달하면 반영해야 함
        if charshape:
            pset.update_from(charshape)

        for key, value in kwargs.items():
            setattr(pset, key, value)

        if action.run():
            return pset
        return None

    def get_parashape(self):
        """
        Retrieves the paragraph shape settings from the current application.

        This method accesses the paragraph shape settings in the Hancom Office Hwp program linked to the `App` instance.
        It extracts these settings and encapsulates them in a `ParaShape` object for further use or analysis.

        Returns
        -------
        ParaShape
            An object representing the paragraph shape settings of the application.

        Examples
        --------
        >>> app = App()
        >>> para_shape = app.get_parashape()
        >>> print(para_shape)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_parashape")
        return self.actions.ParagraphShape.pset

    def set_parashape(self, parashape: ParaShape = None, **kwargs):
        """
        Sets the paragraph shape in the Hancom Office Hwp application using the provided `ParaShape`.

        If `parashape` is None, a default instance of `ParaShape` is created. The function can also accept additional keyword
        arguments (`kwargs`), which are assigned as properties of `parashape`. This modified `parashape` is then used to change
        the current paragraph shape in the Hancom Office Hwp document.

        Parameters
        ----------
        parashape : ParaShape, optional
            The `ParaShape` object to be used for setting paragraph shapes in Hancom Office Hwp. Defaults to `None`.
        **kwargs
            Additional keyword arguments assigned as properties of `parashape`.
        if parashape:
            pset.update(parashape)

        bool
            A boolean value indicating the success of the `set_parashape` operation.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_parashape(parashape=para_shape, align='Left', indent=10)
        >>> print(success)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_parashape")
        action = self.actions.ParagraphShape
        pset = action.pset
        if parashape:
            pset.update_from(parashape)

        for key, value in kwargs.items():
            setattr(pset, key, value)

        return action.run()

    def insert_text(
        self,
        text: str,
        charshape: parametersets.CharShape = None,
        **kwargs,
    ):
        """
        Inserts text into the Hancom Office Hwp document with specified character shape settings.

        This function inserts the given text into the document associated with the `App` instance.
        It allows for optional character shape settings through a `CharShape` object.
        Additional character shape attributes can be specified via keyword arguments.

        Parameters
        ----------
        text : str
            The text to be inserted into the document.
        charshape : CharShape, optional
            An optional `CharShape` object to specify the character shape settings for the inserted text. Defaults to `None`.
        **kwargs
            Additional character shape attributes to be set on the `charshape` object.

        Examples
        --------
        >>> app = App()
        >>> app.insert_text("Hello World", fontName="Arial", fontSize=12)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_text")

        if not charshape:
            charshape = self.get_charshape()
        for key, value in kwargs.items():
            setattr(charshape, key, value)
        self.set_charshape(charshape)

        # NOTE: `_Action` is now doc-aware (lazy-cache keyed by active doc
        # id), so we can safely use `self.actions.InsertText` again —
        # it automatically binds to the currently active document.
        action = self.actions.InsertText
        action.pset.Text = text
        action.run()
        return self  # Fluent — chain 가능 (v0.0.13+)

    def styled_text(self, text: str, **fmt):
        """
        포맷팅이 **그 텍스트에만** 적용되도록 삽입합니다.

        동작 방식 — **snapshot + restore**:

        1. 블록 진입 시 현재 CharShape / ParaShape 를 통째로 snapshot
        2. 텍스트를 snapshot 상태 그대로 삽입 (이전 상태 유지)
        3. 방금 삽입한 영역을 선택해서 요청한 서식만 적용
        4. 커서를 끝으로 이동한 뒤 snapshot 으로 상태 **완전 복원**

        이 방식은 ``shade_color`` 같은 "지울 수 없는" 속성에도 작동합니다.
        snapshot 에 저장된 원래 값으로 돌려놓기만 하면 되기 때문입니다.

        Parameters
        ----------
        text : str
            삽입할 텍스트.
        **fmt
            ``set_charshape()`` 가 받는 키워드 — bold, italic, underline_type,
            strike_out_type, super_script, sub_script, text_color, shade_color,
            height, spacing_hangul 등.

        Examples
        --------
        >>> app.insert_text("앞: ")
        >>> app.styled_text("중요한 부분", bold=True, text_color="#FF0000")
        >>> app.insert_text(" 뒤")
        # "앞: " 기본, "중요한 부분" 빨간 굵은 글씨, " 뒤" 기본

        >>> # shade_color (형광펜) 도 안전
        >>> app.styled_text("하이라이트", shade_color="#FFFF00")
        >>> app.insert_text(" ← 음영이 여기까지 이어지지 않습니다")
        """
        if not text:
            return self  # Fluent — chain 가능 (v0.0.13+)
        # 1) Snapshot CharShape BEFORE any changes
        char_before = self._snapshot_charshape()
        # 2) Insert text using the current (unchanged) cursor state
        self.insert_text(text)
        # 3) Select the just-inserted region backwards
        for _ in range(len(text)):
            self.api.Run("MoveSelLeft")
        # 4) Apply the requested formatting to the selection
        if fmt:
            self.set_charshape(**fmt)
        # 5) Deselect + move cursor to end of the inserted text
        try:
            self.api.Run("Cancel")
        except Exception:
            pass
        for _ in range(len(text)):
            self.api.Run("MoveRight")
        # 6) RESTORE the snapshot — cursor char state == pre-block state
        self._restore_charshape(char_before)
        return self  # Fluent — chain 가능 (v0.0.13+)

    def _snapshot_charshape(self):
        """Capture the current CharShape at cursor as a restorable snapshot."""
        try:
            cs_action = self.actions.CharShape
            cs_action.act.GetDefault(cs_action.pset._raw)
            return cs_action.pset.clone() if hasattr(cs_action.pset, 'clone') else None
        except Exception:
            return None

    def _restore_charshape(self, snapshot):
        """Apply a previously-captured CharShape snapshot at cursor."""
        if snapshot is None:
            return
        try:
            self.set_charshape(snapshot)
        except Exception:
            pass

    def _snapshot_parashape(self):
        """Capture the current ParaShape at cursor as a restorable snapshot."""
        try:
            ps_action = self.actions.ParagraphShape
            ps_action.act.GetDefault(ps_action.pset._raw)
            return ps_action.pset.clone() if hasattr(ps_action.pset, 'clone') else None
        except Exception:
            return None

    def _restore_parashape(self, snapshot):
        """Apply a previously-captured ParaShape snapshot at cursor."""
        if snapshot is None:
            return
        try:
            self.set_parashape(snapshot)
        except Exception:
            pass

    @contextmanager
    def charshape_scope(self, **fmt):
        """
        블록 내부에서 입력된 텍스트에 CharShape 서식을 **스코프** 로 적용.

        진입 시 현재 ParameterSet 을 snapshot 으로 저장하고, 블록 내용을
        선택해서 서식을 적용한 뒤, **snapshot 으로 완전 복원** 합니다.
        이렇게 하면 shade_color 같은 "기본값으로 되돌리기 어려운" 속성도
        원래 값으로 정확히 되돌아갑니다.

        Examples
        --------
        >>> with app.charshape_scope(bold=True, underline_type=1):
        ...     app.insert_text("중요 공지:")
        >>> app.insert_text(" 이후는 원래 서식 그대로")

        >>> with app.charshape_scope(shade_color="#FFFF00"):
        ...     app.insert_text("형광펜 여러 줄도\\n안전하게 처리")
        """
        before_pos = None
        try:
            before_pos = self.api.GetPos()
        except Exception:
            pass
        # Snapshot cursor char shape BEFORE any changes
        char_before = self._snapshot_charshape()

        try:
            yield
        finally:
            after_pos = None
            try:
                after_pos = self.api.GetPos()
            except Exception:
                pass

            # Apply fmt to the block content via selection
            if (before_pos and after_pos and before_pos != after_pos):
                try:
                    self.api.SetPos(*before_pos)
                    try:
                        self.api.SelectText(
                            before_pos[0], before_pos[1], before_pos[2],
                            after_pos[0],  after_pos[1],  after_pos[2],
                        )
                    except Exception:
                        self.api.SetPos(*before_pos)
                        self.api.Run("MoveSelDocEnd")
                    if fmt:
                        self.set_charshape(**fmt)
                    try:
                        self.api.Run("Cancel")
                    except Exception:
                        pass
                except Exception:
                    pass

            # Restore cursor state — this is the KEY step for clean exit
            if after_pos:
                try:
                    self.api.SetPos(*after_pos)
                except Exception:
                    pass
            self._restore_charshape(char_before)

    @contextmanager
    def parashape_scope(self, **fmt):
        """
        블록 내부 문단에 ParaShape 서식을 **스코프** 로 적용.

        정렬, 줄간격, 들여쓰기 등 문단 레벨 서식을 블록 내에서만 적용하고
        블록 종료 시 원래 ParaShape 로 **완전 복원** 합니다.

        Examples
        --------
        >>> with app.parashape_scope(align_type=3, line_spacing=200):
        ...     app.insert_text("가운데 정렬 + 200% 줄간격\\n여러 줄도 OK\\n")
        >>> app.insert_text("이후 문단은 원래 정렬/줄간격 그대로")
        """
        before_pos = None
        try:
            before_pos = self.api.GetPos()
        except Exception:
            pass
        para_before = self._snapshot_parashape()

        try:
            yield
        finally:
            after_pos = None
            try:
                after_pos = self.api.GetPos()
            except Exception:
                pass

            if (before_pos and after_pos and before_pos != after_pos):
                try:
                    self.api.SetPos(*before_pos)
                    try:
                        self.api.SelectText(
                            before_pos[0], before_pos[1], before_pos[2],
                            after_pos[0],  after_pos[1],  after_pos[2],
                        )
                    except Exception:
                        self.api.SetPos(*before_pos)
                        self.api.Run("MoveSelDocEnd")
                    if fmt:
                        self.set_parashape(**fmt)
                    try:
                        self.api.Run("Cancel")
                    except Exception:
                        pass
                except Exception:
                    pass

            if after_pos:
                try:
                    self.api.SetPos(*after_pos)
                except Exception:
                    pass
            self._restore_parashape(para_before)

    @contextmanager
    def use_document(self, doc_or_index):
        """
        지정한 문서를 **임시로 활성창** 으로 만들고, 블록 종료 시 원래 활성
        문서로 복원하는 context manager.

        다중 문서 작업에서 각 문서로 잠깐씩 전환하며 작업할 때 유용합니다.

        Parameters
        ----------
        doc_or_index : Document | int
            ``Document`` 인스턴스 또는 ``app.documents`` 의 인덱스.

        Examples
        --------
        >>> # 인덱스로 전환
        >>> with app.use_document(0):
        ...     app.insert_text("첫 번째 문서에 추가")

        >>> # Document 객체로 전환
        >>> doc2 = app.documents.open("other.hwp")
        >>> with app.use_document(doc2):
        ...     app.replace_all("OLD", "NEW")

        >>> # 여러 문서 순회하며 같은 작업 반복
        >>> for doc in app.documents:
        ...     with app.use_document(doc):
        ...         app.replace_all("2025", "2026")
        ...         doc.save()
        """
        # Remember which document was active before
        prev_active = self.documents.active

        # Resolve target
        if isinstance(doc_or_index, Document):
            target = doc_or_index
        else:
            target = self.documents[int(doc_or_index)]

        target.activate()
        try:
            yield target
        finally:
            if prev_active is not None:
                try:
                    prev_active.activate()
                except Exception:
                    pass

    @contextmanager
    def scan(
        self,
        option=const.MaskOption.All,
        selection=False,
        scan_spos=const.ScanStartPosition.Current,
        scan_epos=const.ScanEndPosition.Document,
        spara=None,
        spos=None,
        epara=None,
        epos=None,
        scan_direction=const.ScanDirection.Forward,
    ):


        logger = get_logger('core')
        logger.debug(f"Calling scan")

        range_ = scan_spos.value + scan_epos.value
        if selection:
            range_ = 0x00FF  # Limit the scanning to the block if selection is True

        range_ += scan_direction.value
        self.api.InitScan(
            option=option.value,
            Range=range_,
            spara=spara,
            spos=spos,
            epara=epara,
            epos=epos,
        )
        yield _get_text(self)
        self.api.ReleaseScan()

    def setup_page(
        self,
        top=20,  # 위쪽 여백 (밀리미터 단위)
        bottom=10,  # 아래쪽 여백 (밀리미터 단위)
        right=20,  # 오른쪽 여백 (밀리미터 단위)
        left=20,  # 왼쪽 여백 (밀리미터 단위)
        header=15,  # 머리글 길이 (밀리미터 단위)
        footer=5,  # 바닥글 길이 (밀리미터 단위)
        gutter=0,  # 제본 여백 (밀리미터 단위)
    ):
        """
        Hancom Office Hwp 문서의 페이지 레이아웃을 설정합니다.

        이 함수는 페이지 여백, 머리글, 바닥글, 제본 여백의 크기를 설정합니다.
        단위는 밀리미터이며, 애플리케이션 내부의 단위 시스템으로 변환됩니다.

        매개변수
        ----------
        top : int, optional
            위쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        bottom : int, optional
            아래쪽 여백 크기 (밀리미터 단위). 기본값은 10mm.
        right : int, optional
            오른쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        left : int, optional
            왼쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        header : int, optional
            머리글 길이 (밀리미터 단위). 기본값은 15mm.
        footer : int, optional
            바닥글 길이 (밀리미터 단위). 기본값은 5mm.
        gutter : int, optional
            제본 여백 크기 (밀리미터 단위). 기본값은 0mm.

        반환값
        -------
        bool
            페이지 설정이 성공하면 True, 실패하면 False.

        사용 예시
        --------
        >>> app = App()
        >>> app.setup_page(top=25, bottom=15, right=20, left=20, header=10, footer=5, gutter=5)
        """


        logger = get_logger('core')
        logger.debug(f"Calling setup_page")

        action = self.actions.PageSetup
        p = action.pset

        # 각 여백 및 길이를 밀리미터 단위에서 한컴오피스 내부 단위로 변환하여 설정
        p.PageDef.TopMargin = self.api.MiliToHwpUnit(top)
        p.PageDef.HeaderLen = self.api.MiliToHwpUnit(header)
        p.PageDef.RightMargin = self.api.MiliToHwpUnit(right)
        p.PageDef.BottomMargin = self.api.MiliToHwpUnit(bottom)
        p.PageDef.FooterLen = self.api.MiliToHwpUnit(footer)
        p.PageDef.LeftMargin = self.api.MiliToHwpUnit(left)
        p.PageDef.GutterLen = self.api.MiliToHwpUnit(gutter)

        return action.run()  # 페이지 설정 실행

    def insert_picture(
        self,
        fpath,
        width=None,
        height=None,
        size_option=const.SizeOption.RealSize,
        reverse=False,
        watermark=False,
        effect=const.Effect.RealPicture,
    ):
        """
        지정된 크기와 효과 옵션으로 문서에 이미지를 삽입합니다.

        매개변수
        ----------
        fpath : str
            삽입할 이미지 파일 경로.
        width : int, optional
            이미지의 너비. None이면 크기 옵션에 따라 결정됩니다. 기본값은 None.
        height : int, optional
            이미지의 높이. None이면 크기 옵션에 따라 결정됩니다. 기본값은 None.
        size_option : SizeOption, optional
            이미지의 크기 옵션으로 `SizeOption` Enum에 정의되어 있습니다. 기본값은 SizeOption.RealSize.
        reverse : bool, optional
            True이면 이미지를 반전합니다. 기본값은 False.
        watermark : bool, optional
            True이면 이미지를 워터마크로 처리합니다. 기본값은 False.
        effect : Effect, optional
            이미지의 시각적 효과로 `Effect` Enum에 정의되어 있습니다. 기본값은 Effect.RealPicture.

        반환값
        -------
        bool
            이미지가 성공적으로 삽입되었으면 True, 그렇지 않으면 False.

        사용 예시
        --------
        >>> app = App()
        >>> success = app.insert_picture('path/to/image.jpg', width=100, height=200, size_option=SizeOption.SpecificSize, effect=Effect.GrayScale)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_picture")

        path = Path(fpath)  # 이미지 파일 경로 처리
        size_option = size_option.value  # 크기 옵션 값 설정
        effect = effect.value  # 효과 옵션 값 설정

        return self.api.InsertPicture(
            path.absolute().as_posix(),  # 파일의 절대 경로
            Width=width,  # 이미지 너비
            Height=height,  # 이미지 높이
            sizeoption=size_option,  # 크기 옵션
            reverse=reverse,  # 반전 여부
            watermark=watermark,  # 워터마크 여부
            effect=effect,  # 시각적 효과
        )

    def select_text(self, option=const.SelectionOption.Line):
        """
        지정된 옵션에 따라 문서에서 텍스트를 선택합니다.

        매개변수
        ----------
        option : SelectionOption, optional
            선택할 텍스트 단위. SelectionOption Enum에 정의된 옵션을 사용. 기본값은 SelectionOption.Line.

        반환값
        -------
        tuple
            시작 및 끝 이동 작업의 결과를 포함하는 튜플 (둘 다 boolean 값).

        사용 예시
        --------
        >>> app = App()
        >>> app.select_text(option=SelectionOption.Para)
        """

        logger = get_logger('core')
        logger.debug(f"Calling select_text with option={option}")

        # 문자열 입력 처리
        if isinstance(option, str):
            try:
                option = const.SelectionOption[option]  # 문자열을 Enum으로 변환
            except KeyError:
                raise ValueError(f"Invalid option string: {option}. Must be one of {[o.name for o in const.SelectionOption]}")

        # Enum 멤버라면 값 꺼내기
        begin_action_name, end_action_name = option.value
        begin_action = getattr(self.actions, begin_action_name)
        end_action = getattr(self.actions, end_action_name)

        return begin_action(), end_action()

    def get_selected_text(self):
        """
        Hancom Office Hwp 문서에서 현재 선택된 영역의 텍스트를 가져옵니다.

        이 함수는 문서에서 선택된 텍스트를 스캔하여 문자열로 반환합니다.
        현재 강조 표시되거나 선택된 텍스트를 처리하는 작업에 특히 유용합니다.

        반환값
        -------
        str
            문서에서 현재 선택된 영역의 텍스트.

        사용 예시
        --------
        >>> app = App()
        >>> selected_text = app.get_selected_text()
        >>> print(selected_text)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_selected_text")

        with self.scan(selection=True) as scan:  # 선택된 영역 스캔
            text = "\n".join(scan)  # 선택된 텍스트를 줄 단위로 결합
        return text  # 결합된 텍스트 반환

    def get_text(self, spos=const.ScanStartPosition.Line, epos=const.ScanEndPosition.Line):
        """
        지정된 시작 및 끝 위치에 따라 문서에서 텍스트를 가져옵니다.

        매개변수
        ----------
        spos : ScanStartPosition, optional
            텍스트 검색을 시작할 위치. 기본값은 ScanStartPosition.Line.
        epos : ScanEndPosition, optional
            텍스트 검색을 종료할 위치. 기본값은 ScanEndPosition.Line.

        반환값
        -------
        str
            지정된 시작 위치부터 끝 위치까지 가져온 텍스트.

        사용 예시
        --------
        >>> app = App()
        >>> text = app.get_text(spos=ScanStartPosition.Paragraph, epos=ScanEndPosition.Paragraph)
        >>> print(text)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_text")

        with self.scan(scan_spos=spos, scan_epos=epos) as txts:  # 지정된 위치 범위에서 텍스트 스캔
            text = "".join(txts)  # 스캔한 텍스트를 하나의 문자열로 결합
        return text  # 결합된 텍스트 반환

    def find_text(
        self,
        text="",  # 찾을 텍스트
        ignore_message=True,  # 메시지 무시 여부
        direction=0,  # 검색 방향
        match_case=None,  # 대소문자 구분
        all_word_forms=None,  # 모든 단어 형태 검색
        several_words=None,  # 여러 단어 검색
        use_wild_cards=None,  # 와일드카드 사용
        whole_word_only=None,  # 전체 단어만 검색
        replace_mode=None,  # 찾아 바꾸기 모드
        ignore_find_string=None,  # 찾을 문자열 무시
        ignore_replace_string=None,  # 바꿀 문자열 무시
        find_style="",  # 찾을 스타일
        replace_style="",  # 바꿀 스타일
        find_jaso=None,  # 자소로 검색
        find_regexp=None,  # 정규표현식으로 검색
        find_type=True,  # 마지막 검색(True), 새 검색(False)
        facename=None,
        text_color=None,
        bold = None,
        charshape: parametersets.CharShape = None,

    ):
        """
        문서에서 다양한 옵션을 사용해 특정 텍스트를 검색합니다.

        매개변수
        ----------
        text : str
            검색할 텍스트 문자열.
        [기타 매개변수...]

        반환값
        -------
        bool
            텍스트를 찾았으면 True, 그렇지 않으면 False.

        사용 예시
        --------
        >>> app = App()
        >>> found = app.find_text(text="Hello", direction=Direction.Forward)
        >>> print(found)
        """


        logger = get_logger('core')
        logger.debug(f"Calling find_text")

        # 반복 검색 액션 생성
        action = self.actions.RepeatFind
        pset = action.pset
        # pset.FindCharShape.update_from(self.charshape())

        if charshape:
            pset.FindCharShape
            pset.FindCharShape.update_from(charshape)

        # 옵션 설정
        if text is not None:
            pset.FindString = text
        if ignore_message is not None:
            pset.ignore_message = ignore_message
        if match_case is not None:
            pset.mach_case = match_case
        if all_word_forms is not None:
            pset.all_word_forms = all_word_forms
        if direction is not None:
            pset.direction = direction
        if several_words is not None:
            pset.several_words = several_words
        if use_wild_cards is not None:
            pset.use_wild_cards = use_wild_cards
        if whole_word_only is not None:
            pset.whole_word_only = whole_word_only
        if replace_mode is not None:
            pset.replace_mode = replace_mode
        if ignore_find_string is not None:
            pset.IgnoreFindString = ignore_find_string
        if ignore_replace_string is not None:
            pset.IgnoreReplaceString = ignore_replace_string
        if find_style is not None:
            pset.FindStyle = find_style
        if replace_style is not None:
            pset.ReplaceStyle = replace_style
        if find_jaso is not None:
            pset.FindJaso = find_jaso
        if find_regexp is not None:
            pset.FindRegexp = find_regexp
        if find_type is not None:
            pset.FindType = find_type
        if facename is not None:
            pset.FindCharShape
            pset.FindCharShape.FaceNameHangul = facename
            from hwpapi.constants import korean_fonts, english_fonts
            fonts = set(korean_fonts + english_fonts)
            pset.FindCharShape.FontTypeHangul = 1 if facename not in fonts else 2
        if text_color is not None:
            pset.FindCharShape.TextColor = text_color
        if bold is not None:
            pset.FindCharShape.Bold = bold

        return action.run()

    def replace_all(
        self,
        old_text="",
        new_text="",
        old_charshape=None,  # 기존 문자 모양
        new_charshape=None,  # 새로운 문자 모양
        ignore_message=True,  # 메시지 무시 여부
        direction=None,  # 검색 방향
        match_case=None,  # 대소문자 구분
        all_word_forms=None,  # 모든 단어 형태
        several_words=None,  # 여러 단어 검색
        use_wild_cards=None,  # 와일드카드 사용
        whole_word_only=None,  # 전체 단어만 검색
        auto_spell=None,  # 자동 철자 교정
        replace_mode=None,  # 찾아 바꾸기 모드
        ignore_find_string=None,  # 찾을 문자열 무시
        ignore_replace_string=None,  # 바꿀 문자열 무시
        find_regexp=None,  # 정규표현식으로 검색
        find_style=None,  # 찾을 스타일
        replace_style=None,  # 바꿀 스타일
        find_jaso=None,  # 자소로 검색
        find_reg_exp=None,  # 정규표현식으로 검색
        find_type=None,  # 마지막 검색 사용(True), 새 검색(False)
    ):
        """
        문서에서 특정 텍스트를 새 텍스트로 모두 교체합니다.

        매개변수
        ----------
        old_text : str
            교체할 텍스트 문자열.
        new_text : str
            대체할 새 텍스트 문자열.
        [기타 매개변수...]

        반환값
        -------
        bool
            교체 작업이 성공하면 True, 실패하면 False.

        사용 예시
        --------
        >>> app = App()
        >>> success = app.replace_all(old_text="Hello", new_text="Hi", direction=Direction.All)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling replace_all")

        # 반복 검색 액션 생성
        action = self.actions.AllReplace
        pset = action.pset

        # 옵션 설정
        if old_text is not None:
            pset.FindString = old_text
        if new_text is not None:
            pset.ReplaceString = new_text
        if ignore_message is not None:
            pset.ignore_message = ignore_message
        if match_case is not None:
            pset.mach_case = match_case
        if all_word_forms is not None:
            pset.all_word_forms = all_word_forms
        if direction is not None:
            pset.direction = direction
        if several_words is not None:
            pset.several_words = several_words
        if use_wild_cards is not None:
            pset.use_wild_cards = use_wild_cards
        if whole_word_only is not None:
            pset.whole_word_only = whole_word_only
        if replace_mode is not None:
            pset.replace_mode = replace_mode
        if ignore_find_string is not None:
            pset.IgnoreFindString = ignore_find_string
        if ignore_replace_string is not None:
            pset.IgnoreReplaceString = ignore_replace_string
        if find_style is not None:
            pset.FindStyle = find_style
        if replace_style is not None:
            pset.ReplaceStyle = replace_style
        if find_jaso is not None:
            pset.FindJaso = find_jaso
        if find_regexp is not None:
            pset.FindRegexp = find_regexp
        if find_type is not None:
            pset.FindType = find_type
        if auto_spell is not None:
            pset.auto_spell = auto_spell

        # 문자 모양 설정
        if old_charshape:
            pset.FindCharShape.update_from(old_charshape)
        if new_charshape:
            pset.replace_charshape.update_from(new_charshape)
        return action.run()

    def insert_file(
        self,
        fpath,
        keep_charshape=False,
        keep_parashape=False,
        keep_section=False,
        keep_style=False,
    ):
        """
        Inserts the contents of a specified file into the current document.

        This function inserts the contents of another file into the current document at the cursor's position.
        It provides options to retain various formatting attributes of the inserted content.

        Parameters
        ----------
        fpath : str
            The file path of the document to be inserted.
        keep_charshape : bool, optional
            If True, retains the original character shapes from the inserted file. Defaults to False.
        keep_parashape : bool, optional
            If True, retains the original paragraph shapes from the inserted file. Defaults to False.
        keep_section : bool, optional
            If True, retains the original section formatting from the inserted file. Defaults to False.
        keep_style : bool, optional
            If True, retains the original styles from the inserted file. Defaults to False.

        Returns
        -------
        bool
            True if the file was inserted successfully, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.insert_file('path/to/file.hwp', keep_style=True)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_file")


        action = self.actions.InsertFile
        p = action.pset
        p.filename = Path(fpath).absolute().as_posix()
        p.KeepCharshape = keep_charshape
        p.KeepParashape = keep_parashape
        p.KeepSection = keep_section
        p.KeepStyle = keep_style

        return action.run()

    def set_cell_border(
        self,
        top=None,
        right=None,
        left=None,
        bottom=None,
        horizontal=None,
        vertical=None,
        top_width=const.Thickness.NULL,
        right_width=const.Thickness.NULL,
        left_width=const.Thickness.NULL,
        bottom_width=const.Thickness.NULL,
        horizontal_width=const.Thickness.NULL,
        vertical_width=const.Thickness.NULL,
        top_color=None,
        bottom_color=None,
        left_color=None,
        right_color=None,
        horizontal_color=None,
        vertical_color=None,
    ):
        """
        Sets the border properties for cells in a table within a Hancom Office Hwp document.

        This function customizes the border types, widths, and colors for different sides of the cells.
        It allows for detailed customization of cell appearance in tables.

        Parameters
        ----------
        top, right, left, bottom, horizontal, vertical : int, optional
            Types of borders for the respective sides and internal lines of the cell (e.g., solid, dashed).
        top_width, right_width, left_width, bottom_width, horizontal_width, vertical_width : Thickness, optional
            Widths of the borders corresponding to the sides and internal lines of the cell. Specified using the `Thickness` Enum.
        top_color, bottom_color, left_color, right_color, horizontal_color, vertical_color : str, optional
            Colors for the respective borders in a string format (e.g., "#RRGGBB").

        Returns
        -------
        bool
            True if the border settings were successfully applied, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_cell_border(top=1, bottom=1, top_width=Thickness._0_25_mm, bottom_width=Thickness._0_25_mm, top_color="#FF0000", bottom_color="#00FF00")
        >>> print(success)

        Notes
        -----
        The function relies on the `self.actions.CellFill` action to set the border properties. The `Thickness` Enum provides predefined thickness levels for the borders. The color parameters should be provided in hex format.
        """
        attrs = {
            "BorderTypeTop": top,
            "BorderTypeRight": right,
            "BorderTypeLeft": left,
            "BorderTypeBottom": bottom,
            "TypeHorz": horizontal,
            "TypeVert": vertical,
            "BorderWidthTop": top_width.value,
            "BorderWidthRight": right_width.value,
            "BorderWidthLeft": left_width.value,
            "BorderWidthBottom": bottom_width.value,
            "WidthHorz": horizontal_width.value,
            "WidthVert": vertical_width.value,
            "BorderColorTop": self.api.RGBColor(get_rgb_tuple(top_color))
            if top_color
            else None,
            "BorderColorRight": self.api.RGBColor(get_rgb_tuple(right_color))
            if right_color
            else None,
            "BorderColorLeft": self.api.RGBColor(get_rgb_tuple(left_color))
            if left_color
            else None,
            "BorderColorBottom": self.api.RGBColor(get_rgb_tuple(bottom_color))
            if bottom_color
            else None,
            "ColorHorz": self.api.RGBColor(get_rgb_tuple(horizontal_color))
            if horizontal_color
            else None,
            "ColorVert": self.api.RGBColor(get_rgb_tuple(vertical_color))
            if vertical_color
            else None,
        }


        logger = get_logger('core')
        logger.debug(f"Calling set_cell_border")

        action = self.actions.CellFill
        p = action.pset

        for key, value in attrs.items():
            if value is None:
                continue
            setattr(p, key, value)

        return action.run()

    def set_cell_color(
        self, bg_color=None, hatch_color="#000000", hatch_style=6, alpha=None
    ):
        """
        Sets the background color and hatch style for cells in a Hancom Office Hwp document.

        This function allows customization of the cell background, including color and hatching pattern,
        providing options to set transparency and hatch color.

        Parameters
        ----------
        bg_color : str, optional
            Hexadecimal color code for the cell background (e.g., "#RRGGBB"). If not specified, the background is not changed.
        hatch_color : str, optional
            Hexadecimal color code for the hatching. Default is black ("#000000").
        hatch_style : int, optional
            Style of the hatching pattern. Default is 6.
        alpha : int, optional
            Alpha value for the background color's transparency (0-255). If not specified, transparency is not changed.

        Returns
        -------
        bool
            True if the cell color settings were applied successfully, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_cell_color(bg_color="#FF0000", hatch_color="#0000FF", hatch_style=5, alpha=128)
        >>> print(success)

        Notes
        -----
        The function uses the `self.actions.CellBorderFill` action for setting the cell properties.
        Colors are specified in hexadecimal format. The alpha parameter controls the transparency of the background color.
        """


        logger = get_logger('core')
        logger.debug(f"Calling set_cell_color")

        fill_type = windows_brush = None
        if bg_color:
            fill_type = 1
            windows_brush = 1

        attrs = {
            "type": fill_type,
            "WindowsBrush": windows_brush,
            "WinBrushFaceColor": self.api.RGBColor(*get_rgb_tuple(bg_color)) if bg_color else None,
            "WinBrushAlpha": alpha,
            "WinBrushHatchColor": self.api.RGBColor(*get_rgb_tuple(hatch_color)) if hatch_color else None,
            "WinBrushFaceStyle": hatch_style,
        }

        action = self.actions.CellBorderFill
        p = action.pset

        for key, value in attrs.items():
            if value is not None:
                setattr(p.FillAttr, key, value)

        return action.run()


def _get_text(app):
    """스캔한 텍스트 텍스트 제너레이터"""
    flag, text = 2, ""
    while flag not in [0, 1, 101, 102]:
        flag, text = app.api.GetText()
        yield text

def move_to_line(app: App, text):
    """인자로 전달한 텍스트가 있는 줄의 시작지점으로 이동합니다."""
    logger = get_logger('core')
    logger.debug(f"Calling move_to_line")
    with app.scan(scan_spos=const.ScanStartPosition.Line) as scan:
        for line in scan:
            if text in line:
                return app.move.scan_pos()
    return False
