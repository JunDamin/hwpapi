"""
:mod:`hwpapi.document` — v3 ``Document`` facade.

ADR-003 의 결정에 따라 v3 부터 ``Document`` 가 모든 문서 단위 작업을
책임집니다 — text I/O / 컬렉션 / save / close / 활성화. ``App`` 은
HWP 프로세스 lifecycle 만.

획득은 ``app.docs.open(path)`` 또는 ``app.docs.add()`` 로:

    from hwpapi import App
    app = App()
    doc = app.docs.open("report.hwp")
    doc.insert_text("hi\\n")
    doc.fields["author"] = "홍길동"
    doc.save()
    doc.close()

다중 문서 시나리오에서는 각 ``Document`` 인스턴스의 메소드를 호출할
때마다 :meth:`Document.activate` 가 자동 실행되어 그 문서가 활성
상태가 됩니다 (xlwings 패턴).
"""
from __future__ import annotations

from functools import cached_property
from pathlib import Path
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App
    from hwpapi.collections.bookmarks import BookmarkCollection
    from hwpapi.collections.fields import FieldCollection
    from hwpapi.collections.hyperlinks import HyperlinkCollection
    from hwpapi.collections.images import ImageCollection
    from hwpapi.collections.paragraphs import ParagraphCollection
    from hwpapi.collections.styles import StyleCollection
    from hwpapi.collections.tables import TableCollection

__all__ = ["Document"]


# ── format maps (mirror App's) ────────────────────────────────────
_SAVE_FORMAT_MAP = {
    ".hwp": "HWP",
    ".hwpx": "HWPX",
    ".hml": "HML",
    ".pdf": "PDF",
    ".png": "PNG",
    ".txt": "TEXT",
    ".docx": "DOCX",
    ".html": "HTML",
    ".htm": "HTML",
}


def _format_from_suffix(suffix: str) -> Optional[str]:
    return _SAVE_FORMAT_MAP.get(suffix.lower())


# ── Document-scoped actions proxy ────────────────────────────────
class _DocCursor:
    """Per-document cursor — 이동 / 위치 검사."""

    __slots__ = ("_doc",)

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def goto_page(self, n: int) -> "Document":
        """페이지 ``n`` (1-based) 로 이동."""
        self._doc.activate()
        try:
            self._doc._app.api.MovePos(23, n - 1, 0)  # MoveToPageStart
        except Exception:
            pass
        return self._doc

    def in_table(self) -> bool:
        """현재 커서가 표 안에 있으면 ``True``."""
        self._doc.activate()
        try:
            ki = self._doc._app.api.KeyIndicator()
            return bool(ki[6])
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"<DocCursor doc={self._doc.name!r}>"


class _DocActions:
    """Document-scoped actions — `doc.actions.X` 가 doc 자동 활성화.

    중요: ``__getattr__`` 시점 (= action 객체 획득 시) 에 즉시
    ``doc.activate()`` 를 호출합니다. 그 이유는 HWP 의 ``HAction`` /
    ``HParameterSet`` 이 활성 문서의 컨텍스트에 묶여 있어, pset.X 값을
    설정한 뒤 doc 을 전환하면 그 값이 무효화되기 때문입니다. 따라서
    "pset 만지기 전에 doc 활성화" 가 정공법.
    """

    __slots__ = ("_doc",)

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __getattr__(self, name: str):
        # 활성화는 여기서 — pset 을 만지기 전에 doc 컨텍스트를 잡음.
        self._doc.activate()
        return getattr(self._doc._app.actions, name)

    def __repr__(self) -> str:
        return f"<DocActions doc={self._doc.name!r}>"


# ── Document ─────────────────────────────────────────────────────
class Document:
    """
    Per-document façade — v3 의 1차 surface.

    ``app.docs.open(path)`` 또는 ``app.docs.add()`` 로 획득합니다.
    모든 doc-scoped 메소드는 호출 직전 :meth:`activate` 를 자동
    실행하므로 다중 문서 환경에서 안전합니다.

    Examples
    --------
    >>> from hwpapi import App
    >>> with App() as app:
    ...     doc = app.docs.open("report.hwp")
    ...     doc.insert_text("hi\\n")
    ...     doc.fields["author"] = "홍길동"
    ...     doc.save()
    ...     doc.close()
    """

    __slots__ = ("_app", "_raw", "__dict__")

    def __init__(self, app: "App", _raw=None, _index: Optional[int] = None) -> None:
        """``_raw`` (IXHwpDocument 핸들) 우선. 없으면 ``_index`` 로 즉시 resolve."""
        self._app = app
        if _raw is None and _index is not None:
            try:
                _raw = app.api.XHwpDocuments.Item(_index)
            except Exception:
                _raw = None
        self._raw = _raw

    # ── meta ─────────────────────────────────────────────────────

    @property
    def index(self) -> int:
        """현재 ``XHwpDocuments`` 안의 인덱스 — 다른 doc 추가/제거 시 변동."""
        if self._raw is None:
            return -1
        try:
            xhd = self._app.api.XHwpDocuments
            for i in range(int(xhd.Count)):
                if xhd.Item(i) == self._raw:
                    return i
        except Exception:
            pass
        return -1

    @property
    def path(self) -> Optional[str]:
        """전체 경로 — 저장된 적 없는 새 문서면 ``None``."""
        if self._raw is None:
            return None
        try:
            full = str(self._raw.FullName)
            return full if full else None
        except Exception:
            return None

    @property
    def name(self) -> str:
        """파일명 — 새 문서면 ``"(untitled)"``."""
        p = self.path
        return Path(p).name if p else "(untitled)"

    @property
    def saved(self) -> bool:
        """``True`` 면 미저장 변경사항이 없음."""
        if self._raw is None:
            return True
        try:
            return not bool(getattr(self._raw, "Modified", False))
        except Exception:
            return True

    @property
    def raw(self):
        """``IXHwpDocument`` COM 핸들 (escape hatch)."""
        return self._raw

    # ── lifecycle ────────────────────────────────────────────────

    def activate(self) -> "Document":
        """이 문서를 HWP 의 활성 문서로 만듦. 자기 자신 반환 (chain)."""
        if self._raw is not None:
            try:
                self._raw.SetActive_XHwpDocument()
            except Exception:
                pass
        return self

    def save(
        self,
        path: Optional[str] = None,
        format: Optional[str] = None,
    ) -> Optional[str]:
        """
        제자리 저장 또는 새 경로로 저장.

        ``path`` 없으면 ``HwpObject.Save`` (현재 경로). ``path`` 있으면
        확장자로 포맷 추론 후 ``SaveAs``.
        """
        from hwpapi.functions import get_absolute_path

        self.activate()
        if not path:
            self._app.api.Save()
            return self.path
        name = get_absolute_path(path)
        fmt = format or _format_from_suffix(Path(name).suffix)
        self._app.api.SaveAs(name, fmt)
        return name

    def save_as(
        self,
        path: str,
        format: Optional[str] = None,
    ) -> Optional[str]:
        """:meth:`save` 의 명시 alias — 항상 ``path`` 인자 필수."""
        return self.save(path, format)

    def close(self, save: bool = False) -> bool:
        """
        문서 닫기. ``save=False`` (기본) 면 변경사항 버리고 닫음.

        직접 ``IXHwpDocument.Close(save)`` 호출 — dialog 가 안 뜨므로
        :func:`SetMessageBoxMode` 영향 없음.
        """
        if self._raw is None:
            return False
        try:
            self._raw.Close(save)
            self._raw = None  # 핸들 무효화
            return True
        except Exception:
            return False

    # ── text I/O ─────────────────────────────────────────────────

    @property
    def text(self) -> str:
        """문서 전체 텍스트 (read)."""
        self.activate()
        try:
            return str(self._app.api.GetTextFile("TEXT", ""))
        except Exception:
            return ""

    def insert_text(self, s: str) -> "Document":
        """
        커서 위치에 텍스트 삽입. ``"\\n"`` 은 ``BreakPara`` 로 변환.
        """
        self.activate()
        parts = s.split("\n")
        for i, part in enumerate(parts):
            if part:
                act = self._app.actions.InsertText
                act.pset.Text = part
                act.run()
            if i < len(parts) - 1:
                self._app.api.Run("BreakPara")
        return self

    def select_all(self) -> "Document":
        """전체 선택."""
        self.activate()
        self._app.api.Run("SelectAll")
        return self

    def get_selected_text(self) -> str:
        """현재 선택된 텍스트."""
        self.activate()
        try:
            return str(self._app.api.GetSelectedText())
        except Exception:
            return ""

    def clear(self) -> "Document":
        """문서 내용 전체 삭제."""
        self.activate()
        self._app.api.Run("SelectAll")
        self._app.api.Run("Delete")
        return self

    # ── 직접 setter (with 없이 즉시 적용) — ADR-005 ─────────────

    def set_charshape(self, **fmt) -> "Document":
        """현재 선택/커서에 글자 서식을 즉시 적용 (with 없이).

        선택 영역이 있으면 그 영역에만, 없으면 cursor 의 다음 입력에
        적용됩니다. 사용 가능 키: ``bold``, ``italic``, ``size``,
        ``height``, ``color`` / ``text_color``, ``shade_color``,
        ``font`` / ``facename``, ``underline_type``, ``strike_out_type``
        등 (charshape_scope 와 동일).

        Examples
        --------
        >>> doc.find_text("강조")    # 선택됨
        >>> doc.set_charshape(bold=True, text_color="#E74C3C")

        >>> # 다음 입력만 적용 (커서)
        >>> doc.set_charshape(bold=True)
        >>> doc.insert_text("이 줄은 굵게")
        """
        from hwpapi.context.scopes import _CHAR_ALIAS, _translate, _apply
        self.activate()
        translated = _translate(fmt, _CHAR_ALIAS)
        _apply(self._app, "CharShape", "CharShape", translated)
        return self

    def set_parashape(self, **fmt) -> "Document":
        """현재 단락에 단락 서식을 즉시 적용 (with 없이).

        사용 가능 키: ``align`` (``"left"|"center"|"right"|"justify"``),
        ``line_spacing``, ``left_margin``, ``right_margin``,
        ``indentation`` 등.

        Examples
        --------
        >>> doc.set_parashape(align="center", line_spacing=180)
        """
        from hwpapi.context.scopes import (
            _PARA_ALIAS, _translate, _apply, _normalise_align,
        )
        self.activate()
        translated = _translate(fmt, _PARA_ALIAS)
        if "AlignType" in translated:
            translated["AlignType"] = _normalise_align(translated["AlignType"])
        _apply(self._app, "ParaShape", "ParaShape", translated)
        return self

    def find_text(self, query: str) -> bool:
        """현재 커서 위치 이후에서 ``query`` 검색. 발견 시 ``True``."""
        self.activate()
        try:
            return bool(self._app.api.HAction.GetDefault) and \
                   bool(self._app.actions.FindDlg.run())
        except Exception:
            return False

    def replace_all(self, find: str, replace: str) -> int:
        """``find`` → ``replace`` 일괄 치환. 치환된 개수 반환 (대략)."""
        self.activate()
        # AllReplace action — ParameterSet 기반.
        try:
            act = self._app.actions.AllReplace
            pset = act.pset
            pset.FindString = find
            pset.ReplaceString = replace
            pset.Direction = 0          # forward
            pset.WholeWordOnly = 0
            pset.IgnoreMessage = 1
            pset.MatchCase = 1
            act.run()
            # HWP 가 정확한 개수를 돌려주지 않음 — 호출 성공 여부만.
            return 1
        except Exception:
            return 0

    def select_text(self, start: int, end: int) -> "Document":
        """문자 위치 ``start..end`` 를 선택."""
        self.activate()
        # SelectByPos 또는 MovePos / MoveSelPos 조합.
        try:
            self._app.api.MovePos(2, 0, 0)              # MoveTopOfFile
            for _ in range(start):
                self._app.api.Run("MoveRight")
            for _ in range(end - start):
                self._app.api.Run("MoveSelRight")
        except Exception:
            pass
        return self

    def copy(self) -> "Document":
        self.activate()
        self._app.api.Run("Copy")
        return self

    def cut(self) -> "Document":
        self.activate()
        self._app.api.Run("Cut")
        return self

    def paste(self) -> "Document":
        self.activate()
        self._app.api.Run("Paste")
        return self

    def delete(self) -> "Document":
        self.activate()
        self._app.api.Run("Delete")
        return self

    def undo(self) -> "Document":
        self.activate()
        self._app.api.Run("Undo")
        return self

    def redo(self) -> "Document":
        self.activate()
        self._app.api.Run("Redo")
        return self

    def insert_line_break(self) -> "Document":
        self.activate()
        self._app.api.Run("BreakLine")
        return self

    def insert_page_break(self) -> "Document":
        self.activate()
        self._app.api.Run("BreakPage")
        return self

    def insert_paragraph_break(self) -> "Document":
        self.activate()
        self._app.api.Run("BreakPara")
        return self

    def insert_tab(self) -> "Document":
        self.activate()
        self._app.api.Run("InsertTab")
        return self

    def insert_picture(self, path: str) -> "Document":
        """그림 삽입 — `path` 의 그림 파일을 커서 위치에 삽입."""
        from hwpapi.functions import get_absolute_path
        self.activate()
        try:
            self._app.api.InsertPicture(get_absolute_path(path), True, 0, 0)
        except Exception as e:
            self._app.logger.debug(f"insert_picture: {e}")
        return self

    def insert_table(self, rows: int, cols: int) -> "Document":
        """``rows × cols`` 표 삽입."""
        self.activate()
        try:
            act = self._app.actions.TableCreate
            act.pset.Rows = rows
            act.pset.Cols = cols
            act.run()
        except Exception as e:
            self._app.logger.debug(f"insert_table: {e}")
        return self

    # ── cursor sub-accessor ─────────────────────────────────────

    @cached_property
    def cursor(self) -> "_DocCursor":
        """:class:`_DocCursor` — 커서 이동/검사."""
        return _DocCursor(self)

    # ── Selection / Range — ADR-005 v3.1 ─────────────────────────

    @property
    def selection(self):
        """현재 활성 선택 영역 핸들 — :class:`hwpapi.selection.Selection`."""
        from hwpapi.selection import Selection
        return Selection(self)

    def range(self, start_para: int, end_para: Optional[int] = None):
        """단락 범위 핸들 — :class:`hwpapi.selection.Range`.

        Parameters
        ----------
        start_para : int
            시작 단락 인덱스 (0-based, inclusive).
        end_para : int, optional
            끝 단락 (inclusive). ``None`` 이면 단일 단락.
        """
        from hwpapi.selection import Range
        return Range(self, start_para, end_para)

    # ── find_all / replace_brackets ─────────────────────────────

    def find_all(self, query: str, max_matches: int = 1000) -> list:
        """문서 전체에서 ``query`` 의 모든 위치를 찾아 list 반환.

        반환값은 (start_pos, end_pos) 튜플 리스트 — HWP 의 GetPos
        반환 형태. 다음 작업의 base 로 사용 (예: 일괄 서식 변경).

        Examples
        --------
        >>> for pos in doc.find_all("강조"):
        ...     # 각 위치에서 일괄 처리
        ...     pass
        """
        self.activate()
        api = self._app.api
        results: list = []
        try:
            # MoveTopOfFile
            api.MovePos(2, 0, 0)
            for _ in range(max_matches):
                # FindCtrl 류는 dialog 띄움 — 직접 탐색은 GetText 후 파싱이
                # 더 안정적이지만 spike 단계에선 raw 사용.
                # 실제 동작 검증 필요 — placeholder 구현.
                ok = bool(api.Run("RepeatFind") if results else api.Run("FindDlg"))
                if not ok:
                    break
                try:
                    pos = api.GetPos()
                    results.append(tuple(pos))
                except Exception:
                    break
        except Exception as e:
            self._app.logger.debug(f"find_all: {e}")
        return results

    def replace_brackets(self, mapping: dict) -> int:
        """``{key}`` 형태의 placeholder 를 ``mapping`` 의 값으로 일괄 치환.

        메일 머지의 가장 흔한 패턴 — 템플릿에 ``{name}``, ``{date}`` 등
        의 마커가 있을 때.

        Parameters
        ----------
        mapping : dict
            ``{"{name}": "홍길동", "{date}": "2026-04-29"}`` 형태.

        Returns
        -------
        int
            치환을 시도한 키 개수 (성공/실패 구분 없음 — HWP 가
            정확한 카운트 미제공).

        Examples
        --------
        >>> doc.replace_brackets({
        ...     "{name}": "홍길동",
        ...     "{date}": "2026-04-29",
        ...     "{amount}": "1,200,000원",
        ... })
        """
        count = 0
        for key, value in mapping.items():
            try:
                self.replace_all(str(key), str(value))
                count += 1
            except Exception as e:
                self._app.logger.debug(f"replace_brackets[{key}]: {e}")
        return count

    # ── action proxy ─────────────────────────────────────────────

    @property
    def actions(self) -> "_DocActions":
        """
        Doc-scoped actions — ``doc.actions.X.run()`` 이 doc 자동 활성화.

        Examples
        --------
        >>> doc.actions.InsertText.pset.Text = "hi"
        >>> doc.actions.InsertText.run()
        """
        return _DocActions(self)

    # ── Collection properties (v2 호환) ──────────────────────────

    @cached_property
    def fields(self) -> "FieldCollection":
        from hwpapi.collections.fields import FieldCollection
        return FieldCollection(self._app)

    @cached_property
    def bookmarks(self) -> "BookmarkCollection":
        from hwpapi.collections.bookmarks import BookmarkCollection
        return BookmarkCollection(self._app)

    @cached_property
    def hyperlinks(self) -> "HyperlinkCollection":
        from hwpapi.collections.hyperlinks import HyperlinkCollection
        return HyperlinkCollection(self._app)

    @cached_property
    def images(self) -> "ImageCollection":
        from hwpapi.collections.images import ImageCollection
        return ImageCollection(self._app)

    @cached_property
    def styles(self) -> "StyleCollection":
        from hwpapi.collections.styles import StyleCollection
        return StyleCollection(self._app)

    @cached_property
    def paragraphs(self) -> "ParagraphCollection":
        from hwpapi.collections.paragraphs import ParagraphCollection
        return ParagraphCollection(self._app)

    @cached_property
    def tables(self) -> "TableCollection":
        from hwpapi.collections.tables import TableCollection
        return TableCollection(self._app)

    # ── escape hatch ─────────────────────────────────────────────

    @property
    def app(self) -> "App":
        """소유 :class:`App` (escape hatch)."""
        return self._app

    def __repr__(self) -> str:
        if self._raw is None:
            return "<Document closed>"
        return f"<Document #{self.index} name={self.name!r}>"
