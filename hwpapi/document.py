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
class _DocActions:
    """Document-scoped actions — `doc.actions.X` 가 doc 자동 활성화."""

    __slots__ = ("_doc",)

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __getattr__(self, name: str):
        action = getattr(self._doc._app.actions, name)
        return _ScopedAction(action, self._doc)

    def __repr__(self) -> str:
        return f"<DocActions doc={self._doc.name!r}>"


class _ScopedAction:
    """`run()` 직전 doc.activate() 자동 호출하는 wrapper."""

    __slots__ = ("_action", "_doc")

    def __init__(self, action, doc: "Document") -> None:
        self._action = action
        self._doc = doc

    @property
    def pset(self):
        return self._action.pset

    @property
    def act(self):
        return self._action.act

    def run(self, *args, **kwargs):
        self._doc.activate()
        return self._action.run(*args, **kwargs)

    def __repr__(self) -> str:
        return f"<ScopedAction doc={self._doc.name!r} action={self._action!r}>"


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
