"""
:mod:`hwpapi.selection` — Selection / Range objects (v3.1).

ADR-005 의 권장 항목 — `doc.selection` 과 `doc.range(...)` 로 raw API
의존을 줄이고 xlwings 의 `app.selection` / `sheet.range(...)` 와 동형
멘탈 모델을 제공합니다.

핵심:
- :class:`Selection` — 현재 활성 선택 영역에 대한 핸들. 텍스트 읽기,
  서식 적용, 삭제 등을 owning doc 자동 활성화로 안전하게 수행.
- :class:`Range` — 명시적 범위 (단락 단위) 핸들. 사용자가 좌표를
  지정해 만들고, 메소드 진입 시 select + 작업 + (옵션) cancel.

모든 메소드는 owning :class:`hwpapi.Document` 의 :meth:`activate` 를
먼저 호출 — 다중 문서 환경에서 안전.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from hwpapi.document import Document

__all__ = ["Selection", "Range"]


class Selection:
    """현재 활성 선택 영역 핸들 — `doc.selection` 으로 접근.

    선택 상태가 없으면 :attr:`exists` 는 ``False``. 선택 상태일 때만
    text / set_charshape / delete 등이 의미 있음.
    """

    __slots__ = ("_doc",)

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    @property
    def exists(self) -> bool:
        """선택 영역이 활성 상태인지."""
        self._doc.activate()
        try:
            text = str(self._doc._app.api.GetSelectedText())
            return len(text) > 0
        except Exception:
            return False

    @property
    def text(self) -> str:
        """선택된 텍스트 (없으면 빈 문자열)."""
        self._doc.activate()
        try:
            return str(self._doc._app.api.GetSelectedText())
        except Exception:
            return ""

    def set_charshape(self, **fmt) -> "Selection":
        """선택 영역에 글자 서식 즉시 적용. 선택 유지."""
        self._doc.set_charshape(**fmt)
        return self

    def set_parashape(self, **fmt) -> "Selection":
        """선택 영역의 단락에 단락 서식 적용."""
        self._doc.set_parashape(**fmt)
        return self

    def delete(self) -> "Selection":
        """선택 영역 삭제."""
        self._doc.activate()
        self._doc._app.api.Run("Delete")
        return self

    def copy(self) -> "Selection":
        self._doc.activate()
        self._doc._app.api.Run("Copy")
        return self

    def cut(self) -> "Selection":
        self._doc.activate()
        self._doc._app.api.Run("Cut")
        return self

    def cancel(self) -> "Selection":
        """선택 해제."""
        self._doc.activate()
        self._doc._app.api.Run("Cancel")
        return self

    def __repr__(self) -> str:
        if not self.exists:
            return "<Selection (empty)>"
        text = self.text
        preview = text[:30] + "..." if len(text) > 30 else text
        return f"<Selection len={len(text)} text={preview!r}>"


class Range:
    """명시적 단락 범위 핸들 — `doc.range(start_para, end_para)` 로 생성.

    Parameters
    ----------
    doc : Document
        Owning document.
    start_para : int
        시작 단락 (0-based, inclusive).
    end_para : int, optional
        끝 단락 (0-based, inclusive). ``None`` 이면 ``start_para`` 와
        같음 (단일 단락).

    Notes
    -----
    내부적으로 메소드 호출 시마다 `MovePos` + `MoveSel` 로 선택 상태를
    재구성합니다. doc 편집 중 단락 인덱스가 밀리면 부정확해질 수 있음.
    """

    __slots__ = ("_doc", "_start", "_end")

    def __init__(
        self,
        doc: "Document",
        start_para: int,
        end_para: Optional[int] = None,
    ) -> None:
        self._doc = doc
        self._start = int(start_para)
        self._end = int(end_para) if end_para is not None else self._start

    def _select(self) -> None:
        """내부: 이 범위를 HWP 에서 선택 상태로."""
        self._doc.activate()
        api = self._doc._app.api
        # 시작 단락 첫 글자로 이동
        try:
            api.MovePos(2, 0, 0)  # MoveTopOfFile
            for _ in range(self._start):
                api.Run("MoveDown")
            api.Run("MoveLineBegin")
            # end 까지 select
            spans = self._end - self._start
            for _ in range(spans):
                api.Run("MoveSelDown")
            api.Run("MoveSelLineEnd")
        except Exception:
            pass

    @property
    def text(self) -> str:
        """범위 안의 텍스트."""
        self._select()
        try:
            return str(self._doc._app.api.GetSelectedText())
        except Exception:
            return ""

    @text.setter
    def text(self, value: str) -> None:
        """범위를 ``value`` 로 교체."""
        self._select()
        api = self._doc._app.api
        api.Run("Delete")
        # insert via Document.insert_text — \n 처리 포함
        self._doc.insert_text(value)

    def set_charshape(self, **fmt) -> "Range":
        """범위에 글자 서식 적용."""
        self._select()
        self._doc.set_charshape(**fmt)
        return self

    def set_parashape(self, **fmt) -> "Range":
        self._select()
        self._doc.set_parashape(**fmt)
        return self

    def delete(self) -> "Range":
        self._select()
        self._doc._app.api.Run("Delete")
        return self

    def copy(self) -> "Range":
        self._select()
        self._doc._app.api.Run("Copy")
        return self

    @property
    def start(self) -> int:
        return self._start

    @property
    def end(self) -> int:
        return self._end

    def __repr__(self) -> str:
        return f"<Range doc={self._doc.name!r} para={self._start}..{self._end}>"
