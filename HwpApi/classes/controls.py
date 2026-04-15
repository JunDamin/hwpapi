"""
Controls accessor — ``app.controls``.

문서 내 모든 컨트롤(표, 그림, 구역정의, 하이퍼링크, 머리말 등)을
linked list (``HeadCtrl`` → ``Next`` 체인) 로 순회 · 검색 · 조작.

Phase E deliverable (PYHWPX_COMPARISON.md roadmap).
"""
from __future__ import annotations
from typing import Optional, Iterator, List

__all__ = ["ControlsAccessor", "Control"]


class Control:
    """
    단일 컨트롤 객체 — ``IDHwpCtrlCode`` COM 래퍼.

    Attributes
    ----------
    raw : COM object
        원시 ``IDHwpCtrlCode``. 직접 속성 접근에 사용.

    Examples
    --------
    >>> for ctrl in app.controls:
    ...     print(ctrl.ctrl_id, ctrl.user_desc)
    >>> ctrl = app.controls.find(desc="표")
    >>> ctrl.select()
    >>> ctrl.delete()
    """

    __slots__ = ("raw", "_app")

    def __init__(self, raw_ctrl, app=None):
        object.__setattr__(self, "raw", raw_ctrl)
        object.__setattr__(self, "_app", app)

    @property
    def ctrl_id(self) -> str:
        """HWP 내부 컨트롤 타입 코드 (예: ``'tbl '``, ``'secd'``, ``'%hlk'``)."""
        try:
            return str(self.raw.CtrlID or "")
        except Exception:
            return ""

    @property
    def user_desc(self) -> str:
        """사람 읽을 수 있는 설명 (예: '표', '구역 정의', '하이퍼링크')."""
        try:
            return str(getattr(self.raw, "UserDesc", "") or "")
        except Exception:
            return ""

    @property
    def properties(self):
        """컨트롤의 ParameterSet 속성 (있는 경우)."""
        try:
            return self.raw.Properties
        except Exception:
            return None

    def select(self) -> bool:
        """
        이 컨트롤을 선택 상태로 만듦.

        Notes
        -----
        HWP 의 ``SelectCtrl`` 액션을 사용. 컨트롤이 보이는 영역에 있어야 함.
        """
        if self._app is None:
            return False
        try:
            self._app.api.SelectCtrl(self.raw)
            return True
        except Exception:
            try:
                # Fallback: move cursor to ctrl then run SelectCtrl action
                self._app.api.SetPosBySet(
                    self.raw.ParaAddr if hasattr(self.raw, "ParaAddr") else None
                )
                self._app.api.Run("SelectCtrlFront")
                return True
            except Exception:
                return False

    def delete(self) -> bool:
        """이 컨트롤을 문서에서 제거."""
        if self._app is None:
            return False
        try:
            self._app.api.DeleteCtrl(self.raw)
            return True
        except Exception:
            return False

    def move_to(self) -> bool:
        """커서를 이 컨트롤 위치로 이동 (선택 없이)."""
        if self._app is None:
            return False
        try:
            self._app.api.SelectCtrl(self.raw)
            self._app.api.Run("Cancel")
            return True
        except Exception:
            return False

    def __repr__(self):
        return f"<Control {self.ctrl_id!r} desc={self.user_desc!r}>"

    def __eq__(self, other):
        if isinstance(other, Control):
            try:
                return self.raw == other.raw
            except Exception:
                return False
        return NotImplemented


class ControlsAccessor:
    """
    ``app.controls`` — 문서 내 컨트롤 iterator.

    Examples
    --------
    >>> len(app.controls)
    >>> for c in app.controls:
    ...     print(c)

    >>> # 타입 코드로 필터
    >>> tables = [c for c in app.controls if c.ctrl_id.strip() == "tbl"]

    >>> # 편의 검색
    >>> tbl = app.controls.find(desc="표")
    >>> tbl.select()

    >>> # 현재 선택된 컨트롤
    >>> app.controls.selected

    >>> # 타입별 묶기
    >>> by_type = app.controls.by_ctrl_id()
    >>> by_type["tbl "]      # list of Control for tables
    """

    # 주요 HWP CtrlID 값 → 한글 설명
    KNOWN_IDS = {
        "secd": "구역 정의",
        "cold": "단 정의",
        "tbl ": "표",
        "gso ": "그림 / 도형",
        "%hlk": "하이퍼링크",
        "bokm": "북마크",
        "%fn ": "각주",
        "%en ": "미주",
        "%cmt": "메모",
        "header": "머리말",
        "footer": "꼬리말",
    }

    def __init__(self, app):
        self._app = app

    def __iter__(self) -> Iterator[Control]:
        """``HeadCtrl`` 에서 시작해 ``Next`` 링크를 따라 순회."""
        try:
            ctrl = self._app.api.HeadCtrl
        except Exception:
            return
        seen = set()  # avoid infinite loops
        while ctrl is not None:
            # Use id() of the COM wrapper for duplicate detection
            key = id(ctrl)
            if key in seen:
                break
            seen.add(key)
            yield Control(ctrl, self._app)
            try:
                ctrl = ctrl.Next
            except Exception:
                break

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __getitem__(self, index: int) -> Control:
        """인덱스로 접근."""
        for i, c in enumerate(self):
            if i == index:
                return c
        raise IndexError(f"Control index {index} out of range")

    @property
    def head(self) -> Optional[Control]:
        """첫 번째 컨트롤."""
        try:
            raw = self._app.api.HeadCtrl
            return Control(raw, self._app) if raw else None
        except Exception:
            return None

    @property
    def last(self) -> Optional[Control]:
        """마지막 컨트롤."""
        try:
            raw = self._app.api.LastCtrl
            return Control(raw, self._app) if raw else None
        except Exception:
            return None

    @property
    def selected(self) -> Optional[Control]:
        """현재 선택된 컨트롤 (없으면 None)."""
        try:
            raw = self._app.api.CurSelectedCtrl
            return Control(raw, self._app) if raw else None
        except Exception:
            return None

    def find(self,
             ctrl_id: Optional[str] = None,
             desc: Optional[str] = None) -> Optional[Control]:
        """
        첫 매치되는 컨트롤 반환.

        Parameters
        ----------
        ctrl_id : str | None
            HWP 내부 타입 코드 (예: ``"tbl "``).
        desc : str | None
            ``user_desc`` 부분 일치.
        """
        for c in self:
            if ctrl_id is not None and c.ctrl_id.strip() != ctrl_id.strip():
                continue
            if desc is not None and desc not in c.user_desc:
                continue
            return c
        return None

    def find_all(self,
                 ctrl_id: Optional[str] = None,
                 desc: Optional[str] = None) -> List[Control]:
        """조건에 맞는 모든 컨트롤."""
        out = []
        for c in self:
            if ctrl_id is not None and c.ctrl_id.strip() != ctrl_id.strip():
                continue
            if desc is not None and desc not in c.user_desc:
                continue
            out.append(c)
        return out

    def by_ctrl_id(self) -> dict:
        """``{ctrl_id: [Control, ...]}`` 그룹화."""
        out = {}
        for c in self:
            out.setdefault(c.ctrl_id, []).append(c)
        return out

    def unselect(self) -> bool:
        """현재 컨트롤 선택 해제."""
        try:
            self._app.api.UnSelectCtrl()
            return True
        except Exception:
            return False

    def __repr__(self):
        return f"<ControlsAccessor count={len(self)}>"
