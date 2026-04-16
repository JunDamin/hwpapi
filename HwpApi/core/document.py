"""
Multi-document support — ``Document`` (single doc handle) and
``Documents`` (collection accessor on ``app.documents``).

Extracted from ``hwpapi/core/app.py`` during the Phase 2 refactor
(REFACTORING_PLAN.md — P1-1) to keep the main App file manageable.

Public types re-exported from :mod:`hwpapi.core`:

- :class:`Document` — xlwings-style single-doc handle. Proxies all App
  methods via ``__getattr__`` with automatic activation.
- :class:`Documents` — list-like collection of all open HWP documents.
"""
from __future__ import annotations

__all__ = ["Document", "Documents"]

class Document:
    """
    여러 HWP 문서 중 **하나** 를 감싼 핸들.

    ``app.documents[i]`` 또는 ``app.documents.active`` 로 얻을 수 있습니다.
    문서 단위의 파일 경로 조회, 저장, 닫기, 활성화 등을 제공합니다.

    Attributes
    ----------
    raw : COM object
        원시 ``IXHwpDocument`` COM 래퍼. 직접 속성 접근이 필요할 때 사용.

    Examples
    --------
    >>> doc = app.documents[0]
    >>> print(doc.path)          # 현재 문서 경로
    >>> print(doc.modified)      # 수정 여부
    >>> doc.save()               # 현재 문서만 저장
    >>> doc.activate()           # 이 문서를 활성창으로 전환
    >>> doc.close()              # 이 문서만 닫기
    """

    # ── xlwings-style: App 메소드를 Document 에서도 쓸 수 있도록 위임 ─
    #
    # Document 자체에 정의된 속성/메소드 (file I/O, 상태 조회) 는 그대로
    # 쓰고, 그 외의 이름 — insert_text, set_charshape, find_text, move,
    # actions 등 — 을 ``doc.foo`` 로 접근하면 자동으로 (1) 이 문서를 활성
    # 창으로 전환, (2) App 의 동일 이름 메소드/속성 호출, (3) 원래 활성
    # 문서로 복원 후 결과 반환하도록 프록시합니다.
    #
    # 예:
    #     doc.insert_text("...")        # app.insert_text 와 동일하지만
    #                                    # doc 에 자동 활성화
    #     doc.replace_all("a", "b")
    #     doc.set_charshape(bold=True)
    #     doc.styled_text("X", bold=True)
    #     doc.actions.SelectAll.run()   # activate → 접근자 반환
    #     doc.move.top_of_file()

    def __init__(self, raw_doc, documents=None):
        # NOTE: set via object.__setattr__ to avoid __getattr__ recursion
        # during init (before attributes exist).
        object.__setattr__(self, "raw", raw_doc)
        object.__setattr__(self, "_collection", documents)
        object.__setattr__(self, "_app",
                           documents._app if documents else None)

    def __setattr__(self, name, value):
        """
        If ``name`` is a property setter on App, activate this doc and
        call the setter so e.g. ``doc.text = "..."`` writes to THIS doc.
        Otherwise fall through to normal attribute assignment.
        """
        # Underscore / own attributes — set directly
        if name.startswith("_") or name == "raw":
            object.__setattr__(self, name, value)
            return
        app = getattr(self, "_app", None)
        if app is not None:
            import inspect as _inspect
            cls_attr = _inspect.getattr_static(type(app), name, None)
            if isinstance(cls_attr, property) and cls_attr.fset is not None:
                self.activate()
                cls_attr.fset(app, value)
                return
        object.__setattr__(self, name, value)

    def __getattr__(self, name):
        """
        Proxy to ``self._app`` for any name not defined on Document itself.

        Callable → wrapped to auto-activate the document before calling.
        Non-callable (e.g. ``app.move`` accessor) → activate then return
        the attribute (caller should use it immediately).
        """
        # Python calls __getattr__ only when normal lookup fails,
        # so we won't interfere with Document's own attrs.
        if name.startswith("_"):
            raise AttributeError(name)

        app = object.__getattribute__(self, "_app")
        if app is None:
            raise AttributeError(
                f"{type(self).__name__!r} object has no attribute {name!r} "
                f"and no bound App to proxy through"
            )

        # If it's a PROPERTY on App, activate this doc FIRST, then
        # evaluate the getter — so ``doc.text`` returns THIS doc's text,
        # not the previously-active doc's.
        import inspect as _inspect
        cls_attr = _inspect.getattr_static(type(app), name, None)
        if isinstance(cls_attr, property):
            self.activate()
            if cls_attr.fget is None:
                raise AttributeError(f"property {name!r} has no getter")
            return cls_attr.fget(app)

        if not hasattr(app, name):
            raise AttributeError(
                f"{type(self).__name__!r} has no attribute {name!r}; "
                f"App also does not have it"
            )

        attr = getattr(app, name)

        # Distinguish REAL methods/functions from callable accessor
        # instances (MoveAccessor, _Actions etc. have __call__ in some
        # cases, but we should NOT wrap those — the user wants
        # `doc.move.top_of_file()`, not `doc.move()`).
        import inspect as _inspect
        is_routine = (
            _inspect.isroutine(attr) or                   # function/method/builtin
            _inspect.ismethoddescriptor(attr) or
            _inspect.isbuiltin(attr)
        )

        if is_routine:
            # Method → wrap with auto-activation
            def wrapped(*args, **kwargs):
                with app.use_document(self):
                    return attr(*args, **kwargs)
            wrapped.__name__ = name
            wrapped.__qualname__ = f"Document.{name} (proxied)"
            wrapped.__doc__ = (
                f"[Proxied from App.{name}] — 호출 시 이 문서를 활성창으로 "
                f"전환한 뒤 App.{name} 을 실행하고 원래 활성 문서로 복원합니다.\n\n"
                f"{getattr(attr, '__doc__', '') or ''}"
            )
            return wrapped

        # Accessor object (app.move, app.actions, app.cell, app.table,
        # app.page, app.documents) or plain attribute (app.api,
        # app.parameters): activate this doc FIRST so subsequent
        # calls on the returned object go to this doc's context.
        self.activate()
        return attr

    # ── 조회용 속성 ──────────────────────────────────────────

    @property
    def full_name(self) -> str:
        """전체 경로 포함 파일 이름. 저장 안된 문서는 빈 문자열."""
        try:
            return str(self.raw.FullName or "")
        except Exception:
            return ""

    @property
    def path(self) -> str:
        """문서 폴더 경로."""
        try:
            return str(self.raw.Path or "")
        except Exception:
            return ""

    @property
    def modified(self) -> bool:
        """마지막 저장 이후 수정 여부."""
        try:
            return bool(self.raw.Modified)
        except Exception:
            return False

    @property
    def saved(self) -> bool:
        """``modified`` 의 반대 — 저장된 상태인지 여부."""
        return not self.modified

    @property
    def name(self) -> str:
        """파일명 (경로 제외). 저장 안된 문서는 빈 문자열."""
        import os
        fn = self.full_name
        return os.path.basename(fn) if fn else ""

    @property
    def document_id(self):
        """HWP 내부 문서 ID."""
        try:
            return self.raw.DocumentID
        except Exception:
            return None

    @property
    def edit_mode(self):
        """편집 모드 (0=일반, 1=읽기 전용, ...)."""
        try:
            return self.raw.EditMode
        except Exception:
            return None

    @property
    def format(self):
        """문서 포맷 문자열 (HWP, HWPX, PDF, ...)."""
        try:
            return self.raw.Format
        except Exception:
            return None

    # ── 행동 ─────────────────────────────────────────────────

    def activate(self) -> "Document":
        """
        이 문서를 활성창으로 전환하고 self 반환.

        HWP 는 ``SetActive_XHwpDocument`` 호출 직후에도 내부 편집 컨텍스트가
        완전히 전환되기까지 짧은 지연이 필요합니다. 호출 직후 insert 나
        action 을 실행하면 이전 문서에 반영되는 현상을 피하기 위해 소폭의
        대기와 이벤트 처리 기회를 제공합니다.
        """
        import time as _time
        try:
            self.raw.SetActive_XHwpDocument()
        except Exception as e:
            self.logger.debug(
                f"activate: {type(e).__name__}: {e}",
                exc_info=True,
            )

        # Nudge message pump — pywin32 pumps messages via PumpWaitingMessages
        try:
            import pythoncom
            pythoncom.PumpWaitingMessages()
        except Exception as e:
            self.logger.debug(
                f"activate: {type(e).__name__}: {e}",
                exc_info=True,
            )
        _time.sleep(0.05)
        return self

    def save(self) -> bool:
        """현재 경로에 저장. 저장된 적 없으면 실패 (``save_as`` 사용)."""
        try:
            return bool(self.raw.Save())
        except Exception:
            return False

    def save_as(self, path: str, format: str = None) -> bool:
        """다른 이름/포맷으로 저장."""
        try:
            if format:
                return bool(self.raw.SaveAs(str(path), format))
            return bool(self.raw.SaveAs(str(path)))
        except Exception:
            return False

    def close(self, save: bool = False) -> bool:
        """이 문서만 닫기. ``save=True`` 면 저장 후 닫기."""
        try:
            return bool(self.raw.Close(save))
        except Exception:
            return False

    def clear(self) -> bool:
        """문서 내용 비우기 (새 문서처럼)."""
        try:
            return bool(self.raw.Clear())
        except Exception:
            return False

    def undo(self) -> bool:
        """이 문서의 되돌리기."""
        try:
            return bool(self.raw.Undo())
        except Exception:
            return False

    def redo(self) -> bool:
        """이 문서의 다시 실행."""
        try:
            return bool(self.raw.Redo())
        except Exception:
            return False

    def __repr__(self):
        name = self.full_name or "(unsaved)"
        mod = " *" if self.modified else ""
        return f"<Document {name}{mod}>"


class Documents:
    """
    열린 모든 HWP 문서를 감싼 컬렉션 (``app.documents``).

    파이썬 list 처럼 인덱싱/반복/`len()` 가능하며 ``add()``, ``open()``,
    ``close_all()``, ``active`` 같은 컬렉션 레벨 메소드를 제공합니다.

    Examples
    --------
    >>> # 전체 열린 문서 목록
    >>> for doc in app.documents:
    ...     print(doc.full_name, doc.modified)

    >>> # 새 빈 문서 추가
    >>> new_doc = app.documents.add()
    >>> new_doc.activate()

    >>> # 파일 열기 (기존 세션에 추가됨)
    >>> doc = app.documents.open("C:/report.hwp")

    >>> # 활성 문서
    >>> active = app.documents.active

    >>> # 전체 저장 후 닫기
    >>> app.documents.save_all()
    >>> app.documents.close_all()
    """

    def __init__(self, app):
        self._app = app

    @property
    def _raw(self):
        return self._app.api.XHwpDocuments

    def __len__(self) -> int:
        try:
            return int(self._raw.Count)
        except Exception:
            return 0

    def __getitem__(self, index: int) -> "Document":
        n = len(self)
        if index < 0:
            index += n
        if not (0 <= index < n):
            raise IndexError(f"Document index {index} out of range (count={n})")
        return Document(self._raw.Item(index), self)

    def __iter__(self):
        for i in range(len(self)):
            yield self[i]

    @property
    def active(self) -> "Document":
        """현재 활성 문서. ``app.api.XHwpDocuments.Active_XHwpDocument`` 기반."""
        try:
            raw = self._raw.Active_XHwpDocument
        except Exception:
            # Fallback — iterate and return the active one
            return self[0] if len(self) else None
        return Document(raw, self)

    def add(self, is_tab: bool = True) -> "Document":
        """새 빈 문서 추가. ``is_tab=True`` 면 같은 창의 새 탭."""
        try:
            raw = self._raw.Add(is_tab)
        except Exception:
            # Fallback: FileNew action
            self._app.api.Run("FileNew")
            raw = self._raw.Item(len(self) - 1)
        return Document(raw, self)

    def open(self, path: str, format: str = None,
             arg: str = "") -> "Document":
        """기존 파일을 새 문서로 열기."""
        try:
            if format:
                raw = self._raw.Open(str(path), format, arg)
            else:
                raw = self._raw.Open(str(path))
        except Exception:
            # Fallback: use app.open, then grab the newest doc
            self._app.open(str(path))
            raw = self._raw.Item(len(self) - 1)
        return Document(raw, self)

    def close_all(self, save: bool = False) -> int:
        """모든 문서를 닫고 닫은 개수를 반환."""
        closed = 0
        # Iterate from the end so indices don't shift while closing
        for i in range(len(self) - 1, -1, -1):
            try:
                if self[i].close(save=save):
                    closed += 1
            except Exception as e:
                self.logger.debug(
                    f"close_all: {type(e).__name__}: {e}",
                    exc_info=True,
                )
        return closed

    def save_all(self) -> int:
        """모든 문서를 저장하고 성공 개수를 반환."""
        saved = 0
        for doc in self:
            if doc.save():
                saved += 1
        return saved

    def find(self, name_substr: str) -> "Document":
        """파일 경로에 ``name_substr`` 이 포함된 첫 번째 문서 반환 (없으면 None)."""
        s = name_substr.lower()
        for doc in self:
            if s in doc.full_name.lower():
                return doc
        return None

    def __repr__(self):
        return f"<Documents count={len(self)}>"


