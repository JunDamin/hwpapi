"""
:mod:`hwpapi.collections.documents` — Documents collection (v3 surface).

xlwings 의 ``app.books`` 에 해당. ``app.docs.open(path)`` / ``add()`` 가
:class:`hwpapi.Document` 인스턴스를 반환하며, 사용자는 그 인스턴스에서
모든 작업을 수행합니다 (저장/닫기/insert_text 등).

ADR-003 참조 — v3 의 3-tier 구조 (App / docs / Document).
"""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Iterator, Optional, Union

if TYPE_CHECKING:
    from hwpapi.core.app import App
    from hwpapi.document import Document

__all__ = ["DocumentCollection"]


class DocumentCollection:
    """
    Documents collection — `app.docs` 로 접근.

    HWP 의 ``XHwpDocuments`` 를 wrap 한 컬렉션. 새 문서를 만들거나
    파일을 열면 ``Document`` 인스턴스를 반환합니다.

    Examples
    --------
    >>> from hwpapi import App
    >>> app = App()
    >>> doc = app.docs.open("report.hwp")     # → Document
    >>> doc.insert_text("Hello\\n")
    >>> doc.save()
    >>> doc.close()
    >>>
    >>> # 다중 문서
    >>> a = app.docs.open("a.hwp")
    >>> b = app.docs.add()
    >>> b.insert_text(a.text)
    >>> b.save_as("merged.hwp")
    >>> a.close(save=False); b.close()
    """

    def __init__(self, app: "App") -> None:
        self._app = app

    # ── factory ───────────────────────────────────────────────────

    def open(
        self,
        path: Union[str, Path],
        format: Optional[str] = None,
        arg: str = "",
    ) -> "Document":
        """파일을 열고 :class:`Document` 인스턴스 반환."""
        from hwpapi.functions import get_absolute_path
        from hwpapi.document import Document

        name = get_absolute_path(path)
        if format:
            self._app.api.Open(name, format, arg)
        else:
            self._app.api.Open(name)
        return Document(self._app, _raw=self._active_raw())

    def add(self) -> "Document":
        """새 빈 문서를 만들고 :class:`Document` 반환."""
        from hwpapi.document import Document

        try:
            self._app.api.Run("FileNew")
        except Exception:
            pass
        return Document(self._app, _raw=self._active_raw())

    # ── access ────────────────────────────────────────────────────

    @property
    def active(self) -> "Document":
        """현재 HWP 가 forefront 로 보여주는 문서."""
        from hwpapi.document import Document
        return Document(self._app, _raw=self._active_raw())

    def __len__(self) -> int:
        try:
            return int(self._app.api.XHwpDocuments.Count)
        except Exception:
            return 0

    def __iter__(self) -> Iterator["Document"]:
        from hwpapi.document import Document
        for i in range(len(self)):
            yield Document(self._app, _index=i)  # _index → resolves to _raw in __init__

    def __getitem__(self, key: Union[int, str]) -> "Document":
        from hwpapi.document import Document

        if isinstance(key, int):
            n = len(self)
            if key < 0:
                key += n
            if not (0 <= key < n):
                raise IndexError(f"document index {key} out of range (count={n})")
            return Document(self._app, _index=key)
        if isinstance(key, str):
            for i in range(len(self)):
                doc = Document(self._app, _index=i)  # _index → resolves to _raw in __init__
                if doc.name == key:
                    return doc
            raise KeyError(f"no open document named {key!r}")
        raise TypeError(f"key must be int or str, got {type(key).__name__}")

    def __contains__(self, key: Union[int, str]) -> bool:
        try:
            self[key]
            return True
        except (IndexError, KeyError):
            return False

    def __repr__(self) -> str:
        names = [d.name for d in self]
        return f"<DocumentCollection count={len(self)} names={names!r}>"

    # ── internal ──────────────────────────────────────────────────

    def _active_raw(self):
        """현재 활성 문서의 IXHwpDocument 핸들."""
        try:
            return self._app.api.XHwpDocuments.Active_XHwpDocument
        except Exception:
            return None
