"""
:mod:`hwpapi.collections` — Phase 3 unified collection surface.

Every per-document aggregation exposed by :class:`hwpapi.document.Document`
(``fields``, ``bookmarks``, ``hyperlinks``, ``images``, ``styles``,
``paragraphs``, ``tables``) is implemented as a separate module in this
package. Each module wires straight to ``self._app.engine.impl`` — the
raw HWP COM handle — so unit tests can drive them with a MagicMock engine
and pywin32 is not required at import time.

The :class:`Collection` Protocol captures the shape every collection
implements. It is intentionally lightweight: structural equality against
six methods (``__getitem__``, ``__iter__``, ``__len__``, ``__contains__``,
``names``, ``filter``). The module is marked ``@runtime_checkable`` so
``isinstance(coll, Collection)`` does structural dispatch without
pulling in the generic ``TypeVar`` machinery that breaks older Python
runtimes.
"""
from __future__ import annotations

from typing import Callable, Iterator, Protocol, TypeVar, runtime_checkable

__all__ = ["Collection"]

T = TypeVar("T")


@runtime_checkable
class Collection(Protocol):
    """
    Protocol for every :mod:`hwpapi.collections` class.

    A *collection* is dict-like, iterable, sized, membership-testable,
    names-exposing and predicate-filterable. Methods intentionally
    match the v1 ``Fields`` surface so users see one shape everywhere.
    """

    def __getitem__(self, key):  # pragma: no cover - Protocol body
        ...

    def __iter__(self) -> Iterator:  # pragma: no cover
        ...

    def __len__(self) -> int:  # pragma: no cover
        ...

    def __contains__(self, key) -> bool:  # pragma: no cover
        ...

    def names(self) -> list[str]:  # pragma: no cover
        ...

    def filter(self, predicate: Callable[..., bool]) -> list:  # pragma: no cover
        ...
