"""
:mod:`hwpapi.collections.fields` — FieldCollection (누름틀) for v2.

Wraps HWP's field (누름틀) API via the raw engine handle. Reads field
names through ``impl.GetFieldList``, values through ``impl.GetFieldText``
/ ``impl.PutFieldText``, and navigates via ``impl.MoveToField``.

All COM access goes through ``self._app.engine.impl`` so unit tests can
inject a MagicMock.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App

__all__ = ["Field", "FieldCollection"]


# HWP's GetFieldList returns names joined by 0x02 (STX). Fall back to
# common separators when the runtime returns something simpler.
_FIELD_SEPARATORS = ("\x02", "\t", "\n", ",")


def _split_field_list(raw: str) -> List[str]:
    """Split a raw HWP field list string on any known separator."""
    if not raw:
        return []
    for sep in _FIELD_SEPARATORS:
        if sep in raw:
            return [n.strip() for n in raw.split(sep) if n.strip()]
    return [raw.strip()] if raw.strip() else []


class Field:
    """Value object for a single HWP field (누름틀)."""

    __slots__ = ("_app", "name")

    def __init__(self, app: "App", name: str):
        self._app = app
        self.name = name

    @property
    def value(self) -> str:
        impl = self._app.engine.impl
        return str(impl.GetFieldText(self.name) or "")

    @value.setter
    def value(self, v) -> None:
        impl = self._app.engine.impl
        impl.PutFieldText(self.name, "" if v is None else str(v))

    def goto(self) -> bool:
        impl = self._app.engine.impl
        try:
            return bool(impl.MoveToField(self.name, True, True, False))
        except Exception:
            return False

    def __repr__(self) -> str:
        try:
            v = self.value
            preview = (v[:20] + "…") if len(v) > 20 else v
            return f"Field({self.name!r}, value={preview!r})"
        except Exception:
            return f"Field({self.name!r})"

    def __eq__(self, other) -> bool:
        if isinstance(other, Field):
            return self.name == other.name
        if isinstance(other, str):
            return self.name == other
        return NotImplemented

    def __hash__(self) -> int:
        return hash(("Field", self.name))


class FieldCollection:
    """
    ``doc.fields`` — collection of HWP fields (누름틀).

    Iterates as :class:`Field` value objects (not bare strings, unlike
    the v1 list-compat surface). Dict-like subscript by name; numeric
    subscript gives the nth :class:`Field` in document order.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App") -> None:
        self._app = app

    # ------------------------------------------------------------------
    # Core protocol
    # ------------------------------------------------------------------

    def _raw_names(self) -> List[str]:
        impl = self._app.engine.impl
        try:
            raw = impl.GetFieldList("") or ""
        except TypeError:
            # Some older bindings take (Number, Option)
            raw = impl.GetFieldList(0, 0) or ""
        except Exception:
            return []
        # Dedup preserving order; HWP lists duplicate entries per occurrence.
        seen: set[str] = set()
        out: List[str] = []
        for n in _split_field_list(str(raw)):
            if n not in seen:
                seen.add(n)
                out.append(n)
        return out

    def names(self) -> List[str]:
        """All field names, document order, deduplicated."""
        return self._raw_names()

    def __iter__(self) -> Iterator[Field]:
        for n in self._raw_names():
            yield Field(self._app, n)

    def __len__(self) -> int:
        return len(self._raw_names())

    def __contains__(self, key) -> bool:
        if not isinstance(key, str):
            return False
        return key in self._raw_names()

    def __getitem__(self, key) -> Field:
        names = self._raw_names()
        if isinstance(key, int):
            return Field(self._app, names[key])
        if isinstance(key, slice):
            return [Field(self._app, n) for n in names[key]]  # type: ignore[return-value]
        if isinstance(key, str):
            if key not in names:
                raise KeyError(f"Field {key!r} not found")
            return Field(self._app, key)
        raise TypeError(
            f"Field key must be str or int, got {type(key).__name__}"
        )

    def __setitem__(self, name, value) -> None:
        if not isinstance(name, str):
            raise TypeError(
                f"Field name must be str, got {type(name).__name__}"
            )
        impl = self._app.engine.impl
        impl.PutFieldText(name, "" if value is None else str(value))

    def __delitem__(self, name) -> None:
        raise NotImplementedError(
            "FieldCollection: field deletion requires a live HWP run; "
            "use doc.actions.FieldDelete or App.engine.impl.Run('FieldDelete')."
        )

    def filter(self, predicate: Callable[[Field], bool]) -> List[Field]:
        """Return the subset of fields for which ``predicate`` is truthy."""
        return [f for f in self if predicate(f)]

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def find(self, name: str) -> Optional[Field]:
        """Return the :class:`Field` with ``name`` or ``None``."""
        if name in self._raw_names():
            return Field(self._app, name)
        return None

    def get(self, name: str, default: str = "") -> str:
        """Return a field's current value, or ``default`` if missing."""
        if name not in self._raw_names():
            return default
        return Field(self._app, name).value

    def to_dict(self) -> dict:
        return {n: Field(self._app, n).value for n in self._raw_names()}

    def update(self, mapping=None, **kwargs) -> None:
        """Dict-style batch assignment."""
        if mapping:
            for k, v in dict(mapping).items():
                self[k] = v
        for k, v in kwargs.items():
            self[k] = v

    def __repr__(self) -> str:
        try:
            names = self._raw_names()
        except Exception:
            return "FieldCollection(<unbound>)"
        if len(names) <= 5:
            preview = ", ".join(repr(n) for n in names)
        else:
            preview = ", ".join(repr(n) for n in names[:3]) + f", … (+{len(names) - 3})"
        return f"FieldCollection([{preview}])"
