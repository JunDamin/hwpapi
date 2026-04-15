"""
Fields collection accessor (v1.0 일관성 청사진의 첫 단계).

`app.fields` 가 단순 list 였던 것을 풍부한 collection 으로 확장.
**기존 list-like 동작은 그대로 보존** — iteration 은 여전히 필드 이름
(str) 을 yield, ``in`` 연산자 / ``len()`` 도 동일.

추가된 collection 인터페이스::

    app.fields["name"]            # → Field 객체
    app.fields["name"] = "값"      # → put_field_text
    "name" in app.fields           # 기존 동작
    app.fields.add("name", memo="…", direction="…")
    app.fields.remove("name")
    app.fields.remove_all()
    app.fields.find("name")        # Optional[Field]
    app.fields.rename(old, new)
    app.fields.from_brackets(r"\\{\\{(\\w+)\\}\\}")
    list(app.fields)               # ['name', 'date', ...]  (legacy)
    for n in app.fields: ...       # 'name', 'date', ...    (legacy)

각 Field 는 값 객체:

    fld = app.fields["name"]
    fld.name                       # str
    fld.value                      # str (get/set)
    fld.value = "홍길동"
    fld.goto()                     # 커서 이동
    fld.remove()                   # 삭제
"""
from __future__ import annotations

import warnings
from typing import TYPE_CHECKING, Iterator, List, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


class Field:
    """단일 누름틀(필드) 값 객체."""

    __slots__ = ("_app", "name")

    def __init__(self, app: "App", name: str):
        self._app = app
        self.name = name

    @property
    def value(self) -> str:
        """현재 값."""
        return self._app.get_field(self.name)

    @value.setter
    def value(self, v) -> None:
        self._app.set_field(self.name, v)

    def goto(self, *, text: bool = True, front: bool = True,
             select: bool = False) -> bool:
        """커서를 필드 위치로 이동."""
        return self._app.move_to_field(
            self.name, text=text, front=front, select=select
        )

    def remove(self) -> bool:
        """필드 삭제."""
        return self._app.delete_field(self.name)

    def __repr__(self) -> str:
        try:
            v = self.value
            v_short = (v[:20] + "…") if len(v) > 20 else v
            return f"Field({self.name!r}, value={v_short!r})"
        except Exception:
            return f"Field({self.name!r})"


class Fields:
    """
    누름틀(필드) 컬렉션 — list-like + dict-like + collection methods.

    하위 호환을 위해 iterate 시 필드 이름 문자열을 yield.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    # ─── list-like (legacy compat) ────────────────────────────────

    def _names(self) -> List[str]:
        """Underlying list of field names (legacy `app.fields` 로직 재사용)."""
        # Use the raw _names_impl on App to avoid recursion if `app.fields`
        # gets aliased to this class
        raw = self._app.api.GetFieldList(0, 0) or ""
        for sep in ("\x02", "\t", "\n"):
            if sep in raw:
                names = raw.split(sep)
                break
        else:
            names = [raw] if raw else []
        seen = set()
        out = []
        for n in names:
            n = n.strip()
            if n and n not in seen:
                seen.add(n)
                out.append(n)
        return out

    def __iter__(self) -> Iterator[str]:
        """Yield field names (legacy 동작 — `for n in app.fields`)."""
        return iter(self._names())

    def __len__(self) -> int:
        return len(self._names())

    def __contains__(self, name) -> bool:
        if not isinstance(name, str):
            return False
        return self._app.field_exists(name)

    def __bool__(self) -> bool:
        return len(self) > 0

    # ─── dict-like (new) ──────────────────────────────────────────

    def __getitem__(self, key):
        """``app.fields["name"]`` → Field. ``app.fields[0]`` → 이름 (deprecated)."""
        if isinstance(key, int):
            warnings.warn(
                "Indexing app.fields with int is deprecated; "
                "use list(app.fields)[i] for a name, or app.fields['name'] "
                "for a Field object.",
                DeprecationWarning,
                stacklevel=2,
            )
            return self._names()[key]
        if isinstance(key, slice):
            warnings.warn(
                "Slicing app.fields is deprecated; use list(app.fields)[…].",
                DeprecationWarning,
                stacklevel=2,
            )
            return self._names()[key]
        if not isinstance(key, str):
            raise TypeError(f"Field key must be str, got {type(key).__name__}")
        if not self._app.field_exists(key):
            raise KeyError(f"Field {key!r} not found in document")
        return Field(self._app, key)

    def __setitem__(self, name, value) -> None:
        """``app.fields["name"] = "값"`` — 값 주입 (필요시 자동 생성)."""
        if not isinstance(name, str):
            raise TypeError(f"Field name must be str, got {type(name).__name__}")
        if not self._app.field_exists(name):
            self._app.create_field(name)
        self._app.set_field(name, value)

    def __delitem__(self, name) -> None:
        """``del app.fields["name"]``."""
        if not isinstance(name, str):
            raise TypeError(f"Field name must be str, got {type(name).__name__}")
        if not self._app.delete_field(name):
            raise KeyError(f"Field {name!r} not found or could not be deleted")

    # ─── collection methods (new) ─────────────────────────────────

    def add(self, name: str, *, memo: str = "",
            direction: str = "") -> Field:
        """
        현재 커서 위치에 새 필드 생성.

        Examples
        --------
        >>> app.fields.add("customer", direction="고객명")
        Field('customer', value='')
        """
        self._app.create_field(name, memo=memo, direction=direction)
        return Field(self._app, name)

    def remove(self, name: str) -> bool:
        """필드 삭제. 성공 시 True."""
        return self._app.delete_field(name)

    def remove_all(self) -> int:
        """모든 필드 삭제, 삭제된 개수 반환."""
        return self._app.delete_all_fields()

    def find(self, name: str) -> Optional[Field]:
        """이름으로 검색. 없으면 ``None``."""
        if self._app.field_exists(name):
            return Field(self._app, name)
        return None

    def rename(self, old: str, new: str) -> bool:
        """필드 이름 변경."""
        return self._app.rename_field(old, new)

    def from_brackets(self, pattern: str = r"\{\{(\w+)\}\}",
                      memo: str = "") -> List[str]:
        """
        ``{{name}}`` 형태 브래킷을 모두 필드로 변환. 변환된 이름 반환.

        :meth:`App.replace_brackets_with_fields` 의 별칭.
        """
        return self._app.replace_brackets_with_fields(pattern=pattern, memo=memo)

    def to_dict(self) -> dict:
        """``{name: value}`` 딕셔너리로 변환."""
        return {n: self._app.get_field(n) for n in self._names()}

    def update(self, mapping=None, **kwargs) -> None:
        """
        dict-style 일괄 업데이트.

        Examples
        --------
        >>> app.fields.update({"name": "홍길동", "date": "2026-04-15"})
        >>> app.fields.update(name="홍길동", date="2026-04-15")
        """
        if mapping:
            for k, v in dict(mapping).items():
                self[k] = v
        for k, v in kwargs.items():
            self[k] = v

    # ─── repr ─────────────────────────────────────────────────────

    def __repr__(self) -> str:
        names = self._names()
        if len(names) <= 5:
            preview = ", ".join(repr(n) for n in names)
        else:
            preview = ", ".join(repr(n) for n in names[:3]) + f", … (+{len(names) - 3})"
        return f"Fields([{preview}])"


class Bookmarks:
    """
    책갈피 컬렉션. (HWP COM 의 InsertBookMark / DeleteBookMark / SelectBookMark / ExistBookMark)

    Examples
    --------
    >>> app.bookmarks.add("ch1")
    >>> "ch1" in app.bookmarks
    True
    >>> app.bookmarks.goto("ch1")
    >>> app.bookmarks.remove("ch1")
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    def add(self, name: str) -> bool:
        """현재 커서 위치에 책갈피 생성."""
        return self._app.insert_bookmark(name)

    def remove(self, name: str) -> bool:
        """책갈피 삭제."""
        try:
            return bool(self._app.api.DeleteBookMark(name))
        except Exception:
            return False

    def goto(self, name: str) -> bool:
        """책갈피로 이동."""
        try:
            return bool(self._app.api.SelectBookMark(name))
        except Exception:
            return False

    def __contains__(self, name) -> bool:
        if not isinstance(name, str):
            return False
        try:
            return bool(self._app.api.ExistBookMark(name))
        except Exception:
            return False

    def __repr__(self) -> str:
        return f"Bookmarks(<app {id(self._app):x}>)"


class Hyperlink:
    """단일 하이퍼링크 값 객체."""

    __slots__ = ("text", "url")

    def __init__(self, text: str, url: str):
        self.text = text
        self.url = url

    def __repr__(self) -> str:
        return f"Hyperlink(text={self.text!r}, url={self.url!r})"


class Hyperlinks:
    """
    하이퍼링크 컬렉션.

    Examples
    --------
    >>> app.hyperlinks.add("Anthropic", "https://anthropic.com")
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    def add(self, text: str, url: str) -> Hyperlink:
        """현재 커서 위치에 하이퍼링크 삽입."""
        self._app.insert_hyperlink(text, url)
        return Hyperlink(text, url)

    def __repr__(self) -> str:
        return f"Hyperlinks(<app {id(self._app):x}>)"
