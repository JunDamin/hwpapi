"""
ParameterSet backends — abstraction over HWP COM / plain-Python object access.

Four backend types exist to support different underlying storage:

- ``PsetBackend``: Wraps pset COM objects from ``action.CreateSet()``.
  Uses ``Item()``, ``SetItem()``, ``CreateItemSet()`` natively. **Preferred.**

- ``HParamBackend``: Wraps ``HParameterSet`` COM objects with dotted path
  navigation (e.g. ``"HFindReplace.FindString"``). Legacy; used when
  working with global ``HParameterSet`` state.

- ``ComBackend``: Generic fallback for any COM object with
  Item/SetItem/RemoveItem methods.

- ``AttrBackend``: For plain Python objects (attributes, dicts).
  Used in tests and when no COM object is bound.

All backends implement the ``ParameterBackend`` protocol:
``get(key)``, ``set(key, value)``, ``delete(key) -> bool``.

See ``hwpapi/parametersets/__init__.py :: make_backend()`` for auto-selection.
"""
from __future__ import annotations
from typing import Any, Protocol


class ParameterBackend(Protocol):
    """Protocol for parameter backends."""
    def get(self, key: str) -> Any: ...
    def set(self, key: str, value: Any) -> None: ...
    def delete(self, key: str) -> bool: ...


class ComBackend:
    """win32com-style backend using Item/SetItem/RemoveItem."""
    def __init__(self, obj: Any):
        self._obj = obj

    def get(self, key: str) -> Any:
        return self._obj.Item(key)

    def set(self, key: str, value: Any) -> None:
        self._obj.SetItem(key, value)

    def delete(self, key: str) -> bool:
        try:
            self._obj.RemoveItem(key)
            return True
        except Exception:
            return False


class AttrBackend:
    """
    Fallback for plain Python objects:
    - attribute access if hasattr
    - dict-like access otherwise
    """
    def __init__(self, obj: Any):
        self._obj = obj

    def get(self, key: str) -> Any:
        if hasattr(self._obj, key):
            return getattr(self._obj, key)
        if isinstance(self._obj, dict):
            return self._obj.get(key, None)
        return self._obj[key]

    def set(self, key: str, value: Any) -> None:
        if hasattr(self._obj, key):
            setattr(self._obj, key, value)
        elif isinstance(self._obj, dict):
            self._obj[key] = value
        else:
            self._obj[key] = value

    def delete(self, key: str) -> bool:
        if hasattr(self._obj, key):
            try:
                delattr(self._obj, key)
                return True
            except Exception:
                return False
        if isinstance(self._obj, dict):
            return self._obj.pop(key, None) is not None
        try:
            del self._obj[key]
            return True
        except Exception:
            return False


class PsetBackend:
    """
    Backend for pset objects created via action.CreateSet().
    Supports direct pset operations: Item, SetItem, CreateItemSet.
    """
    def __init__(self, pset: Any):
        self._pset = pset

    def get(self, key: str) -> Any:
        """Get value using pset.Item(key)."""
        try:
            return self._pset.Item(key)
        except Exception as e:
            raise KeyError(f"Cannot get '{key}': {e}") from e

    def set(self, key: str, value: Any) -> None:
        """Set value using pset.SetItem(key, value)."""
        try:
            self._pset.SetItem(key, value)
        except Exception as e:
            raise KeyError(f"Cannot set '{key}': {e}") from e

    def delete(self, key: str) -> bool:
        """Delete item if supported by pset."""
        try:
            if hasattr(self._pset, 'RemoveItem'):
                self._pset.RemoveItem(key)
                return True
            return False
        except Exception:
            return False

    def create_itemset(self, key: str, setid: str) -> Any:
        """Create nested parameter set using pset.CreateItemSet(key, setid)."""
        try:
            return self._pset.CreateItemSet(key, setid)
        except Exception as e:
            raise KeyError(f"Cannot create itemset '{key}' with SetID '{setid}': {e}") from e

    def item_exists(self, key: str) -> bool:
        """Check if item exists using pset.ItemExist(key) if available."""
        try:
            if hasattr(self._pset, 'ItemExist'):
                return self._pset.ItemExist(key)
            self._pset.Item(key)
            return True
        except Exception:
            return False


class HParamBackend:
    """
    Backend for HParameterSet objects (legacy HSet-based approach).
    Supports dotted key paths for accessing nested parameter set structures.
    Example: "HFindReplace.FindString" navigates to HFindReplace.FindString attribute.
    """
    def __init__(self, root: Any):
        """Initialize with root HParameterSet or any child node."""
        self._root = root

    def get(self, key: str) -> Any:
        """Get value using dotted key path."""
        try:
            parent, leaf = self._resolve_parent_and_leaf(key)
            return getattr(parent, leaf)
        except (AttributeError, KeyError) as e:
            raise KeyError(f"Cannot get '{key}': {e}") from e

    def set(self, key: str, value: Any) -> None:
        """Set value using dotted key path with type-aware coercion."""
        try:
            parent, leaf = self._resolve_parent_and_leaf(key)
            coerced_value = self._coerce_for_put(parent, leaf, value)
            setattr(parent, leaf, coerced_value)
        except (AttributeError, KeyError) as e:
            raise KeyError(f"Cannot set '{key}': {e}") from e
        except TypeError as e:
            raise TypeError(f"Cannot set '{key}': {e}") from e

    def delete(self, key: str) -> bool:
        """
        Best-effort neutralization (not true deletion).
        Sets property to neutral value: 0/False/"" based on current type.
        """
        try:
            parent, leaf = self._resolve_parent_and_leaf(key)
            current = getattr(parent, leaf)

            if isinstance(current, bool):
                neutral = False
            elif isinstance(current, (int, float)):
                neutral = 0
            elif isinstance(current, str):
                neutral = ""
            else:
                neutral = None

            if neutral is not None:
                setattr(parent, leaf, neutral)
                return True
            return False
        except (AttributeError, KeyError):
            return False

    def _resolve_parent_and_leaf(self, key: str):
        """Resolve dotted key path to (parent_object, leaf_attribute)."""
        if '.' not in key:
            return (self._root, key)

        parts = key.split('.')
        current = self._root
        for part in parts[:-1]:
            current = getattr(current, part)
        return (current, parts[-1])

    def _coerce_for_put(self, parent: Any, leaf: str, value: Any) -> Any:
        """Type-aware coercion for HParameterSet attributes."""
        try:
            current = getattr(parent, leaf)
            current_type = type(current)

            if isinstance(value, current_type):
                return value

            if current_type == bool:
                return bool(value)
            elif current_type == int:
                return int(value)
            elif current_type == float:
                return float(value)
            elif current_type == str:
                return str(value)
            else:
                return value
        except (AttributeError, TypeError):
            return value


# ── COM object detection helpers ─────────────────────────────────────────

def _is_com(obj: Any) -> bool:
    """Check if object is a COM object."""
    if obj is None:
        return False
    return hasattr(obj, '_oleobj_') or 'com_gen_py' in str(type(obj))


def _looks_like_pset(obj: Any) -> bool:
    """
    Check if object looks like a pset created by action.CreateSet().

    Pset objects have methods like Item, SetItem, CreateItemSet, and SetID property.
    """
    if not _is_com(obj):
        return False

    pset_methods = ["Item", "SetItem", "CreateItemSet"]
    pset_properties = ["SetID"]

    for method in pset_methods:
        if not hasattr(obj, method):
            return False

    for prop in pset_properties:
        if not hasattr(obj, prop):
            return False

    return True


def make_backend(obj: Any) -> ParameterBackend:
    """
    Backend factory with pset priority.

    Selection order:
    1. pset objects (from action.CreateSet()) → PsetBackend
    2. COM objects → ComBackend
    3. Plain Python objects → AttrBackend
    """
    if _looks_like_pset(obj):
        return PsetBackend(obj)
    if _is_com(obj):
        return ComBackend(obj)
    return AttrBackend(obj)


__all__ = [
    "ParameterBackend",
    "ComBackend",
    "AttrBackend",
    "PsetBackend",
    "HParamBackend",
    "_is_com",
    "_looks_like_pset",
    "make_backend",
]
