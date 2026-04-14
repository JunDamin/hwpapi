"""
Property descriptor classes for ParameterSet.

Defines the 10+ descriptor types used by ParameterSet subclasses to map
Python-level attribute access to HWP COM pset Item/SetItem operations with
type-aware conversion and validation:

- PropertyDescriptor:  Base class (generic)
- IntProperty:         Integers with optional range validation
- BoolProperty:        Booleans ↔ 0/1
- StringProperty:      Strings
- ColorProperty:       #RRGGBB ↔ BBGGRR integer
- UnitProperty:        mm/cm/in/pt ↔ HWPUNIT
- MappedProperty:      String key ↔ int (via mapping dict)
- TypedProperty:       Generic typed (alias of PropertyDescriptor)
- NestedProperty:      Auto-creating nested ParameterSet
- ArrayProperty:       HArray wrapper (list-like COM)
- HArrayWrapper:       The Pythonic wrapper returned by ArrayProperty
- ListProperty:        Plain Python list

ParameterSet is referenced only at runtime (isinstance checks); the lazy
``_get_parameterset_cls()`` helper avoids circular import with ``base.py``.
"""
from __future__ import annotations
import re
from typing import Any, Dict, List, Optional, Union, Callable, Type, Iterable

from hwpapi.functions import from_hwpunit, to_hwpunit, convert_hwp_color_to_hex, convert_to_hwp_color
from .backends import PsetBackend, _looks_like_pset, _is_com, make_backend


def _get_parameterset_cls():
    """Lazy lookup of ParameterSet class to avoid circular import."""
    from .base import ParameterSet
    return ParameterSet


class MissingRequiredError(ValueError):
    """Raised when required parameters are missing during apply()."""
    pass


class PropertyDescriptor:
    """
    Descriptor for parameter properties backed by a staged ParameterSet.

    Features:
    - Reads prefer staged values, then snapshot, then default.
    - Writes stage the value; nothing is sent until ParameterSet.apply().
    - Optional automatic wrapping of nested ParameterSets via `wrap=...`.
    """

    @property
    def logger(self):
        from hwpapi.logging import get_logger
        return get_logger("parametersets.PropertyDescriptor")
    def __init__(
        self,
        key: str,
        doc: str = "",
        *,
        to_python: Optional[Callable[[Any], Any]] = None,
        to_backend: Optional[Callable[[Any], Any]] = None,
        default: Any = None,
        readonly: bool = False,
        required: bool = False,
        wrap: Optional[Type["ParameterSet"]] = None,  # <-- NEW: nested ParameterSet subclass
    ):
        self.key = key
        self.doc = doc
        self.name: Optional[str] = None
        self.to_python = to_python
        self.to_backend = to_backend
        self.default = default
        self.readonly = readonly
        self.required = required
        self.wrap = wrap

    def __set_name__(self, owner, name):
        self.name = name
        reg = getattr(owner, "_property_registry", None)
        if reg is None:
            reg = {}
            setattr(owner, "_property_registry", reg)
        reg[name] = self

    def __get__(self, instance: Optional["ParameterSet"], owner):
        if instance is None:
            return self

        # Get staged/snapshot value
        val = instance._ps_get(self)

        # Automatic wrapping for nested ParameterSets
        if self.wrap is not None:
            # Serve from cache if we already wrapped this key
            cached = instance._wrapper_cache.get(self.key)
            if cached is not None:
                return cached

            # If staged/snapshot holds an existing ParameterSet, cache and return it
            if isinstance(val, _get_parameterset_cls()):
                instance._wrapper_cache[self.key] = val
                return val

            # If backend returned a raw nested object (COM or dict-like), wrap it now
            if val is not None:
                # Handle both lambda functions (like lambda: CharShape) and direct class references
                if callable(self.wrap):
                    try:
                        # Try calling with no arguments first (for lambda: ClassName patterns)
                        wrapper_class = self.wrap()
                        wrapped = wrapper_class(val) if val != {} else wrapper_class()
                    except TypeError:
                        # If that fails, try calling with val as argument (for direct callable patterns)
                        wrapped = self.wrap(val)
                else:
                    # Direct class reference
                    wrapped = self.wrap(val) if val != {} else self.wrap()

                instance._wrapper_cache[self.key] = wrapped
                # Also stage the wrapper so reads stay consistent
                instance._staged[self.key] = wrapped
                return wrapped

            # If value is None, fall through to default handling below

        # Non-wrapped (or None) path - but for wrapped properties, create empty object
        if val is None:
            # For wrapped properties, return an empty instance instead of None
            if self.wrap is not None:
                # Create empty wrapped object and cache it
                if callable(self.wrap):
                    try:
                        wrapper_class = self.wrap()
                        wrapped = wrapper_class()
                    except TypeError:
                        wrapped = self.wrap()
                else:
                    wrapped = self.wrap()

                instance._wrapper_cache[self.key] = wrapped
                instance._staged[self.key] = wrapped
                return wrapped

            # For non-wrapped properties, return default if available
            if self.default is not None:
                return self.default

        # Auto-wrap pset objects using registry (Tier 1/2)
        if val is not None and _looks_like_pset(val):
            val = wrap_parameterset(val, self.key)

        return self.to_python(val) if (val is not None and self.to_python) else val

    def __set__(self, instance: "ParameterSet", value: Any):
        if self.readonly:
            raise AttributeError(f"'{self.name}' is read-only")

        # If this property is a nested ParameterSet, normalize on assignment:
        if self.wrap is not None:
            # Allow passing: ParameterSet, raw COM object, or dict-like
            if isinstance(value, _get_parameterset_cls()):
                wrapped = value
            else:
                # If dict/raw object given, create a wrapper
                wrapped = self.wrap(value if value is not None else {})
            # Keep cache consistent so subsequent gets reuse same object
            instance._wrapper_cache[self.key] = wrapped
            # Stage the wrapper (not the raw) — parent apply() will unwrap
            instance._ps_set(self, wrapped)
            return

        # Primitive path: run to_backend if provided
        v = self.to_backend(value) if (value is not None and self.to_backend) else value
        instance._ps_set(self, v)

    def _get_value(self, instance):
        return instance._ps_get(self)
    def _set_value(self, instance, value):
        instance._ps_set(self, value)
    def _del_value(self, instance): return instance._ps_del(self)


class IntProperty(PropertyDescriptor):
    """Property descriptor for integer values with optional range validation."""

    def __init__(self, key: str, doc: str, min_val: Optional[int] = None, max_val: Optional[int] = None):
        super().__init__(key, doc)
        self.min_val = min_val
        self.max_val = max_val

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return int(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, (int, float)):
            raise TypeError(f"Value for '{self.key}' must be numeric")

        value = int(value)

        if self.min_val is not None and value < self.min_val:
            raise ValueError(f"Value {value} for '{self.key}' is below minimum {self.min_val}")
        if self.max_val is not None and value > self.max_val:
            raise ValueError(f"Value {value} for '{self.key}' is above maximum {self.max_val}")

        return self._set_value(instance, value)


class BoolProperty(PropertyDescriptor):
    """Property descriptor for boolean values (0 or 1)."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return bool(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if isinstance(value, bool):
            numeric_value = 1 if value else 0
        elif isinstance(value, (int, float)):
            numeric_value = 1 if value else 0
        else:
            raise TypeError(f"Value for '{self.key}' must be boolean or numeric")

        return self._set_value(instance, numeric_value)


class StringProperty(PropertyDescriptor):
    """Property descriptor for string values."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return str(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, str):
            value = str(value)

        return self._set_value(instance, value)


class ColorProperty(PropertyDescriptor):
    """Property descriptor for color values with hex conversion."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return convert_hwp_color_to_hex(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        numeric_value = convert_to_hwp_color(value)
        return self._set_value(instance, numeric_value)


class UnitProperty(PropertyDescriptor):
    """
    Property descriptor for unit-based values with automatic conversion.

    Supports automatic conversion between user-friendly units (mm, cm, in, pt)
    and HWP's internal HWPUNIT format.

    Attributes:
        key (str): Parameter key in HWP
        doc (str): Documentation string
        default_unit (str): Default unit when bare numbers are provided (default: "mm")

    Example:
        >>> class PageDef(ParameterSet):
        ...     width = UnitProperty("Width", "Page width", default_unit="mm")
        >>> page = PageDef(action.CreateSet())
        >>> page.width = "210mm"   # String with unit
        >>> page.width = "21cm"    # Auto-converts to HWPUNIT
        >>> page.width = 210       # Bare number, assumes mm
        >>> print(page.width)      # Returns value in mm
        210.0
    """

    def __init__(self, key: str, doc: str, default_unit: str = "mm"):
        super().__init__(key, doc)
        self.default_unit = default_unit

    def __get__(self, instance, owner):
        """Get value in default unit."""
        if instance is None:
            return self
        value = self._get_value(instance)
        if value is None:
            return None
        # Convert from HWPUNIT to default unit
        return from_hwpunit(value, self.default_unit)

    def __set__(self, instance, value):
        """
        Set value from number or string with unit.

        Args:
            value: Number or string like "210mm", "21cm", "8.27in", "12pt"
        """
        if value is None:
            return self._del_value(instance)

        # Use enhanced to_hwpunit which handles both numbers and strings
        try:
            hwp_value = to_hwpunit(value, self.default_unit)
        except (ValueError, TypeError) as e:
            raise TypeError(
                f"Value for '{self.key}' must be numeric or unit string "
                f"(e.g., '210mm', '21cm', 210). Got: {value}"
            ) from e

        return self._set_value(instance, hwp_value)


class MappedProperty(PropertyDescriptor):
    """Property descriptor for mapped values (string <-> integer)."""

    def __init__(self, key: str, mapping: Dict[str, int], doc: str):
        super().__init__(key, doc)
        self.mapping = mapping
        self.reverse_mapping = {v: k for k, v in mapping.items()}

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        if value is None:
            return None
        return self.reverse_mapping.get(value, value)

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if isinstance(value, str):
            if value not in self.mapping:
                valid_keys = list(self.mapping.keys())
                raise ValueError(f"Invalid value '{value}' for '{self.key}'. Valid options: {valid_keys}")
            numeric_value = self.mapping[value]
        elif isinstance(value, (int, float)):
            numeric_value = int(value)
            if numeric_value not in self.reverse_mapping:
                valid_values = list(self.reverse_mapping.keys())
                raise ValueError(f"Invalid numeric value '{numeric_value}' for '{self.key}'. Valid values: {valid_values}")
        else:
            raise TypeError(f"Value for '{self.key}' must be string or numeric")

        return self._set_value(instance, numeric_value)


class TypedProperty(PropertyDescriptor):
    """
    Alias for PropertyDescriptor for typed parameter sets.
    Inherit all logic from PropertyDescriptor; do not override __get__ or __set__.
    Use this for clarity when defining typed sub-ParameterSets.
    """
    def __init__(self, key: str, doc: str = "", wrap=None, **kwargs):
        """Initialize TypedProperty with support for positional wrap argument."""
        super().__init__(key, doc, wrap=wrap, **kwargs)


    # Inherits __get__ and __set__ from PropertyDescriptor


class NestedProperty(PropertyDescriptor):
    """
    Auto-creating nested ParameterSet property.

    Automatically calls backend.create_itemset(key, setid) on first access,
    wraps the result in the specified ParameterSet class, and caches the instance.

    Attributes:
        key (str): Parameter key in HWP (e.g., "FindCharShape")
        setid (str): SetID for CreateItemSet call (e.g., "CharShape")
        param_class (Type[ParameterSet]): ParameterSet class to wrap
        doc (str): Documentation string
        _cache_attr (str): Internal cache attribute name

    Example:
        >>> class FindReplace(ParameterSet):
        ...     find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)
        >>> pset = FindReplace(action.CreateSet())
        >>> pset.find_char_shape.bold = True  # Auto-creates CharShape!
    """

    def __init__(self, key: str, setid: str, param_class: Type["ParameterSet"], doc: str = ""):
        super().__init__(key, doc)
        self.setid = setid
        self.param_class = param_class
        self._cache_attr = f"_nested_cache_{key}"

    def __get__(self, instance: "ParameterSet", owner) -> "ParameterSet":
        """
        Get nested ParameterSet, creating it if needed.

        Returns cached instance if available, otherwise:
        1. Calls backend.create_itemset(key, setid) to create COM object
        2. Wraps in param_class
        3. Caches for future access
        4. Returns wrapped instance

        Raises:
            RuntimeError: If ParameterSet not bound to backend
        """
        if instance is None:
            return self

        # Check cache first (subsequent access)
        if hasattr(instance, self._cache_attr):
            return getattr(instance, self._cache_attr)

        # Verify backend is available
        if instance._backend is None:
            raise RuntimeError(
                f"Cannot access nested property '{self.key}': "
                "ParameterSet not bound. Call .bind() or pass parameterset= to __init__"
            )

        # Create via CreateItemSet (first access)
        if hasattr(instance._backend, 'create_itemset'):
            # PsetBackend - use CreateItemSet
            nested_pset_com = instance._backend.create_itemset(self.key, self.setid)
            nested_wrapped = self.param_class(nested_pset_com)
        else:
            # Fallback for HParamBackend or other backends
            try:
                nested_com = instance._backend.get(self.key)
                nested_wrapped = self.param_class(nested_com)
            except (KeyError, AttributeError):
                # Create unbound instance as last resort
                nested_wrapped = self.param_class()

        # Cache for subsequent access
        setattr(instance, self._cache_attr, nested_wrapped)
        return nested_wrapped

    def __set__(self, instance: "ParameterSet", value: "ParameterSet"):
        """
        Allow direct assignment of ParameterSet instance.

        Args:
            instance: Parent ParameterSet
            value: ParameterSet instance to assign

        Raises:
            TypeError: If value is not the correct ParameterSet type
        """
        if not isinstance(value, self.param_class):
            raise TypeError(
                f"Value for '{self.key}' must be {self.param_class.__name__}, "
                f"got {type(value).__name__}"
            )

        # Cache the provided instance
        setattr(instance, self._cache_attr, value)

        # Stage for apply()
        instance._staged[self.key] = value

class ArrayProperty(PropertyDescriptor):
    """
    Auto-creating array property for HArray COM objects.

    Provides Pythonic list-like interface with automatic type conversion
    and optional length validation.

    Attributes:
        key (str): Parameter key in HWP
        item_type (Type): Type of array elements (int, float, str, tuple)
        doc (str): Documentation string
        min_length (Optional[int]): Minimum array length
        max_length (Optional[int]): Maximum array length
        _cache_attr (str): Internal cache attribute name

    Example:
        >>> class TabDef(ParameterSet):
        ...     tab_stops = ArrayProperty("TabStops", int, "Tab stop positions")
        >>> tab_def = TabDef(action.CreateSet())
        >>> tab_def.tab_stops = [1000, 2000, 3000]
        >>> tab_def.tab_stops.append(4000)
    """

    def __init__(self, key: str, item_type: Type, doc: str = "",
                 min_length: Optional[int] = None, max_length: Optional[int] = None):
        super().__init__(key, doc)
        self.item_type = item_type
        self.min_length = min_length
        self.max_length = max_length
        self._cache_attr = f"_array_cache_{key}"

    def __get__(self, instance: "ParameterSet", owner) -> "HArrayWrapper":
        """
        Get HArrayWrapper, creating it if needed.

        Returns:
            HArrayWrapper providing list-like interface to HArray
        """
        if instance is None:
            return self

        # Check cache first
        if hasattr(instance, self._cache_attr):
            return getattr(instance, self._cache_attr)

        # Try to get existing HArray from backend
        if instance._backend is not None:
            try:
                harray_com = instance._backend.get(self.key)
                if harray_com is not None:
                    wrapper = HArrayWrapper(harray_com, self.item_type,
                                           instance._backend, self.key)
                    setattr(instance, self._cache_attr, wrapper)
                    return wrapper
            except (KeyError, AttributeError):
                pass

        # Return empty wrapper (will create HArray on modification)
        wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key)
        setattr(instance, self._cache_attr, wrapper)
        return wrapper

    def __set__(self, instance: "ParameterSet", value: Union[List, Tuple, None]):
        """
        Set array from Python list/tuple.

        Args:
            value: Python list/tuple to set, or None to clear

        Raises:
            TypeError: If value is not list/tuple
            ValueError: If length constraints violated
        """
        if value is None:
            # Clear array
            wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key)
            setattr(instance, self._cache_attr, wrapper)
            instance._staged[self.key] = []
            return

        # Validate type
        if not isinstance(value, (list, tuple)):
            raise TypeError(
                f"Value for '{self.key}' must be list or tuple, "
                f"got {type(value).__name__}"
            )

        value_list = list(value)

        # Validate length
        if self.min_length is not None and len(value_list) < self.min_length:
            raise ValueError(
                f"Array '{self.key}' must have at least {self.min_length} items, "
                f"got {len(value_list)}"
            )
        if self.max_length is not None and len(value_list) > self.max_length:
            raise ValueError(
                f"Array '{self.key}' must have at most {self.max_length} items, "
                f"got {len(value_list)}"
            )

        # Create wrapper with initial values
        wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key,
                                initial_values=value_list)
        setattr(instance, self._cache_attr, wrapper)
        instance._staged[self.key] = value_list

class HArrayWrapper:
    """
    Pythonic wrapper around HWP's HArray COM object.

    Provides full list interface: indexing, iteration, append, insert, etc.
    Automatically syncs changes with underlying COM HArray when available.

    Attributes:
        _harray: COM HArray object (or None if not created yet)
        _item_type: Python type for array elements
        _backend: ParameterSet backend (for array creation)
        _key: Parameter key
        _local_cache: Python list holding current values
    """

    def __init__(self, harray_com: Any, item_type: Type,
                 backend: Optional[Any] = None, key: Optional[str] = None,
                 initial_values: Optional[List] = None):
        self._harray = harray_com
        self._item_type = item_type
        _backend = backend
        self._key = key
        self._local_cache = list(initial_values) if initial_values else []

        # Sync from COM if available
        if self._harray is not None:
            self._sync_from_com()

    def _sync_from_com(self):
        """Load values from COM HArray into local cache."""
        if self._harray is None:
            return
        try:
            count = self._harray.Count
            self._local_cache = [
                self._convert_from_com(self._harray.Item(i))
                for i in range(count)
            ]
        except Exception:
            pass  # Keep local cache if COM access fails

    def _sync_to_com(self):
        """Push local cache to COM HArray."""
        if self._harray is None:
            return
        try:
            # Clear existing
            while self._harray.Count > 0:
                self._harray.RemoveAt(0)
            # Add all from cache
            for value in self._local_cache:
                self._harray.Add(self._convert_to_com(value))
        except Exception:
            pass  # Keep local cache if sync fails

    def _convert_to_com(self, value: Any) -> Any:
        """Convert Python value to COM-compatible type."""
        if self._item_type == tuple:
            return value  # Tuples may need special handling
        return self._item_type(value)

    def _convert_from_com(self, value: Any) -> Any:
        """Convert COM value to Python type."""
        if self._item_type == tuple:
            return value
        return self._item_type(value)

    # List interface implementation
    def __len__(self) -> int:
        return len(self._local_cache)

    def __getitem__(self, index: int) -> Any:
        return self._local_cache[index]

    def __setitem__(self, index: int, value: Any):
        self._local_cache[index] = self._convert_to_com(value)
        self._sync_to_com()

    def __delitem__(self, index: int):
        del self._local_cache[index]
        self._sync_to_com()

    def __iter__(self):
        return iter(self._local_cache)

    def __repr__(self) -> str:
        return f"HArrayWrapper({self._local_cache})"

    def append(self, value: Any):
        """Add item to end of array."""
        self._local_cache.append(self._convert_to_com(value))
        self._sync_to_com()

    def insert(self, index: int, value: Any):
        """Insert item at index."""
        self._local_cache.insert(index, self._convert_to_com(value))
        self._sync_to_com()

    def remove(self, value: Any):
        """Remove first occurrence of value."""
        self._local_cache.remove(value)
        self._sync_to_com()

    def pop(self, index: int = -1) -> Any:
        """Remove and return item at index."""
        value = self._local_cache.pop(index)
        self._sync_to_com()
        return value

    def clear(self):
        """Remove all items."""
        self._local_cache.clear()
        self._sync_to_com()

    def to_list(self) -> List:
        """Convert to plain Python list."""
        return list(self._local_cache)

class ListProperty(PropertyDescriptor):
    """Property descriptor for list values."""

    def __init__(self, key: str, doc: str, item_type: Optional[Type] = None,
                 min_length: Optional[int] = None, max_length: Optional[int] = None):
        super().__init__(key, doc)
        self.item_type = item_type
        self.min_length = min_length
        self.max_length = max_length

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return list(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, (list, tuple)):
            raise TypeError(f"Value for '{self.key}' must be a list or tuple")

        value = list(value)

        # Length validation
        if self.min_length is not None and len(value) < self.min_length:
            raise ValueError(f"List for '{self.key}' must have at least {self.min_length} items")
        if self.max_length is not None and len(value) > self.max_length:
            raise ValueError(f"List for '{self.key}' must have at most {self.max_length} items")

        # Type validation
        if self.item_type is not None:
            for i, item in enumerate(value):
                if not isinstance(item, self.item_type):
                    if self.item_type == tuple and isinstance(item, (list, tuple)) and len(item) == 2:
                        value[i] = tuple(item)
                    else:
                        try:
                            value[i] = self.item_type(item)
                        except (ValueError, TypeError):
                            raise TypeError(f"Item {i} in list for '{self.key}' must be of type {self.item_type.__name__}")

        return self._set_value(instance, value)


