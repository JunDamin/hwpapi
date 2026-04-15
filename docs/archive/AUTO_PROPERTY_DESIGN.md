# Auto-Creating Properties Design Specification

**Version:** 1.0
**Date:** 2025-01-08
**Status:** Design Phase

## Executive Summary

This document specifies the design and implementation of **auto-creating property descriptors** for hwpapi, enabling intuitive Pythonic access to HWP's nested parameter sets and arrays with full IDE tab completion support.

**Goals:**
- ✅ Eliminate manual `create_itemset()` calls
- ✅ Provide full IDE tab completion for nested properties
- ✅ Pythonic list interface for HWP arrays (HArray)
- ✅ Type-safe property access
- ✅ Backward compatible with existing code

---

## Problem Statement

### Current Pain Points

**1. Verbose Nested Parameter Set Creation:**
```python
# Current approach - 4 steps, no tab completion
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")  # Manual
char_shape = CharShape(char_com)  # Manual wrapping
char_shape.bold = True  # Finally set the value!
```

**2. No IDE Support:**
- Tab completion breaks after `create_itemset()`
- IDE doesn't know the type of `char_shape`
- No autocomplete for nested properties

**3. Array Handling is Primitive:**
- No Pythonic interface for HArray (COM arrays)
- Manual COM array manipulation required
- No list methods (append, insert, pop, etc.)

### Desired Developer Experience

```python
# Ideal approach - direct access, full tab completion
pset = FindReplace(action.CreateSet())
pset.find_char_shape.bold = True  # Auto-creates! IDE knows it's CharShape!
pset.find_para_shape.align = "center"  # Tab completion works!
pset.column_widths = [2000, 3000, 2500]  # Array just works like a list!
pset.column_widths.append(1500)  # Standard list methods!
```

---

## Solution Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────┐
│                    ParameterSet                         │
│  ┌───────────────────────────────────────────────────┐ │
│  │            Property Descriptors                   │ │
│  ├───────────────────────────────────────────────────┤ │
│  │  IntProperty      → Simple int values             │ │
│  │  BoolProperty     → Boolean values                │ │
│  │  StringProperty   → String values                 │ │
│  │  NestedProperty   → Auto-creating nested psets 🆕│ │
│  │  ArrayProperty    → Auto-creating arrays 🆕       │ │
│  └───────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
                           │
                           ├─ Nested: Creates via backend.create_itemset()
                           └─ Array:  Wraps HArray with HArrayWrapper
```

### Key Components

1. **NestedProperty** - Descriptor for nested ParameterSets
2. **ArrayProperty** - Descriptor for HArray COM arrays
3. **HArrayWrapper** - Pythonic list interface for HArray
4. **PsetBackend** - Backend supporting `create_itemset()`

---

## Technical Specification

### 1. NestedProperty Class

**Purpose:** Auto-create nested ParameterSet instances via `CreateItemSet` on first access.

**Class Definition:**
```python
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
```

**Behavior:**
- **Lazy creation:** Only creates when first accessed
- **Cached:** Subsequent access returns same instance
- **Type-safe:** Enforces correct ParameterSet class
- **IDE-friendly:** Return type is known at class definition time

### 2. ArrayProperty Class

**Purpose:** Provide Pythonic list interface for HWP's HArray (PIT_ARRAY) parameters.

**Class Definition:**
```python
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
```

### 3. HArrayWrapper Class

**Purpose:** Provide full Python list interface for HArray COM objects.

**Class Definition:**
```python
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
```

---

## Implementation Plan

### Phase 1: Core Implementation

**File:** `nbs/02_api/02_parameters.ipynb`

**Tasks:**
1. Add `NestedProperty` class after `TypedProperty` (cell ~19)
2. Add `ArrayProperty` class after `ListProperty` (cell ~20)
3. Add `HArrayWrapper` class after `ArrayProperty`
4. Add to `__all__` exports
5. Run `nbdev_export` to generate Python code

**Estimated LOC:** ~300 lines total

### Phase 2: Update Existing ParameterSet Classes

**Priority Classes to Update:**

1. **FindReplace** - Common use case
   ```python
   find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)
   find_para_shape = NestedProperty("FindParaShape", "ParaShape", ParaShape)
   ```

2. **Table** - Has both nested and arrays
   ```python
   border_fill = NestedProperty("BorderFill", "BorderFill", BorderFill)
   column_widths = ArrayProperty("ColWidths", int, "Column widths")
   ```

3. **TabDef** - Array example
   ```python
   tab_stops = ArrayProperty("TabStops", int, "Tab stop positions")
   ```

4. **BorderFill** - Nested in many places
   ```python
   # If has border widths array:
   border_widths = ArrayProperty("BorderWidth", int, "", min_length=4, max_length=4)
   ```

### Phase 3: Testing

**Test Cases:**

1. **NestedProperty Tests:**
   - Auto-creation on first access
   - Caching works (same instance returned)
   - Type enforcement on assignment
   - Works with both PsetBackend and HParamBackend
   - Error handling when not bound

2. **ArrayProperty Tests:**
   - List assignment works
   - Append/insert/remove methods work
   - Index access works
   - Length validation works
   - Type validation works
   - Iteration works

3. **Integration Tests:**
   - Complete workflow with real HWP actions
   - Nested + array in same ParameterSet
   - Tab completion in IDE (manual verification)

### Phase 4: Documentation

**Files to Update:**
- [x] `CLAUDE.md` - Comprehensive guide (DONE)
- [ ] `README.md` - Add examples
- [ ] `nbs/index.ipynb` - Update overview
- [ ] Docstrings in all new classes
- [ ] Example notebooks showing usage

---

## Migration Strategy

### Backward Compatibility

**Existing code continues to work:**
```python
# Old manual approach still works
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")
char_shape = CharShape(char_com)
char_shape.bold = True
```

**New auto-creating approach:**
```python
# New approach - simpler!
pset = FindReplace(action.CreateSet())
pset.find_char_shape.bold = True
```

### Gradual Migration Path

1. **Phase 1:** Add new property types (no breaking changes)
2. **Phase 2:** Update ParameterSet classes to use new properties
3. **Phase 3:** Deprecate manual `create_itemset()` pattern (optional)
4. **Phase 4:** Remove `TypedProperty` (far future, if ever)

### TypedProperty vs NestedProperty

**Keep both for now:**
- `TypedProperty` - Manual, requires `create_itemset()` call
- `NestedProperty` - Auto-creating, preferred for new code

**Deprecation (optional, future):**
```python
# Add warning to TypedProperty
warnings.warn(
    "TypedProperty is deprecated. Use NestedProperty for auto-creating nested sets.",
    DeprecationWarning
)
```

---

## Example Usage Scenarios

### Scenario 1: Find and Replace

```python
from hwpapi import App
from hwpapi.parametersets import FindReplace

app = App()

# Create parameter set
pset = app.actions.repeat_find.create_set()

# Simple properties
pset.find_string = "Python"
pset.replace_string = "Python 3.11"
pset.whole_word_only = True
pset.match_case = False

# Auto-creating nested properties
pset.find_char_shape.bold = True
pset.find_char_shape.fontsize = 1200  # 12pt
pset.find_char_shape.text_color = "#FF0000"

pset.find_para_shape.align = "center"
pset.find_para_shape.line_spacing = 160

# Execute
result = pset.run()
```

### Scenario 2: Table Creation

```python
from hwpapi import App
from hwpapi.parametersets import Table

app = App()

# Create table
table = app.actions.table_creation.create_set()

# Simple properties
table.rows = 5
table.cols = 4
table.has_header = True

# Array property
table.column_widths = [2000, 3000, 2500, 2000]  # One per column

# Modify array
table.column_widths[2] = 3500  # Adjust third column
table.column_widths.append(1500)  # Add fifth column

# Nested property
table.border_fill.border_type = "solid"
table.border_fill.fill_color = "#EEEEEE"
table.border_fill.border_widths = [10, 10, 10, 10]

# Execute
table.run()
```

### Scenario 3: Drawing Coordinates

```python
from hwpapi import App
from hwpapi.parametersets import DrawCoordInfo

app = App()

# Create drawing
coords = app.actions.insert_shape.create_set()

# Array of coordinate tuples
coords.points = [(0, 0), (100, 100), (200, 50), (300, 150)]

# Modify points
coords.points.append((400, 200))
coords.points.insert(2, (150, 75))

# Iterate over points
for i, (x, y) in enumerate(coords.points):
    print(f"Point {i}: ({x}, {y})")

# Execute
coords.run()
```

---

## Performance Considerations

### Caching Strategy

- **NestedProperty:** Caches wrapped instance in `_nested_cache_{key}` attribute
- **ArrayProperty:** Caches HArrayWrapper in `_array_cache_{key}` attribute
- **No redundant creation:** Subsequent access returns cached instance

### Memory Impact

- Minimal - only caches created instances
- Lazy creation - only creates what's accessed
- Cache cleared when ParameterSet garbage collected

### COM Performance

- `CreateItemSet` called once per nested property
- HArray operations buffered in local cache
- Sync to COM only on modification (not on every access)

---

## Success Criteria

**Must Have:**
- ✅ NestedProperty auto-creates via CreateItemSet
- ✅ ArrayProperty provides list interface
- ✅ Tab completion works in IDEs (VS Code, PyCharm)
- ✅ Backward compatible with existing code
- ✅ Type-safe property access
- ✅ Comprehensive tests pass

**Nice to Have:**
- ✅ Performance equal to or better than manual approach
- ✅ Clear error messages for common mistakes
- ✅ Examples in documentation
- ✅ Migration guide for existing code

---

## Risks and Mitigation

### Risk 1: COM Object Lifetime

**Risk:** Nested COM objects might be garbage collected unexpectedly

**Mitigation:**
- Cache wrapped instances in parent ParameterSet
- Parent ParameterSet holds reference to backend (which holds COM object)
- COM reference counting handles lifecycle

### Risk 2: HArray API Variations

**Risk:** Different HArray implementations might have different APIs

**Mitigation:**
- Use try/except blocks for COM operations
- Fall back to local cache if COM access fails
- Provide `to_list()` method for escape hatch

### Risk 3: Type Confusion

**Risk:** Users might not understand difference between NestedProperty and TypedProperty

**Mitigation:**
- Clear documentation
- Naming convention (`NestedProperty` suggests auto-creation)
- Deprecation warnings on old pattern (future)

---

## Future Enhancements

### Phase 5: Type Hints for IDE

```python
from typing import TypeVar, Generic

T = TypeVar('T', bound='ParameterSet')

class NestedProperty(Generic[T], PropertyDescriptor):
    def __get__(self, instance: "ParameterSet", owner) -> T:
        ...
```

### Phase 6: Validation Decorators

```python
class FindReplace(ParameterSet):
    @validate_length(min=1, max=255)
    find_string = StringProperty("FindString", "Text to find")
```

### Phase 7: Auto-Discovery from HWP Docs

```python
# Generate ParameterSet classes from official docs
generate_parameterset_from_docs("CharShape")
```

---

## Conclusion

The auto-creating property pattern provides:

1. **Intuitive API** - Direct property access, no manual creation
2. **IDE Support** - Full tab completion and type hints
3. **Type Safety** - Enforced at property definition time
4. **Backward Compatible** - Existing code continues to work
5. **Pythonic** - Feels natural for Python developers

This design aligns hwpapi with modern Python best practices while maintaining the powerful HWP automation capabilities.

---

**Document Status:** Ready for implementation
**Next Step:** Begin Phase 1 - Core Implementation in `02_parameters.ipynb`
