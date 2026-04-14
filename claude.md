# Claude's Guide to Working with hwpapi

This document captures critical knowledge, patterns, and best practices for working with the hwpapi codebase. Follow these guidelines to work effectively and avoid common pitfalls.

---

## 🚨 Project Structure - Standard Python Package

### Source Files

Python files in `hwpapi/` are the **direct source code**. Edit them directly.

| File / Directory | Purpose |
|------------------|---------|
| `hwpapi/core.py` | App, Engine classes - main API |
| `hwpapi/actions.py` | HWP action wrappers (704 actions via __getattr__) |
| `hwpapi/parametersets/` | ParameterSet package (split from single file) |
| `hwpapi/parametersets/__init__.py` | ParameterSet base class, Meta, PropertyDescriptors, GenericParameterSet |
| `hwpapi/parametersets/mappings.py` | 35 string↔int MAPs (DIRECTION_MAP, ALIGN_MAP, etc.) |
| `hwpapi/parametersets/backends.py` | 4 backends (PsetBackend, HParamBackend, ComBackend, AttrBackend) |
| `hwpapi/parametersets/sets/primitives.py` | 8 base classes (CharShape, ParaShape, BorderFill, Cell, Caption, CtrlData, Password, Style) |
| `hwpapi/parametersets/sets/drawing.py` | 25 classes: ShapeObject, Draw* |
| `hwpapi/parametersets/sets/text.py` | 20 classes: character ops |
| `hwpapi/parametersets/sets/paragraph.py` | 4 classes: TabDef, NumberingShape, etc. |
| `hwpapi/parametersets/sets/table.py` | 12 classes: Table, CellBorderFill, etc. |
| `hwpapi/parametersets/sets/document.py` | 17 classes: DocumentInfo, PageDef, SecDef, etc. |
| `hwpapi/parametersets/sets/formatting.py` | 3 classes: BorderFillExt, StyleDelete, StyleTemplate |
| `hwpapi/parametersets/sets/file_ops.py` | 14 classes: FileOpen, Print, etc. |
| `hwpapi/parametersets/sets/find_edit.py` | 13 classes: FindReplace, BookMark, etc. |
| `hwpapi/parametersets/sets/media_misc.py` | 27 classes: OleCreation, HyperLink, Ftp*, etc. |
| `hwpapi/functions.py` | Utility functions |
| `hwpapi/classes.py` | Accessor classes (Move, Cell, Table, Page) |
| `hwpapi/constants.py` | Constants and enums |
| `hwpapi/logging.py` | Logging configuration |

See `docs/API_GUIDE.md` for the full API reference with recipes.

### Workflow for Code Changes

```bash
# 1. Edit the .py file directly
# 2. Test your changes
python -m pytest tests/

# 3. Commit
git add hwpapi/changed_file.py
git commit -m "Your change description"
```

### Note on nbs/ Directory

The `nbs/` directory contains archived Jupyter notebooks from the previous nbdev-based workflow. These are kept for reference only and are **not** the source of truth.

---

## 🏗️ Architecture Deep Dive

### Core Design Patterns

#### 1. Backend Abstraction Pattern

The codebase uses multiple backends to handle different parameter set types:

```python
# Backend hierarchy
ParameterBackend (Protocol)
├── PsetBackend      # For pset objects (preferred, modern)
├── HParamBackend    # For HParameterSet (legacy)
├── ComBackend       # For generic COM objects
└── AttrBackend      # For plain Python objects
```

**Key Functions:**
- `_is_com(obj)` - Checks if object is COM (has `_oleobj_` or 'com_gen_py')
- `_looks_like_pset(obj)` - Checks for pset-specific methods (Item, SetItem, CreateItemSet)
- `make_backend(obj)` - Factory that auto-detects and returns appropriate backend

**Important**: Backend selection is automatic. Trust the factory function.

#### 2. Property Descriptor System

Type-safe properties with automatic validation and conversion:

```python
class CharShape(ParameterSet):
    bold = BoolProperty("Bold", "Bold formatting")
    fontsize = IntProperty("Size", "Font size in points")
    text_color = ColorProperty("TextColor", "Text color")
```

**Property Types:**
- `IntProperty` - Integer values
- `BoolProperty` - Boolean values
- `StringProperty` - String values
- `ColorProperty` - Hex color ↔ HWP color conversion
- **`UnitProperty`** - 🆕 Smart unit conversion (mm, cm, in, pt ↔ HWPUNIT)
- `MappedProperty` - String ↔ Integer via mapping dict
- `TypedProperty` - Nested ParameterSet (manual)
- `ListProperty` - List of values (basic Python lists)
- **`NestedProperty`** - 🆕 Auto-creating nested ParameterSet with tab completion
- **`ArrayProperty`** - 🆕 Auto-creating HArray with list-like interface

**Auto-Generated Attributes:**
- `attributes_names` property returns `list(self._property_registry.keys())`
- ParameterSetMeta metaclass auto-populates `_property_registry`
- **NEVER** manually set `self.attributes_names = [...]` in subclasses

**Auto-Creating Properties (New Pattern):**
- `NestedProperty` and `ArrayProperty` automatically create underlying COM objects
- No manual `create_itemset()` or array initialization needed
- Full IDE tab completion and type hints
- Lazy creation on first access

#### 3. Staging vs Immediate Mode

**Two execution modes:**

1. **Pset-based (Modern, Preferred)**
   - Changes apply immediately
   - No staging required
   - Simpler mental model

2. **HSet-based (Legacy)**
   - Changes are staged first
   - Must call `apply()` to commit
   - Supports transactional changes

**Code Pattern:**
```python
# Check backend type
if isinstance(self._backend, PsetBackend):
    # Immediate mode
else:
    # Staging mode - accumulate in self._staged
```

---

## 🔍 Common Issues and Solutions

### Issue 1: Missing Function Definitions

**Symptom:** `NameError: name '_is_com' is not defined`

**Cause:** Helper function used but not defined in module

**Solution:**
```python
# Add the missing function in the appropriate .py file

def _is_com(obj: Any) -> bool:
    """Check if object is a COM object."""
    if obj is None:
        return False
    return hasattr(obj, '_oleobj_') or 'com_gen_py' in str(type(obj))
```

### Issue 2: AttributeError with attributes_names

**Symptom:** `AttributeError: property 'attributes_names' of 'X' object has no setter`

**Cause:** Trying to set `self.attributes_names = [...]` after it became a read-only property

**Solution:** Define properties instead of setting attributes_names:
```python
# ❌ OLD WAY (broken)
class MyPS(ParameterSet):
    def __init__(self):
        super().__init__()
        self.attributes_names = ["a", "b"]
        self.a = None
        self.b = None

# ✅ NEW WAY (correct)
class MyPS(ParameterSet):
    a = IntProperty("a", "Value a")
    b = IntProperty("b", "Value b")

    def __init__(self):
        super().__init__()
        # attributes_names auto-generated from properties
```

### Issue 3: Backend is None

**Symptom:** `AttributeError: 'NoneType' object has no attribute 'delete'`

**Cause:** Unbound ParameterSet (created without COM object)

**Solution:** Add None checks:
```python
def _del_value(self, name):
    """Legacy method - use backend instead."""
    if self._backend is None:
        return False
    return self._backend.delete(name)
```

### Issue 4: Import Errors Between Modules

**Symptom:** Circular import or missing imports

**Solution:** Check module imports and definition order:
- Functions must be defined/imported before use
- Use `from typing import TYPE_CHECKING` for type-only imports

### Issue 5: Duplicate Class/Method Definitions in Notebook

**Symptom:** Method doesn't behave as expected after edits; old logic still runs despite changes

**Cause:** Class or methods duplicated (common during copy-paste refactoring)

**How to Detect:**
```bash
# Count occurrences of method definition
python -c "with open('hwpapi/parametersets.py', encoding='utf-8') as f: print(f.read().count('def _format_int_value'))"
# Output: 2 (should be 1!)
```

**Prevention:**
- Use `grep -c "def method_name" hwpapi/parametersets.py` to count definitions
- Be careful with copy-paste

---

## 🛠️ Simplification Strategies

### Successfully Applied Simplifications

#### ✅ Auto-Generate attributes_names (Completed)

**Before:**
```python
class CharShape(ParameterSet):
    def __init__(self):
        super().__init__()
        self.attributes_names = [
            "facename_hangul", "facename_latin", ...,  # 67 lines!
        ]
```

**After:**
```python
# In ParameterSet base class
@property
def attributes_names(self):
    """Auto-generated list of attribute names from property registry."""
    return list(self._property_registry.keys())

# In subclasses - just define properties, no manual list!
class CharShape(ParameterSet):
    facename_hangul = StringProperty("FaceNameHangul", "...")
    facename_latin = StringProperty("FaceNameLatin", "...")
    # No self.attributes_names needed!
```

**Result:** Removed ~500 lines, eliminated maintenance burden

### Planned Simplifications (Priority Order)

#### Priority 1: Unify Backend Modes
- Remove dual staging behavior (immediate vs delayed)
- Pick one model and stick with it
- Impact: ~200 lines saved, 50% complexity reduction

#### Priority 2: Consolidate Property Types
- Replace 8 property classes with converter pattern
- Use `PropertyDescriptor("key", doc, converter=int)`
- Impact: ~200 lines saved, removes 6 classes

#### Priority 3: Remove Forward Declarations
- Use `TYPE_CHECKING` or reorder definitions
- Impact: ~25 lines saved, eliminates confusion

---

## 🧪 Testing Strategy

### Test Structure

```
tests/test_hparam.py
├── Unit tests (with mocks) - Run without HWP
├── Integration tests - Require HWP installed
└── Graceful skipping - Tests skip if dependencies unavailable
```

### Running Tests

```bash
# All tests
python -m pytest tests/test_hparam.py -v

# Specific test class
python -m pytest tests/test_hparam.py::TestParameterSetUpdateFrom -v

# Show skipped tests
python -m pytest tests/test_hparam.py -v -ra
```

### Test Requirements

**For Unit Tests:** Just Python + pytest
**For Integration Tests:**
- Windows OS
- HWP installed
- pywin32

### Writing New Tests

```python
import unittest
from hwpapi.parametersets import ParameterSet, IntProperty

class TestMyFeature(unittest.TestCase):
    def test_feature(self):
        # Use real ParameterSet subclasses that have properties defined
        from hwpapi.parametersets import CharShape

        ps = CharShape()
        ps.bold = True
        self.assertEqual(ps.bold, True)
```

**Important:** When testing ParameterSet, use classes with actual property descriptors, not manual attributes_names lists.

---

## 📝 Code Patterns to Follow

### Pattern 1: Adding a New ParameterSet Class

```python
class MyParameterSet(ParameterSet):
    """
    ### MyParameterSet

    123) MyParameterSet : 내 파라미터셋

    Description of what this parameter set does.
    """

    # Define properties (NOT attributes_names)
    my_int = IntProperty("MyInt", "Integer value")
    my_bool = BoolProperty("MyBool", "Boolean flag")
    my_color = ColorProperty("MyColor", "Color value")

    def __init__(self, parameterset=None, **kwargs):
        super().__init__(parameterset, **kwargs)
        # NO self.attributes_names = [...] needed!
        # Stage initial values if needed
        if 'my_int' in kwargs:
            self.my_int = kwargs['my_int']
```

### Pattern 2: Auto-Creating Nested ParameterSets

```python
class NestedProperty(PropertyDescriptor):
    """
    Auto-creating nested ParameterSet property.
    Automatically calls CreateItemSet when first accessed.

    Example:
        class FindReplace(ParameterSet):
            find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)

        pset = FindReplace(action.CreateSet())
        pset.find_char_shape.bold = True  # Auto-creates! Tab completion works!
    """

    def __init__(self, key: str, setid: str, param_class: Type["ParameterSet"], doc: str = ""):
        super().__init__(key, doc)
        self.setid = setid
        self.param_class = param_class
        self._cache_attr = f"_nested_cache_{key}"

    def __get__(self, instance, owner):
        if instance is None:
            return self

        # Check cache first
        if hasattr(instance, self._cache_attr):
            return getattr(instance, self._cache_attr)

        # Auto-create via CreateItemSet
        if instance._backend and hasattr(instance._backend, 'create_itemset'):
            nested_pset_com = instance._backend.create_itemset(self.key, self.setid)
            nested_wrapped = self.param_class(nested_pset_com)
        else:
            # Fallback: create unbound instance
            nested_wrapped = self.param_class()

        # Cache for future access
        setattr(instance, self._cache_attr, nested_wrapped)
        return nested_wrapped
```

**Usage:**
```python
class FindReplace(ParameterSet):
    find_string = StringProperty("FindString", "Text to find")
    find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)

pset = FindReplace(action.CreateSet())
pset.find_char_shape.bold = True  # Simple! Tab completion works!
```

### Pattern 3: Adding a Custom Property Type

```python
class MyProperty(PropertyDescriptor):
    """Custom property with special conversion."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        if value is None:
            return self.default
        # Your conversion logic
        return my_conversion(value)

    def __set__(self, instance, value):
        if value is None:
            self._del_value(instance)
            return
        # Your validation logic
        converted = my_conversion(value)
        self._set_value(instance, converted)
```

### Pattern 4: Checking COM Objects

```python
# Always use the helper function
if _is_com(obj):
    # Handle COM object

# For pset specifically
if _looks_like_pset(obj):
    # Handle pset object

# Let factory decide
backend = make_backend(obj)  # Automatic detection
```

### Pattern 5: Handling Optional Backend

```python
def my_method(self):
    """Method that accesses backend."""
    if self._backend is None:
        # Handle unbound case
        return default_value

    # Proceed with backend operations
    return self._backend.get(self.key)
```

---

## 🚀 Auto-Creating Properties: NestedProperty & ArrayProperty

### Overview

**Problem:** Manual nested parameter set creation is verbose and breaks tab completion:
```python
# ❌ Old way - too complicated!
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")
char_shape = CharShape(char_com)
char_shape.bold = True
```

**Solution:** Auto-creating properties that work like regular Python attributes:
```python
# ✅ New way - simple and intuitive!
pset = FindReplace(action.CreateSet())
pset.find_char_shape.bold = True  # Auto-creates! Tab completion works!
```

### NestedProperty - Auto-Creating Nested ParameterSets

**Purpose:** Automatically create nested parameter sets via `CreateItemSet` on first access.

**Signature:**
```python
NestedProperty(key: str, setid: str, param_class: Type[ParameterSet], doc: str = "")
```

**Parameters:**
- `key` - Parameter key in HWP (e.g., "FindCharShape")
- `setid` - SetID for CreateItemSet call (e.g., "CharShape")
- `param_class` - ParameterSet class to wrap (e.g., `CharShape`)
- `doc` - Documentation string

**Example Definition:**
```python
class FindReplace(ParameterSet):
    """Find and replace parameters."""

    # Simple properties
    find_string = StringProperty("FindString", "Text to find")

    # Auto-creating nested properties
    find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape,
                                     "Character formatting to find")
    find_para_shape = NestedProperty("FindParaShape", "ParaShape", ParaShape,
                                     "Paragraph formatting to find")
```

**Example Usage:**
```python
pset = app.actions.repeat_find.create_set()

# Access nested property - auto-creates CharShape via CreateItemSet!
pset.find_char_shape.bold = True
pset.find_char_shape.italic = False
pset.find_char_shape.text_color = "#FF0000"

# IDE provides full tab completion on find_char_shape!
# No manual create_itemset() call needed!
```

**How It Works:**
1. First access to `pset.find_char_shape` triggers `NestedProperty.__get__`
2. Calls `backend.create_itemset("FindCharShape", "CharShape")` to create COM object
3. Wraps result in `CharShape` Python class
4. Caches instance for future access
5. Returns fully-typed instance with all properties available

**Benefits:**
- ✅ **Tab completion** - IDE knows exact type and shows all properties
- ✅ **No manual creation** - CreateItemSet called automatically
- ✅ **Type safety** - Enforces correct ParameterSet class
- ✅ **Cached** - Subsequent access returns same instance
- ✅ **Lazy** - Only created when actually accessed

### UnitProperty - Smart Unit Conversion

**Purpose:** Automatically convert between user-friendly units (mm, cm, in, pt) and HWPUNIT.

**Problem:** HWPUNIT is not intuitive (1mm = 283 HWPUNIT, 1pt = 100 HWPUNIT)

**Solution:** Accept familiar units, auto-convert internally

**Signature:**
```python
UnitProperty(key: str, doc: str,
             default_unit: str = "mm",
             output_unit: Optional[str] = None,
             min_value: Optional[float] = None,
             max_value: Optional[float] = None)
```

**Example Definition:**
```python
class PageDef(ParameterSet):
    """Page layout."""

    # Dimensions in millimeters (most common for paper)
    width = UnitProperty("Width", "Page width", default_unit="mm")
    height = UnitProperty("Height", "Page height", default_unit="mm")

    # Margins in millimeters
    left_margin = UnitProperty("LeftMargin", "Left margin", default_unit="mm")

class CharShape(ParameterSet):
    """Character formatting."""

    # Font size in points (standard for typography)
    fontsize = UnitProperty("Height", "Font size", default_unit="pt")
```

**Example Usage:**
```python
# Page dimensions - ALL of these work!
page = PageDef(action.CreateSet())

# String with unit (most explicit)
page.width = "210mm"   # A4 width
page.height = "297mm"  # A4 height

# Different units (auto-converts)
page.width = "21cm"    # Same as 210mm
page.width = "8.27in"  # Same as 210mm

# Bare number (uses default_unit = mm)
page.width = 210       # Assumes mm

# Set margins with mixed units
page.left_margin = 25        # 25mm (bare number)
page.right_margin = "2.5cm"  # 25mm (converted)
page.top_margin = "1in"      # ~25.4mm (converted)

# Get value (returns in mm)
print(f"Width: {page.width}mm")  # Output: Width: 210.0mm

# Font size in points
char = CharShape(action.CreateSet())
char.fontsize = 12       # 12pt
char.fontsize = "12pt"   # Same
char.fontsize = "4.23mm" # Converts to pt internally
print(f"Font: {char.fontsize}pt")  # Output: Font: 12.0pt
```

**Supported Units:**
- `mm` - Millimeters (1mm = 283 HWPUNIT) - **Default for dimensions**
- `cm` - Centimeters (1cm = 2830 HWPUNIT)
- `in` - Inches (1in = 7200 HWPUNIT)
- `pt` - Points (1pt = 100 HWPUNIT) - **Default for fonts**

**Benefits:**
- ✅ **Intuitive** - Use familiar units (210mm instead of 59430 HWPUNIT)
- ✅ **Flexible** - String "210mm" or number 210 both work
- ✅ **Auto-converting** - Handles HWPUNIT internally
- ✅ **Validated** - Optional min/max in user units
- ✅ **Standard units** - mm for paper, pt for fonts

### ArrayProperty - Auto-Creating HArray with List Interface

**Purpose:** Provide Pythonic list interface for HWP's HArray (PIT_ARRAY) parameters.

**Signature:**
```python
ArrayProperty(key: str, item_type: Type, doc: str = "",
              min_length: Optional[int] = None, max_length: Optional[int] = None)
```

**Parameters:**
- `key` - Parameter key in HWP (e.g., "TabStops", "Point")
- `item_type` - Type of array elements (`int`, `float`, `str`, `tuple`, etc.)
- `doc` - Documentation string
- `min_length` - Minimum array length (optional validation)
- `max_length` - Maximum array length (optional validation)

**Example Definition:**
```python
class TabDef(ParameterSet):
    """Tab definition."""

    # Array of tab stop positions (in HWPUNIT)
    tab_stops = ArrayProperty("TabStops", int, "Tab stop positions")

class BorderFill(ParameterSet):
    """Border and fill settings."""

    # Array of 4 border widths: [left, right, top, bottom]
    border_widths = ArrayProperty("BorderWidth", int, "Border widths for each side",
                                  min_length=4, max_length=4)

class DrawCoordInfo(ParameterSet):
    """Drawing coordinates."""

    # Array of (X, Y) coordinate tuples
    points = ArrayProperty("Point", tuple, "Coordinate points")
```

**Example Usage:**
```python
# Tab stops
tab_def = TabDef(action.CreateSet())
tab_def.tab_stops = [1000, 2000, 3000, 4000]  # Set entire array
tab_def.tab_stops.append(5000)  # Standard list method
print(tab_def.tab_stops[0])  # Index access: 1000

# Border widths
border = BorderFill(action.CreateSet())
border.border_widths = [10, 10, 20, 20]  # left, right, top, bottom
border.border_widths[2] = 25  # Update top border

# Coordinates
coords = DrawCoordInfo(action.CreateSet())
coords.points = [(0, 0), (100, 100), (200, 50)]
coords.points.append((300, 75))
for i, (x, y) in enumerate(coords.points):
    print(f"Point {i}: ({x}, {y})")
```

**List-Like Methods:**
```python
# HArrayWrapper provides full list interface:
array.append(item)           # Add to end
array.insert(index, item)    # Insert at position
array.remove(item)           # Remove first occurrence
array.pop(index)             # Remove and return
array.clear()                # Remove all
array[index]                 # Get item
array[index] = value         # Set item
len(array)                   # Array length
for item in array: ...       # Iteration
```

**How It Works:**
1. Assignment `array = [...]` triggers `ArrayProperty.__set__`
2. Creates `HArrayWrapper` instance wrapping COM HArray
3. Wrapper provides list-like interface
4. All modifications sync to underlying HArray
5. Full Python list semantics

**Benefits:**
- ✅ **Pythonic** - Works exactly like Python lists
- ✅ **Tab completion** - IDE shows all list methods
- ✅ **Type validation** - Ensures all items match `item_type`
- ✅ **Length validation** - Optional min/max constraints
- ✅ **No COM knowledge needed** - Pure Python interface

### Complete Example: All Property Types Together

```python
class AdvancedTable(ParameterSet):
    """Table with all property types demonstrated."""

    # Simple properties
    rows = IntProperty("Rows", "Number of rows")
    cols = IntProperty("Cols", "Number of columns")
    has_header = BoolProperty("HasHeader", "First row is header")
    title = StringProperty("Title", "Table title")
    align = MappedProperty("Align", "Alignment", ALIGN_MAP)

    # Unit properties - AUTO-CONVERTING!
    table_width = UnitProperty("Width", "Table width", default_unit="mm")
    table_height = UnitProperty("Height", "Table height", default_unit="mm")

    # Array properties - AUTO-CREATING!
    column_widths = ArrayProperty("ColWidths", int, "Width of each column in HWPUNIT")
    row_heights = ArrayProperty("RowHeights", int, "Height of each row in HWPUNIT")

    # Nested property - AUTO-CREATING!
    border_fill = NestedProperty("BorderFill", "BorderFill", BorderFill,
                                 "Border and fill settings")

# Usage - everything just works!
table = AdvancedTable(action.CreateSet())

# Simple properties
table.rows = 3
table.cols = 4
table.has_header = True
table.title = "Sales Report"
table.align = "center"

# Unit properties (auto-converts to HWPUNIT)
table.table_width = "150mm"   # String with unit
table.table_height = 80       # Bare number, assumes mm
# OR use different units:
table.table_width = "15cm"    # Same as 150mm
table.table_width = "5.91in"  # Same as 150mm

# Array assignments (auto-creates HArray)
table.column_widths = [2000, 3000, 2500, 2000]  # HWPUNIT values
table.row_heights = [1000, 1000, 1000]

# Array modifications
table.column_widths.append(1500)
table.column_widths[2] = 3500

# Nested object access (auto-creates BorderFill via CreateItemSet)
table.border_fill.border_type = "solid"
table.border_fill.fill_color = "#EEEEEE"

# If border_fill has UnitProperty for border widths:
table.border_fill.border_left = "2mm"
table.border_fill.border_right = "0.2cm"  # Same as 2mm

# Execute
table.run()
```

### Migration Guide

#### From TypedProperty to NestedProperty

**Before:**
```python
class FindReplace(ParameterSet):
    find_char_shape = TypedProperty("FindCharShape", "Character formatting", CharShape)

# Usage - manual creation
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")
char_shape = CharShape(char_com)
char_shape.bold = True
```

**After:**
```python
class FindReplace(ParameterSet):
    find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape,
                                     "Character formatting")

# Usage - automatic!
pset = FindReplace(action.CreateSet())
pset.find_char_shape.bold = True  # Auto-creates!
```

**Migration Steps:**
1. Change `TypedProperty(key, doc, ParamClass)` to `NestedProperty(key, setid, ParamClass, doc)`
2. Add `setid` parameter (usually matches the class name)
3. Remove manual `create_itemset()` calls in usage code

#### From ListProperty to ArrayProperty

**Before:**
```python
class TabDef(ParameterSet):
    tab_stops = ListProperty("TabStops", "Tab positions", item_type=int)

# Usage - basic Python list (no COM sync)
tab_def = TabDef()
tab_def.tab_stops = [1000, 2000, 3000]
```

**After:**
```python
class TabDef(ParameterSet):
    tab_stops = ArrayProperty("TabStops", int, "Tab positions")

# Usage - syncs with HArray
tab_def = TabDef(action.CreateSet())
tab_def.tab_stops = [1000, 2000, 3000]  # Syncs to COM HArray
```

**Key Differences:**
- `ArrayProperty` requires binding to COM object (HArray)
- `ListProperty` is pure Python (no COM sync)
- Use `ArrayProperty` for HWP parameters that are PIT_ARRAY type
- Use `ListProperty` for internal Python-only lists

### Property Type Decision Tree

```
Does this parameter exist in HWP documentation?
├─ NO → Use regular Python attribute or ListProperty
└─ YES → What type is it?
    ├─ Simple value (int, bool, string) → Use IntProperty, BoolProperty, StringProperty
    ├─ Enum/mapped value → Use MappedProperty
    ├─ Nested ParameterSet → Use NestedProperty (auto-creating!)
    ├─ Array (PIT_ARRAY) → Use ArrayProperty (auto-creating!)
    ├─ Unit value (HWPUNIT) → Use UnitProperty (auto-converts mm/cm/in/pt!)
    └─ Color value → Use ColorProperty
```

**Unit Selection Guide:**
- **Page/table dimensions** → UnitProperty with `default_unit="mm"`
- **Margins/spacing** → UnitProperty with `default_unit="mm"`
- **Font size** → UnitProperty with `default_unit="pt"`
- **Border widths** → UnitProperty with `default_unit="mm"`
- **Line spacing** → UnitProperty with `default_unit="pt"` or `"mm"`

### Implementation Checklist

When adding auto-creating properties to a ParameterSet class:

**For NestedProperty:**
- [ ] Identify nested parameter sets in HWP documentation
- [ ] Find the `SetID` for `CreateItemSet` (usually matches class name)
- [ ] Import the nested ParameterSet class
- [ ] Define: `name = NestedProperty(key, setid, ParamClass, doc)`
- [ ] Test: `pset.name.some_property = value` works without manual creation

**For ArrayProperty:**
- [ ] Identify array parameters (PIT_ARRAY type in docs)
- [ ] Determine element type (int, float, str, tuple)
- [ ] Determine length constraints (if any)
- [ ] Define: `name = ArrayProperty(key, item_type, doc, min_length, max_length)`
- [ ] Test: `pset.name = [...]` and `pset.name.append(...)` work

### Best Practices

**DO ✅:**
1. Use `NestedProperty` for all nested ParameterSets (not `TypedProperty`)
2. Use `ArrayProperty` for HWP array parameters (not `ListProperty`)
3. Specify correct `setid` matching HWP documentation
4. Provide clear documentation strings
5. Add type hints for better IDE support
6. Test tab completion works in your IDE

**DON'T ❌:**
1. Don't use `TypedProperty` for new code (use `NestedProperty`)
2. Don't manually call `create_itemset()` when using `NestedProperty`
3. Don't use `ListProperty` for HWP array parameters (use `ArrayProperty`)
4. Don't forget to bind ParameterSet before accessing auto-creating properties
5. Don't mix up `key` and `setid` parameters

---

## 📺 ParameterSet Display Enhancements

### Overview

The `ParameterSet.__repr__()` method has been enhanced with three powerful features that create self-documenting, human-readable output. These enhancements work together to make debugging and learning much easier.

### Enhancement 1: Human-Readable Value Formatting

**Purpose:** Convert internal HWP values to intuitive, human-readable formats.

**Conversions:**

| Property Type | Internal Value | Display Format | Conversion |
|--------------|----------------|----------------|------------|
| **Colors** | `0x0000FF` (BBGGRR) | `#FF0000` | BBGGRR → #RRGGBB hex |
| **Font Sizes** | `1200` (HWPUNIT) | `12.0pt` | HWPUNIT ÷ 100 |
| **Dimensions** | `59430` (HWPUNIT) | `210.0mm` | via `from_hwpunit()` |
| **Booleans** | `True`/`False` | `True`/`False` | Direct display |

**Implementation:**
- `_format_int_value()` method detects property type
- Checks property name patterns: 'Size', 'Color', 'Width', 'Height', etc.
- Checks property descriptor type: `ColorProperty`, `UnitProperty`, etc.
- Applies appropriate conversion

**Example:**
```python
pset = CharFormat()
pset.FontSize = 1200
pset.TextColor = 0x0000FF
pset.Width = 59430

print(pset)
# Output:
# CharFormat(
#   FontSize=12.0pt
#   TextColor="#ff0000"
#   Width=210.0mm
# )
```

### Enhancement 2: Enum Display for MappedProperty

**Purpose:** Show both numeric value and mapped name for enum-like properties.

**Format:** `{numeric_value} ({mapped_name})`

**How It Works:**
1. Detects when property descriptor is `MappedProperty`
2. Retrieves raw numeric value from backend or staging dict
3. Gets mapped string name from property getter
4. Formats as `value (name)`

**Example:**
```python
class BookMark(ParameterSet):
    Type = MappedProperty("Type", {
        "일반책갈피": 0,
        "블록책갈피": 1
    }, "Bookmark type")

bookmark = BookMark()
bookmark.Type = "블록책갈피"  # Set using string name

print(bookmark)
# Output:
# BookMark(
#   Type=1 (블록책갈피)
#   ...
# )
```

**Benefits:**
- See internal numeric value HWP uses
- See human-readable name simultaneously
- Understand enum mappings without checking docs
- Works with any language (Korean, English, etc.)

**Common Use Cases:**
```python
# Search direction
Direction=0 (down)
Direction=1 (up)
Direction=2 (all)

# Text alignment
Align=0 (left)
Align=1 (center)
Align=2 (right)
Align=3 (justify)

# Bookmark types (Korean)
Type=0 (일반책갈피)
Type=1 (블록책갈피)
```

### Enhancement 3: Property Description Comments

**Purpose:** Display inline documentation for every property.

**Format:** `property=value  # description`

**How It Works:**
1. Checks if property descriptor has `doc` attribute
2. Appends as inline comment after the formatted value
3. Works with all property types

**Example:**
```python
class VideoInsert(ParameterSet):
    Base = StringProperty("Base", "동영상 파일의 경로")
    Format = MappedProperty("Format", {"mp4": 0, "avi": 1}, "동영상 형식")
    Width = IntProperty("Width", "동영상 너비 (HWPUNIT)")

video = VideoInsert()
video.Base = "C:/Videos/sample.mp4"
video.Format = "mp4"
video.Width = 59430

print(video)
# Output:
# VideoInsert(
#   Base="C:/Videos/sample.mp4"  # 동영상 파일의 경로
#   Format=0 (mp4)  # 동영상 형식
#   Width=210.0mm  # 동영상 너비 (HWPUNIT)
# )
```

**Benefits:**
- Self-documenting: No need to check external docs
- Units clarified: Know if it's HWPUNIT, pt, mm, etc.
- Format explained: Understand BBGGRR, enum values, ranges
- Context provided: Hints, constraints, valid values
- Multilingual: Works with Korean and English descriptions

### All Three Enhancements Together

**Complete Example:**
```python
class CharFormat(ParameterSet):
    FontName = StringProperty("FontName", "Font family name")
    FontSize = IntProperty("FontSize", "Font size in HWPUNIT (100 = 1pt)")
    TextColor = ColorProperty("TextColor", "Text color in BBGGRR format")
    Bold = BoolProperty("Bold", "Bold formatting")
    Underline = MappedProperty("Underline", {
        "none": 0, "single": 1, "double": 2
    }, "Underline style")

char = CharFormat()
char.FontName = "Arial"
char.FontSize = 1200
char.TextColor = 0x0000FF
char.Bold = True
char.Underline = "single"

print(char)
# Output:
# CharFormat(
#   Bold=True  # Bold formatting
#   FontName="Arial"  # Font family name
#   FontSize=12.0pt  # Font size in HWPUNIT (100 = 1pt)
#   TextColor="#ff0000"  # Text color in BBGGRR format
#   Underline=1 (single)  # Underline style
#   [staged changes: 5]
# )
```

**Notice:**
1. `12.0pt` - Human-readable value (Enhancement 1)
2. `#ff0000` - Color converted to hex (Enhancement 1)
3. `1 (single)` - Enum shows value + name (Enhancement 2)
4. `# Font size...` - Description explains everything (Enhancement 3)

### Implementation Details

**Location:** `ParameterSet._format_repr()` method in `hwpapi/parametersets.py`

**Key Methods:**
```python
def __repr__(self):
    """Return human-readable representation."""
    return self._format_repr()

def _format_repr(self, indent=0, max_depth=3):
    """Format ParameterSet with all enhancements."""
    # 1. Get all properties from registry
    # 2. Format each value based on type
    # 3. Add enum display for MappedProperty
    # 4. Append description comment
    # 5. Return complete formatted string

def _format_int_value(self, prop_name, prop_descriptor, value):
    """Format integer values based on property type."""
    # Detect colors, sizes, dimensions
    # Apply appropriate conversion
    # Return formatted string
```

**Testing:**
```bash
# Run demos
python examples/nested_property_demo.py
python examples/mapped_property_display_demo.py
python examples/property_description_display_demo.py
```

### Best Practices for Property Definitions

**DO ✅:**
```python
# Provide clear, informative descriptions
FontSize = IntProperty("FontSize", "Font size in HWPUNIT (100 = 1pt)")
Width = IntProperty("Width", "Table width in HWPUNIT (283 = 1mm)")
Rows = IntProperty("Rows", "Number of rows (1-500)")

# Include units, formats, ranges in descriptions
TextColor = ColorProperty("TextColor", "Text color in BBGGRR format")
Direction = MappedProperty("Direction", {...}, "Search direction (down=forward, up=backward)")

# Use descriptive enum values
Align = MappedProperty("Align", {
    "left": 0,
    "center": 1,
    "right": 2
}, "Text alignment on page")
```

**DON'T ❌:**
```python
# Don't leave descriptions empty
FontSize = IntProperty("FontSize", "")  # No help for users!

# Don't omit units/constraints
Width = IntProperty("Width", "Width")  # Width in what unit?

# Don't use cryptic enum values
Mode = MappedProperty("Mode", {
    "m1": 0,  # What is m1?
    "m2": 1   # What is m2?
}, "Mode")
```

### Benefits Summary

**For Users:**
- ✅ Understand parameters without checking docs
- ✅ See values in familiar units (pt, mm, #RRGGBB)
- ✅ Learn API while debugging
- ✅ Verify correct values are being set

**For Developers:**
- ✅ Self-documenting code
- ✅ Easier debugging
- ✅ Better error messages possible
- ✅ Reduced support questions

**For Documentation:**
- ✅ Examples show real, understandable values
- ✅ Screenshots are more informative
- ✅ API is more discoverable

---

## 🎯 Best Practices

### DO ✅

1. **Edit .py files directly** (this is a standard Python project)
2. **Test after every change** (at least run imports)
4. **Use property descriptors** for new ParameterSet attributes
5. **Check for None backend** in methods that access it
6. **Trust the backend factory** (make_backend)
7. **Follow existing patterns** in similar code
8. **Use type hints** (already set up with `from __future__ import annotations`)
10. **Keep backward compatibility** when refactoring

### DON'T ❌

1. **DON'T set `self.attributes_names = [...]` in subclasses**
3. **DON'T assume backend is always present** (can be None)
4. **DON'T mix staging modes** without understanding
5. **DON'T add features without tests**
6. **DON'T break existing API** without migration path
7. **DON'T use `isinstance` checks** unless necessary
10. **DON'T create new COM detection logic** (use `_is_com`)

---

## 🚀 Development Workflow

### Making a Change (Step by Step)

```bash
# 1. Identify what needs changing
# 2. Edit the .py file directly
# 3. Test your changes
python -c "import hwpapi; from hwpapi.parametersets import CharShape"
python -m pytest tests/test_hparam.py -v

# 4. Test in actual use (if possible)
python examples/your_example.py

# 5. Commit
git add hwpapi/changed_file.py
git commit -m "Description of change"
```

---

## 📚 Key Reference Information

### Important Files

| File | Purpose |
|------|---------|
| `pyproject.toml` | Project configuration, version, dependencies |
| `claude.md` | This file - working guidelines |
| `PSET_MIGRATION_SUMMARY.md` | Context on pset refactoring |
| `REFACTORING_SUMMARY.md` | Recent refactoring documentation |
| `DUPLICATE_FIX_SUMMARY.md` | Duplicate class bug fix and display formatting (2025-12-09) |
| `AUTO_PROPERTY_DESIGN.md` | NestedProperty & ArrayProperty design |
| `UNIT_PROPERTY_ENHANCEMENT.md` | Smart unit conversion specification |

### Key Classes

| Class | Location | Purpose |
|-------|----------|---------|
| `App` | `core.py` | High-level API for users |
| `Engine` | `core.py` | Mid-level wrapper around HwpObject |
| `ParameterSet` | `parametersets.py` | Base class for all parameter sets |
| `PropertyDescriptor` | `parametersets.py` | Base for property types |
| `ParameterSetMeta` | `parametersets.py` | Metaclass for auto-registration |

### Key Functions

| Function | Location | Purpose |
|----------|----------|---------|
| `_is_com(obj)` | `parametersets.py` | Check if object is COM |
| `_looks_like_pset(obj)` | `parametersets.py` | Check if object is pset |
| `make_backend(obj)` | `parametersets.py` | Create appropriate backend |
| `resolve_action_args()` | `parametersets.py` | Resolve action arguments |

### Environment Variables

```bash
# Logging Configuration
HWPAPI_LOG_LEVEL=DEBUG      # Set logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
                            # Default: WARNING (production-friendly, only shows warnings/errors)
                            # Use DEBUG or INFO for development/troubleshooting

# Examples:
# Development - show all logs
export HWPAPI_LOG_LEVEL=DEBUG

# Production - only warnings and errors (default if not set)
export HWPAPI_LOG_LEVEL=WARNING

# Quiet mode - only errors and critical
export HWPAPI_LOG_LEVEL=ERROR
```

**Important**: The default log level is `WARNING`, which means normal users only see warnings, errors, and critical messages. This is intentional to avoid cluttering output in production. Set `HWPAPI_LOG_LEVEL=DEBUG` or `INFO` when you need detailed logging for development or troubleshooting.

---

## 🐛 Debugging Tips

### Issue: Import errors

```python
# Check what's exported
python -c "import hwpapi.parametersets; print(dir(hwpapi.parametersets))"

# Check __all__ in generated file
grep "__all__" hwpapi/parametersets.py
```

### Issue: Test failures

```bash
# Run with verbose output
python -m pytest tests/test_hparam.py -vv -s

# Run specific test
python -m pytest tests/test_hparam.py::TestClass::test_method -vv

# Show full traceback
python -m pytest tests/test_hparam.py --tb=long
```

### Issue: Duplicate definitions in code

```bash
# Quick check for duplicate method names in generated file
grep -c "def _format_int_value" hwpapi/parametersets.py  # Should be 1

# Find all duplicate methods in a file
python << 'EOF'
import re
from collections import Counter

with open('hwpapi/parametersets.py', encoding='utf-8') as f:
    content = f.read()

# Find all method definitions
methods = re.findall(r'\n    def (\w+)', content)
method_counts = Counter(methods)

# Show duplicates
duplicates = {name: count for name, count in method_counts.items() if count > 1}
if duplicates:
    print("DUPLICATES FOUND:")
    for name, count in duplicates.items():
        print(f"  {name}: {count} times")
else:
    print("No duplicates found.")
EOF

# After fixing, verify
grep -c "def _format_int_value" hwpapi/parametersets.py  # Should be 1
```

---

## 💡 Lessons Learned

### From Recent Simplifications

1. **Auto-generation beats manual maintenance**
   - `attributes_names` property eliminated 500+ lines
   - No sync issues between properties and attribute lists
   - Single source of truth (property registry)

2. **Property system is powerful**
   - Declarative is better than imperative
   - Type safety and validation in one place
   - Easier to extend and maintain

3. **Edge cases matter**
   - Always check for None backend
   - Handle unbound ParameterSets
   - Tests catch these issues

4. **Simplification requires careful testing**
   - Changed behavior can break tests
   - Update tests to match new patterns
   - Verify with real usage, not just unit tests

6. **Human-readable display is valuable**
   - Raw HWPUNIT/BBGGRR values are not intuitive
   - Smart formatting (`_format_int_value`) makes debugging easier
   - `__repr__` showing properties helps users understand ParameterSet state
   - Context-aware formatting: colors as hex, sizes as pt, dimensions as mm

7. **Duplicate detection is critical after refactoring**
   - Count method definitions: `grep -c "def method_name" file.py`
   - Duplicates can silently override correct implementations
   - Second definition always wins in Python class definitions

### Pitfalls Encountered

1. ❌ Setting `self.attributes_names` → AttributeError (property has no setter)
3. ❌ Missing `_is_com` definition → NameError
4. ❌ Not checking for None backend → AttributeError
5. ❌ Manual attribute lists out of sync → Runtime errors
6. ❌ Duplicate class/method definitions → Second definition overrides first, causing bugs
7. ❌ Copy-paste refactoring without checking for duplicates → Hard-to-debug issues

---

## 🎓 Understanding the Domain

### HWP (Hancom Office)
- Korean word processor (like MS Word for Korea)
- COM automation via `HwpObject`
- Actions executed via `Run()` with parameter sets

### win32com Interface
- PyWin32 provides COM bridge
- COM objects have `_oleobj_` attribute
- Generated COM classes have 'com_gen_py' in type string

### Parameter Sets
- Configure HWP actions (like "InsertText", "FindReplace")
- Two flavors: pset (modern) and HSet (legacy)
- Properties map Python names to COM property names

---

## 📊 Codebase Metrics

**Current State:**
- Total lines: ~15,000
- parametersets.py: ~4,100 lines
- ParameterSet subclasses: 29
- Property descriptors: 438
- Action definitions: 899+

**After Priority 2 Simplification:**
- Estimated: ~500 lines removed
- Maintenance burden: Significantly reduced
- Complexity: Much lower

---

## 🏛️ HWP Object Model: Official vs Current Architecture

### Overview

This section compares the **official HWP Automation Object Model** (from HwpAutomation_2504.pdf) with the **current hwpapi implementation** to identify gaps, misalignments, and opportunities for better code organization.

**Documentation Source:** `hwp_docs/HwpAutomation_2504.pdf` (Korean, dated 2025-04-15)

---

### Official HWP Object Model Structure

The official HWP automation follows a **hierarchical object model** similar to Microsoft Office:

```
IHwpObject (Root COM Object)
│
├── IXHwpDocuments (Collection)
│   └── IXHwpDocument (Single)
│       ├── Properties: FullName, Name, Path, Saved, etc.
│       └── Methods: Save(), SaveAs(), Close(), Print(), etc.
│
├── IXHwpWindows (Collection)
│   └── IXHwpWindow (Single)
│       ├── Properties: Width, Height, Left, Top, Active, etc.
│       └── Methods: Activate(), Close(), etc.
│
├── IXHwpForms (Collection)
│   └── Form Controls (Various types)
│       ├── IXHwpFormPushButtons (Collection)
│       ├── IXHwpFormCheckButtons (Collection)
│       ├── IXHwpFormRadioButtons (Collection)
│       ├── IXHwpFormComboBoxes (Collection)
│       └── etc.
│
├── HAction (Action Execution System)
│   ├── GetActionIDByName(name) → ActionID
│   ├── Run(ActionID)
│   └── Execute(ActionID, ParameterSet)
│
├── HParameterSet (Parameter Management)
│   ├── CreateItemSet(SetID, ParamIndex) → Creates nested parameter set
│   ├── Item(ParamIndex) → Get parameter value
│   ├── SetItem(ParamIndex, Value) → Set parameter value
│   └── Clear() → Clear all parameters
│
├── HSet (Parameter Collection - Legacy)
│   └── Collection of parameters for complex actions
│
└── HArray (Parameter Arrays - PIT_ARRAY type)
    ├── Count → Number of elements
    ├── Item(index) → Get element at index
    ├── SetItem(index, value) → Set element at index
    ├── Add(value) → Append element
    └── RemoveAt(index) → Remove element at index
```

**Key Characteristics:**
1. **Collection Pattern**: Documents, Windows, Forms follow "Collection → Single Object" pattern
2. **Hierarchical Navigation**: Document → Sections → Paragraphs → Characters
3. **Action System**: Centralized via HAction.Execute() with HParameterSet
4. **900+ Actions**: Each with specific parameter requirements
5. **Type-Safe Parameters**: Strongly typed via HParameterSet interface

---

### Current hwpapi Architecture

The current implementation uses a **wrapper-based approach** with custom patterns:

```
App (Main Entry Point)
│
├── Engine
│   └── impl (HwpObject COM object)
│       └── Direct COM access: self.api.MovePos(), self.api.Run(), etc.
│
├── _Actions (900+ actions as properties)
│   ├── CharShape → _Action("CharShape", CharShape parameterset)
│   ├── ParaShape → _Action("ParaShape", ParaShape parameterset)
│   └── [899+ more actions...]
│
├── ParameterSet System (130+ classes in parametersets.py)
│   ├── Base: ParameterSet, ParameterSetMeta
│   ├── Backend Abstraction:
│   │   ├── PsetBackend (modern, immediate)
│   │   ├── HParamBackend (legacy, staging)
│   │   ├── ComBackend (generic COM)
│   │   └── AttrBackend (pure Python)
│   ├── Property Descriptors:
│   │   ├── IntProperty, BoolProperty, StringProperty
│   │   ├── ColorProperty, UnitProperty
│   │   ├── MappedProperty, TypedProperty, ListProperty
│   │   └── Auto-registration via ParameterSetMeta
│   └── 130+ ParameterSet Subclasses:
│       ├── Text/Char: CharShape, ParaShape, BulletShape, etc.
│       ├── Tables: Table, Cell, TableCreation, etc.
│       ├── Drawing: ShapeObject, DrawLineAttr, DrawImageAttr, etc.
│       ├── Document: DocumentInfo, PageDef, SecDef, etc.
│       └── [All mixed in single 3,357-line file]
│
├── Custom Accessors (Pythonic convenience layer)
│   ├── MoveAccessor: Navigation (move.top_of_file(), move.bottom(), etc.)
│   ├── CellAccessor: Table cell operations
│   ├── TableAccessor: Table operations
│   └── PageAccessor: Page operations
│
└── Dataclasses (Alternative representation)
    ├── Character, CharShape (dataclass)
    ├── Paragraph, ParaShape (dataclass)
    └── PageShape (dataclass)
```

**Key Characteristics:**
1. **Flat Entry Point**: Single `App` object, no collections exposed
2. **Action Properties**: 900+ actions as dynamic properties on `_Actions`
3. **Backend Polymorphism**: 4 backend types handle different parameter storage
4. **Pythonic Wrappers**: Custom accessors hide COM complexity
5. **Monolithic ParameterSets**: All 130+ classes in one file

---

### Comparison Matrix

| Aspect | Official HWP Model | Current hwpapi | Alignment |
|--------|-------------------|----------------|-----------|
| **Entry Point** | `IHwpObject` COM object | `App` wrapper around `Engine` | ✅ Aligned (wrapped) |
| **Document Access** | `IXHwpDocuments` collection | `App.api` direct access | ❌ Collection pattern not exposed |
| **Window Management** | `IXHwpWindows` collection | `App.set_visible()` only | ⚠️ Partial (no multi-window support) |
| **Form Controls** | `IXHwpForms` collection | Not exposed | ❌ Missing |
| **Action Execution** | `HAction.Execute(id, pset)` | `app.actions.ActionName(pset)` | ✅ Aligned (pythonic wrapper) |
| **Parameter Sets** | `HParameterSet` COM object | `ParameterSet` Python classes | ✅ Well abstracted |
| **Parameter Typing** | COM types | Python property descriptors | ✅ Excellent (better than COM) |
| **Nested Parameters** | `CreateItemSet` method | `NestedProperty` auto-creates | ✅ Enhanced (auto-creating) |
| **Arrays (HArray)** | COM array methods | `ArrayProperty` + `HArrayWrapper` | ✅ Enhanced (Pythonic list) |
| **Navigation** | Object hierarchy | Custom accessors | ⚠️ Different paradigm |
| **Organization** | Domain-based modules | Single monolithic file | ❌ Poor organization |

---

### Identified Gaps and Misalignments

#### 1. Missing Collection Objects ❌

**Issue:** hwpapi doesn't expose collection objects like `IXHwpDocuments`, `IXHwpWindows`, `IXHwpForms`

**Impact:**
- Cannot enumerate open documents
- Cannot manage multiple windows
- No access to form controls
- Limits multi-document workflows

**Example (What's Missing):**
```python
# This is possible in official HWP but not in hwpapi:
documents = hwp.Documents  # Collection of all open documents
doc = documents.Item(0)    # Get first document
doc.Save()                 # Save specific document
```

**Current hwpapi:**
```python
# Only single document access:
app = App()  # Always refers to "current" document
app.save()   # Saves "current" document only
```

#### 2. Monolithic ParameterSets Module ❌

**Issue:** All 130+ ParameterSet classes crammed into single 3,357-line file

**Impact:**
- Hard to navigate and maintain
- No logical grouping by domain
- Merge conflicts in team development
- Slow IDE performance

**Breakdown:**
```
parametersets.py (3,357 lines):
├── Mappings (147 lines): All DIRECTION_MAP, ALIGNMENT_MAP, etc.
├── Backend System (350 lines): Protocols, backend classes
├── Property Descriptors (250 lines): IntProperty, BoolProperty, etc.
├── ParameterSet Base (150 lines): Base class, metaclass
└── 130+ ParameterSet Classes (2,460 lines):
    ├── Text/Character (15 classes): CharShape, ParaShape, BulletShape, etc.
    ├── Tables (12 classes): Table, Cell, TableCreation, etc.
    ├── Drawing/Shapes (25 classes): ShapeObject, DrawLineAttr, etc.
    ├── Document (18 classes): DocumentInfo, PageDef, SecDef, etc.
    ├── Find/Replace (5 classes): FindReplace, DocFindInfo, etc.
    ├── Forms (8 classes): AutoFill, AutoNum, FieldCtrl, etc.
    ├── Formatting (12 classes): BorderFill, Caption, DropCap, etc.
    ├── Actions (15 classes): FileOpen, FileSaveAs, Print, etc.
    └── Misc (20 classes): Everything else
```

#### 3. Navigation Paradigm Mismatch ⚠️

**Issue:** Official model uses object hierarchy, hwpapi uses position-based accessors

**Official Model (Object-Based):**
```python
# Hypothetical object-based navigation:
document = app.ActiveDocument
section = document.Sections[0]
paragraph = section.Paragraphs[5]
text = paragraph.Text
```

**Current hwpapi (Position-Based):**
```python
# Position-based navigation:
app.move.current_list(para=5, pos=0)
app.actions.CharShape(...)
```

**Analysis:** Current approach is more pragmatic for HWP's position-based model. No change needed.

#### 4. Form Controls Not Exposed ❌

**Issue:** No access to IXHwpForms, form button controls, etc.

**Impact:**
- Cannot automate form-based documents
- Cannot create interactive PDFs with forms
- Missing feature parity with official API

---

### Proposed Restructuring Plan

#### Phase 1: Reorganize ParameterSets Module (High Priority)

**Goal:** Split monolithic `parametersets.py` into domain-based submodules

**New Structure:**
```
hwpapi/
├── parametersets/
│   ├── __init__.py              # Re-export all classes for compatibility
│   ├── base.py                  # ParameterSet base class, metaclass
│   ├── backends.py              # Backend protocol, implementations
│   ├── properties.py            # Property descriptors
│   ├── mappings.py              # All DIRECTION_MAP, ALIGNMENT_MAP, etc.
│   ├── text/
│   │   ├── __init__.py
│   │   ├── character.py         # CharShape, BulletShape
│   │   ├── paragraph.py         # ParaShape, TabDef, ListProperties
│   │   └── numbering.py         # NumberingShape, AutoNum
│   ├── table/
│   │   ├── __init__.py
│   │   ├── table.py             # Table, TableCreation
│   │   └── cell.py              # Cell, CellBorderFill
│   ├── drawing/
│   │   ├── __init__.py
│   │   ├── shape.py             # ShapeObject, DrawLayout
│   │   ├── line.py              # DrawLineAttr
│   │   ├── image.py             # DrawImageAttr, DrawImageScissoring
│   │   └── effects.py           # DrawShadow, DrawRotate, DrawTextart
│   ├── document/
│   │   ├── __init__.py
│   │   ├── info.py              # DocumentInfo, SummaryInfo, VersionInfo
│   │   ├── page.py              # PageDef, PageBorderFill, MasterPage
│   │   └── section.py           # SecDef, ColDef
│   ├── formatting/
│   │   ├── __init__.py
│   │   ├── border.py            # BorderFill, BorderFillExt
│   │   ├── caption.py           # Caption, FootnoteShape
│   │   └── style.py             # Style, StyleTemplate
│   ├── actions/
│   │   ├── __init__.py
│   │   ├── file.py              # FileOpen, FileSaveAs, FileConvert
│   │   ├── edit.py              # FindReplace, ConvertCase, etc.
│   │   └── print.py             # Print, PrintToImage, PrintWatermark
│   └── forms/
│       ├── __init__.py
│       └── fields.py            # AutoFill, FieldCtrl, HyperLink
```

**Benefits:**
- **Maintainability**: Find CharShape in `text/character.py`, not line 850 of monolith
- **Team Development**: Fewer merge conflicts with separated files
- **IDE Performance**: Faster autocomplete, syntax highlighting
- **Logical Grouping**: Related classes together by domain
- **Backward Compatible**: Re-export from `__init__.py` preserves existing imports

**Migration:**
```python
# Old import (still works):
from hwpapi.parametersets import CharShape, ParaShape, Table

# New import (also works):
from hwpapi.parametersets.text.character import CharShape
from hwpapi.parametersets.text.paragraph import ParaShape
from hwpapi.parametersets.table.table import Table
```

**Estimated Impact:**
- Lines reduced: 0 (reorganization, not deletion)
- Files created: ~20 new files
- Maintainability: 🔼 Significantly improved
- IDE performance: 🔼 Improved

#### Phase 2: Expose Collection Objects (Medium Priority)

**Goal:** Expose `Documents`, `Windows` collections to match official API

**New API:**
```python
# Add to App class:
class App:
    @property
    def documents(self):
        """Access to IXHwpDocuments collection."""
        return DocumentsCollection(self.api)

    @property
    def windows(self):
        """Access to IXHwpWindows collection."""
        return WindowsCollection(self.api)

    @property
    def active_document(self):
        """Currently active document."""
        return Document(self.api.ActiveDocument)

# New collection classes:
class DocumentsCollection:
    def __init__(self, hwp_object):
        self._hwp = hwp_object

    def __len__(self):
        return self._hwp.Documents.Count

    def __getitem__(self, index):
        return Document(self._hwp.Documents.Item(index))

    def add(self):
        """Create new document."""
        return Document(self._hwp.Documents.Add())

class Document:
    def __init__(self, doc_com_object):
        self._doc = doc_com_object

    @property
    def full_name(self):
        return self._doc.FullName

    def save(self):
        return self._doc.Save()

    def close(self):
        return self._doc.Close()
```

**Usage:**
```python
app = App()

# Access collections:
print(f"Open documents: {len(app.documents)}")
doc1 = app.documents[0]
doc2 = app.documents[1]

# Multi-document workflows:
for doc in app.documents:
    print(doc.full_name)
    doc.save()

# Create new document:
new_doc = app.documents.add()
```

**Benefits:**
- Feature parity with official API
- Multi-document support
- More explicit than implicit "current document"
- Better for automation scripts

#### Phase 3: Add Form Controls Support (Low Priority)

**Goal:** Expose form controls for interactive documents

**New Classes:**
```python
# Add to App:
class App:
    @property
    def forms(self):
        """Access to form controls."""
        return FormsCollection(self.api)

class FormsCollection:
    def __init__(self, hwp_object):
        self._hwp = hwp_object

    @property
    def push_buttons(self):
        return PushButtonsCollection(self._hwp.Forms.PushButtons)

    @property
    def check_buttons(self):
        return CheckButtonsCollection(self._hwp.Forms.CheckButtons)

    # etc...
```

**Benefits:**
- Support form-based documents
- Enable interactive workflows
- Complete API coverage

---

### Restructuring Priorities (Updated)

| Priority | Task | Lines Saved | Complexity Reduction | User Impact |
|----------|------|-------------|---------------------|-------------|
| **1** | Split parametersets.py by domain | 0 (reorg) | 🔼🔼🔼 High | Low (internal) |
| **2** | Unify backend modes | ~200 | 🔼🔼 Medium | Low (internal) |
| **3** | Expose Documents/Windows collections | +150 | 🔽 Slight increase | 🔼🔼 High (feature) |
| **4** | Consolidate property types | ~200 | 🔼 Medium | Low (internal) |
| **5** | Add Form controls support | +200 | 🔽 Slight increase | 🔼 Medium (feature) |
| **6** | Remove forward declarations | ~25 | 🔼 Small | None |

**Recommendation:** Start with Priority 1 (split parametersets.py) as it has:
- Highest maintainability impact
- Zero breaking changes
- Easiest to implement (move code, update imports)

---

### Implementation Strategy

#### Step 1: Prepare New Structure (parametersets/ package)

1. Create directory structure:
   ```bash
   mkdir -p hwpapi/parametersets/{text,table,drawing,document,formatting,actions,forms}
   ```

2. Create `__init__.py` files with re-exports:
   ```python
   # hwpapi/parametersets/__init__.py
   from .base import ParameterSet, ParameterSetMeta
   from .backends import *
   from .properties import *
   from .text.character import CharShape, BulletShape
   from .text.paragraph import ParaShape, TabDef
   # ... (re-export all classes to preserve imports)

   __all__ = ['ParameterSet', 'CharShape', 'ParaShape', ...]  # Full list
   ```

3. Move classes to domain files:
   - Extract CharShape, BulletShape → `text/character.py`
   - Extract ParaShape, TabDef → `text/paragraph.py`
   - Continue for all 130+ classes

4. Test imports:
   ```python
   # Ensure backward compatibility:
   from hwpapi.parametersets import CharShape  # Should still work
   ```

#### Step 2: Document Migration

1. Update CLAUDE.md with new file mappings
2. Update README with new structure
3. Create migration guide for contributors

#### Step 3: Gradual Rollout

1. **Phase 1a**: Move base classes, backends, properties (low risk)
2. **Phase 1b**: Move text-related classes (medium risk)
3. **Phase 1c**: Move remaining classes (high risk, test thoroughly)
4. **Phase 1d**: Update documentation, examples

---

### Benefits Summary

**Immediate Benefits (Phase 1 - Reorganization):**
- ✅ **Navigability**: Find classes 10x faster
- ✅ **Maintainability**: Logical grouping by domain
- ✅ **Team Collaboration**: Fewer merge conflicts
- ✅ **IDE Performance**: Faster autocomplete
- ✅ **Code Reviews**: Easier to review focused changes
- ✅ **Zero Breaking Changes**: Backward compatible via re-exports

**Long-term Benefits (Phase 2-3 - New Features):**
- ✅ **API Completeness**: Match official HWP API surface
- ✅ **Multi-document Support**: Automate across multiple files
- ✅ **Form Support**: Interactive document automation
- ✅ **Better Alignment**: Official docs map directly to code structure

**Non-Goals:**
- ❌ Don't change the property descriptor system (it's excellent)
- ❌ Don't change the backend abstraction (it works well)
- ❌ Don't change the Actions pattern (pythonic and convenient)
- ❌ Don't add complexity for theoretical future needs

---

### Migration Checklist

When implementing the restructuring:

**Preparation:**
- [ ] Read official HwpAutomation_2504.pdf documentation
- [ ] Understand current parametersets.py organization
- [ ] Create domain-based file structure

**Phase 1 Execution:**
- [ ] Create `hwpapi/parametersets/` package structure
- [ ] Move base classes to `base.py`
- [ ] Move backend classes to `backends.py`
- [ ] Move property descriptors to `properties.py`
- [ ] Move mappings to `mappings.py`
- [ ] Move ParameterSet subclasses to domain files
- [ ] Create `__init__.py` with full re-exports
- [ ] Test all imports: `python -c "from hwpapi.parametersets import CharShape, ParaShape, Table"`
- [ ] Run full test suite: `python -m pytest tests/ -v`
- [ ] Update CLAUDE.md file mapping table
- [ ] Commit changes

**Phase 2 Execution (Optional):**
- [ ] Design DocumentsCollection, WindowsCollection classes
- [ ] Add `app.documents`, `app.windows` properties
- [ ] Write tests for multi-document workflows
- [ ] Update documentation with examples
- [ ] Commit changes

**Phase 3 Execution (Optional):**
- [ ] Design FormsCollection classes
- [ ] Add `app.forms` property
- [ ] Write tests for form controls
- [ ] Update documentation
- [ ] Commit changes

---

## 🔄 Version History

### Recent Changes

**2025-12-09 - Logging System Improvements**
- **Default Log Level Changed**: Changed from `INFO` to `WARNING`
  - Production-friendly: Normal users only see warnings, errors, and critical messages
  - Reduces log clutter in release builds
  - Developers can still enable detailed logging via `HWPAPI_LOG_LEVEL=DEBUG` or `INFO`
- **Enhanced Documentation**: Added comprehensive logging configuration examples
- **Environment Variable**: `HWPAPI_LOG_LEVEL` now defaults to `WARNING` instead of `INFO`
- Files: `hwpapi/logging.py`
- Impact: Cleaner output for end users, opt-in verbose logging for developers

**2025-12-09 - Complete Display Enhancement Suite**
- **Critical Bug Fix**: Removed duplicate ParameterSet class definition in cell 26
  - Entire class (27,304 characters) was duplicated
  - Second `_format_int_value` overrode first with old logic
  - Caused `FontSize` to show as `1200` instead of `12.0pt`

- **Enhancement 1: Human-Readable Value Formatting**
  - Colors: `0x0000FF` → `#FF0000` (BBGGRR to hex)
  - Font sizes: `1200` → `12.0pt` (HWPUNIT to pt)
  - Dimensions: `59430` → `210.0mm` (HWPUNIT to mm)
  - Booleans: Display as `True`/`False`

- **Enhancement 2: Enum Display for MappedProperty**
  - Before: `Direction="down"`
  - After: `Direction=0 (down)` (shows both value and name)
  - Works with Korean: `Type=0 (일반책갈피)`
  - Automatically detects MappedProperty and formats accordingly

- **Enhancement 3: Property Description Comments**
  - Before: `FontSize=12.0pt`
  - After: `FontSize=12.0pt  # Font size in HWPUNIT (100 = 1pt)`
  - Shows inline documentation for every property
  - Supports multilingual descriptions (Korean, English)

- **Complete Example**:
  ```python
  CharFormat(
    FontSize=12.0pt  # Font size in HWPUNIT (100 = 1pt)
    TextColor="#ff0000"  # Text color in BBGGRR format
    Direction=0 (down)  # Search direction (down=forward, up=backward)
  )
  ```

- **Detection Tools**: Added scripts to detect duplicate method definitions
- **Documentation**: Updated CLAUDE.md with Issue 5, debugging tips, examples
- Result: Self-documenting, human-readable parameter display
- Files: `hwpapi/parametersets.py`
- Examples: `nested_property_demo.py`, `mapped_property_display_demo.py`, `property_description_display_demo.py`

**2025-01-08 - Auto-Creating Properties Design**
- Designed `NestedProperty` for auto-creating nested ParameterSets
- Designed `ArrayProperty` for Pythonic HArray interface
- **Enhanced `UnitProperty`** for smart unit conversion (mm, cm, in, pt ↔ HWPUNIT)
- Created comprehensive design documents:
  - AUTO_PROPERTY_DESIGN.md - NestedProperty & ArrayProperty specification
  - UNIT_PROPERTY_ENHANCEMENT.md - Smart unit conversion specification
- Updated CLAUDE.md with complete documentation:
  - "Auto-Creating Properties" section with full examples
  - UnitProperty section with unit conversion examples
  - Migration guides for all property types
  - Property type decision tree with unit selection guide
- Key improvements:
  - Tab completion for nested properties
  - No manual create_itemset() calls
  - Pythonic array interface (append, insert, pop, etc.)
  - **Intuitive units: "210mm", "21cm", "8.27in" instead of HWPUNIT**
- Result: Intuitive API that feels natural for Python developers
- Status: Design complete, ready for implementation

**2025-01-08 - Architecture Analysis & Restructuring Plan**
- Analyzed official HWP Automation Object Model (HwpAutomation_2504.pdf)
- Compared official structure with current hwpapi implementation
- Identified 4 major gaps: Collection objects, monolithic parametersets.py, form controls, organization
- Designed 3-phase restructuring plan:
  - Phase 1: Split parametersets.py into domain-based modules (highest priority)
  - Phase 2: Expose Documents/Windows collections (medium priority)
  - Phase 3: Add Form controls support (low priority)
- Added comprehensive "HWP Object Model: Official vs Current Architecture" section to CLAUDE.md
- Result: Clear roadmap for better alignment with official API and improved maintainability

**2024 - Auto-Generated attributes_names**
- Removed manual `self.attributes_names = [...]` from 9+ classes
- Added `@property attributes_names` to ParameterSet
- Fixed `_del_value` to handle None backend
- Updated tests to use property descriptors
- Result: -500 lines, eliminated sync issues

**2024 - Added _is_com Function**
- Fixed NameError in `_looks_like_pset`
- Added proper COM object detection
- Now properly distinguishes COM vs non-COM objects

**Earlier - Pset Migration**
- See PSET_MIGRATION_SUMMARY.md
- Migrated from HSet-based to pset-based approach
- Maintained backward compatibility

---

## ✅ Checklist for New Contributors

Before making changes:
- [ ] Read this entire document
- [ ] Set up development environment
- [ ] Can run tests

Before committing:
- [ ] Edited .py files directly
- [ ] Ran tests (`python -m pytest tests/ -v`)
- [ ] Verified imports work
- [ ] Added/updated tests if needed
- [ ] Commit message describes what and why

---

## 📞 Quick Reference Commands

```bash
# Run all tests
python -m pytest tests/ -v

# Test specific module
python -m pytest tests/test_hparam.py -v

# Verify imports
python -c "import hwpapi; print('OK')"

# Install in dev mode
pip install -e .
```

---

## 🎯 Remember

1. **Standard Python project**: .py files are the source of truth
2. **Backend abstraction works**: Trust `make_backend()` factory
3. **Properties are auto-registered**: No manual `attributes_names` needed
4. **Always check for None**: Backend might not be initialized
5. **Test your changes**: Don't break existing functionality
6. **Keep it simple**: Prefer simplification over clever abstractions
7. **Document decisions**: Update this file when you learn something new

---

*This document is a living guide. Update it as you learn more about the codebase.*

**Last Updated:** 2025-12-09 (After fixing duplicate class and display formatting)
**Next Review:** After Phase 1 restructuring (split parametersets.py by domain)
