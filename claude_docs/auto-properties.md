# Auto-Creating Properties: NestedProperty, UnitProperty & ArrayProperty

## Overview

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

## NestedProperty - Auto-Creating Nested ParameterSets

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

## UnitProperty - Smart Unit Conversion

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

## ArrayProperty - Auto-Creating HArray with List Interface

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

**Benefits:**
- ✅ **Pythonic** - Works exactly like Python lists
- ✅ **Tab completion** - IDE shows all list methods
- ✅ **Type validation** - Ensures all items match `item_type`
- ✅ **Length validation** - Optional min/max constraints
- ✅ **No COM knowledge needed** - Pure Python interface

## Complete Example: All Property Types Together

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

# Array assignments (auto-creates HArray)
table.column_widths = [2000, 3000, 2500, 2000]  # HWPUNIT values
table.row_heights = [1000, 1000, 1000]

# Array modifications
table.column_widths.append(1500)
table.column_widths[2] = 3500

# Nested object access (auto-creates BorderFill via CreateItemSet)
table.border_fill.border_type = "solid"
table.border_fill.fill_color = "#EEEEEE"

# Execute
table.run()
```

## Migration Guide

### From TypedProperty to NestedProperty

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

### From ListProperty to ArrayProperty

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

## Property Type Decision Tree

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

## Implementation Checklist

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

## Best Practices

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
