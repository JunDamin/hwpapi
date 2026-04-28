# ParameterSet Display Enhancements

## Overview

The `ParameterSet.__repr__()` method has been enhanced with three powerful features that create self-documenting, human-readable output. These enhancements work together to make debugging and learning much easier.

## Enhancement 1: Human-Readable Value Formatting

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

## Enhancement 2: Enum Display for MappedProperty

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

## Enhancement 3: Property Description Comments

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

## All Three Enhancements Together

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

## Implementation Details

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
python examples/nested_property_demo.py
python examples/mapped_property_display_demo.py
python examples/property_description_display_demo.py
```

## Best Practices for Property Definitions

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
