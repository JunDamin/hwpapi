# Enhanced UnitProperty Design

**Version:** 1.0
**Date:** 2025-01-08
**Purpose:** Make unit-based properties intuitive with automatic mm/cm/in/pt conversion

---

## Problem

**Current UnitProperty limitations:**
```python
# Current - only accepts numbers, unit is fixed
class PageDef(ParameterSet):
    width = UnitProperty("Width", "mm", "Page width")

# Usage - not intuitive!
page.width = 210  # Is this mm? User has to remember!
# Cannot use: page.width = "210mm" ❌
# Cannot use: page.width = "21cm" ❌
```

**HWPUNIT is not intuitive:**
- 1 mm = 283 HWPUNIT
- 1 pt = 100 HWPUNIT
- 1 cm = 2830 HWPUNIT
- 1 inch = 7200 HWPUNIT (approximately)

Users want to work with familiar units: **millimeters, centimeters, inches, points**.

---

## Solution: Smart Unit Conversion

### Enhanced UnitProperty Features

1. **String with units:** `"210mm"`, `"21cm"`, `"8.27in"`, `"100pt"`
2. **Numbers with default unit:** `210` (assumes mm by default)
3. **Auto-conversion to HWPUNIT:** Stores internally as HWPUNIT
4. **Returns in original or specified unit:** Flexible output
5. **Validation:** Ensures values are positive, reasonable ranges

---

## Implementation

### 1. Enhanced Conversion Functions

```python
# In functions.py - extend existing conversions

HWPUNIT_PER_MM = 283
HWPUNIT_PER_PT = 100
HWPUNIT_PER_CM = 2830
HWPUNIT_PER_INCH = 7200  # 1 inch = 25.4mm * 283

def parse_unit_string(value: Union[str, int, float], default_unit: str = "mm") -> Tuple[float, str]:
    """
    Parse a value with optional unit suffix.

    Args:
        value: Number or string like "210mm", "21cm", "8.27in", "100pt"
        default_unit: Unit to assume if value is numeric (default: "mm")

    Returns:
        (numeric_value, unit)

    Examples:
        >>> parse_unit_string("210mm")
        (210.0, "mm")
        >>> parse_unit_string("21cm")
        (21.0, "cm")
        >>> parse_unit_string(210)
        (210.0, "mm")  # default unit
        >>> parse_unit_string("8.27in")
        (8.27, "in")
    """
    if isinstance(value, (int, float)):
        return (float(value), default_unit)

    if isinstance(value, str):
        # Try to parse unit suffix
        match = re.match(r'^([0-9.]+)\s*(mm|cm|in|inch|pt|point)?$', value.lower().strip())
        if match:
            number = float(match.group(1))
            unit = match.group(2) or default_unit
            # Normalize unit names
            if unit in ['inch']:
                unit = 'in'
            elif unit in ['point']:
                unit = 'pt'
            return (number, unit)

    raise ValueError(f"Invalid unit value: {value}")


def to_hwpunit(value: Union[str, int, float], default_unit: str = "mm") -> int:
    """
    Convert any unit to HWPUNIT.

    Args:
        value: Number or string with unit ("210mm", "21cm", 210, etc.)
        default_unit: Unit to assume for bare numbers

    Returns:
        HWPUNIT value (integer)

    Examples:
        >>> to_hwpunit("210mm")
        59430  # 210 * 283
        >>> to_hwpunit("21cm")
        59430  # 21 * 2830
        >>> to_hwpunit(210)  # assumes mm
        59430
        >>> to_hwpunit("8.27in")
        59544  # 8.27 * 7200
    """
    number, unit = parse_unit_string(value, default_unit)

    conversions = {
        'mm': HWPUNIT_PER_MM,
        'cm': HWPUNIT_PER_CM,
        'in': HWPUNIT_PER_INCH,
        'pt': HWPUNIT_PER_PT,
    }

    if unit not in conversions:
        raise ValueError(f"Unknown unit: {unit}. Supported: mm, cm, in, pt")

    return int(round(number * conversions[unit]))


def from_hwpunit(hwpunit_value: int, target_unit: str = "mm") -> float:
    """
    Convert HWPUNIT to specified unit.

    Args:
        hwpunit_value: Value in HWPUNIT
        target_unit: Desired output unit ("mm", "cm", "in", "pt")

    Returns:
        Value in target unit

    Examples:
        >>> from_hwpunit(59430, "mm")
        210.0
        >>> from_hwpunit(59430, "cm")
        21.0
        >>> from_hwpunit(59430, "in")
        8.27
    """
    conversions = {
        'mm': HWPUNIT_PER_MM,
        'cm': HWPUNIT_PER_CM,
        'in': HWPUNIT_PER_INCH,
        'pt': HWPUNIT_PER_PT,
    }

    if target_unit not in conversions:
        raise ValueError(f"Unknown unit: {target_unit}. Supported: mm, cm, in, pt")

    return round(hwpunit_value / conversions[target_unit], 2)
```

### 2. Enhanced UnitProperty Class

```python
class UnitProperty(PropertyDescriptor):
    """
    Property descriptor for unit-based values with automatic conversion.

    Accepts values with units (e.g., "210mm", "21cm", "8.27in") or bare numbers
    (which use the default unit). Automatically converts to/from HWPUNIT.

    Attributes:
        key (str): Parameter key in HWP
        doc (str): Documentation string
        default_unit (str): Unit to assume for bare numbers (default: "mm")
        output_unit (str): Unit to return when getting value (default: same as default_unit)
        min_value (float): Minimum allowed value in output units (optional)
        max_value (float): Maximum allowed value in output units (optional)

    Example:
        >>> class PageDef(ParameterSet):
        ...     width = UnitProperty("Width", "Page width", default_unit="mm")
        ...     height = UnitProperty("Height", "Page height", default_unit="mm")

        >>> page = PageDef(action.CreateSet())
        >>> page.width = "210mm"  # String with unit
        >>> page.width = "21cm"   # Different unit, auto-converts
        >>> page.width = 210      # Bare number, assumes mm
        >>> print(page.width)     # Returns in mm: 210.0
    """

    def __init__(self, key: str, doc: str,
                 default_unit: str = "mm",
                 output_unit: Optional[str] = None,
                 min_value: Optional[float] = None,
                 max_value: Optional[float] = None):
        super().__init__(key, doc)
        self.default_unit = default_unit
        self.output_unit = output_unit or default_unit
        self.min_value = min_value
        self.max_value = max_value

    def __get__(self, instance, owner) -> Optional[float]:
        """
        Get value in output unit.

        Returns:
            Value in output_unit, or None if not set
        """
        if instance is None:
            return self

        hwpunit_value = self._get_value(instance)
        if hwpunit_value is None:
            return None

        # Convert from HWPUNIT to output unit
        return from_hwpunit(hwpunit_value, self.output_unit)

    def __set__(self, instance, value: Union[str, int, float, None]):
        """
        Set value from any unit format.

        Args:
            value: Can be:
                - String with unit: "210mm", "21cm", "8.27in", "100pt"
                - Number: assumes default_unit
                - None: clears the value

        Raises:
            TypeError: If value is not str/int/float/None
            ValueError: If value is out of range or invalid unit
        """
        if value is None:
            return self._del_value(instance)

        # Convert to HWPUNIT
        try:
            hwpunit_value = to_hwpunit(value, self.default_unit)
        except ValueError as e:
            raise ValueError(f"Invalid value for '{self.key}': {e}")

        # Validate range (convert back to output unit for comparison)
        if self.min_value is not None or self.max_value is not None:
            check_value = from_hwpunit(hwpunit_value, self.output_unit)

            if self.min_value is not None and check_value < self.min_value:
                raise ValueError(
                    f"Value {check_value}{self.output_unit} for '{self.key}' is below "
                    f"minimum {self.min_value}{self.output_unit}"
                )

            if self.max_value is not None and check_value > self.max_value:
                raise ValueError(
                    f"Value {check_value}{self.output_unit} for '{self.key}' is above "
                    f"maximum {self.max_value}{self.output_unit}"
                )

        # Store as HWPUNIT
        return self._set_value(instance, hwpunit_value)
```

---

## Usage Examples

### Example 1: Page Dimensions

```python
class PageDef(ParameterSet):
    """Page layout definition."""

    # Width/height in millimeters (most common for paper)
    width = UnitProperty("Width", "Page width", default_unit="mm", min_value=10, max_value=500)
    height = UnitProperty("Height", "Page height", default_unit="mm", min_value=10, max_value=500)

    # Margins in millimeters
    left_margin = UnitProperty("LeftMargin", "Left margin", default_unit="mm")
    right_margin = UnitProperty("RightMargin", "Right margin", default_unit="mm")
    top_margin = UnitProperty("TopMargin", "Top margin", default_unit="mm")
    bottom_margin = UnitProperty("BottomMargin", "Bottom margin", default_unit="mm")

# Usage - all of these work!
page = PageDef(action.CreateSet())

# A4 paper size
page.width = "210mm"  # String with unit
page.height = "297mm"

# Or use centimeters
page.width = "21cm"   # Auto-converts to HWPUNIT
page.height = "29.7cm"

# Or bare numbers (assumes mm)
page.width = 210      # Assumes mm
page.height = 297

# Or use inches (Letter size)
page.width = "8.5in"
page.height = "11in"

# Set margins
page.left_margin = 25    # 25mm (bare number)
page.right_margin = "2.5cm"  # 25mm (converted from cm)
page.top_margin = "1in"      # ~25.4mm (converted from inches)
page.bottom_margin = 25

# Get values (returns in mm by default)
print(f"Page size: {page.width}mm × {page.height}mm")
# Output: Page size: 210.0mm × 297.0mm
```

### Example 2: Table Dimensions

```python
class Table(ParameterSet):
    """Table creation parameters."""

    rows = IntProperty("Rows", "Number of rows")
    cols = IntProperty("Cols", "Number of columns")

    # Table width/height in millimeters
    table_width = UnitProperty("Width", "Table width", default_unit="mm")
    table_height = UnitProperty("Height", "Table height", default_unit="mm")

    # Column widths array (each in mm)
    column_widths = ArrayProperty("ColWidths", int, "Column widths in HWPUNIT")

# Usage
table = Table(action.CreateSet())
table.rows = 3
table.cols = 4

# Set table size - all valid!
table.table_width = "150mm"  # String with unit
table.table_height = 80      # Bare number, assumes mm

# Or use centimeters
table.table_width = "15cm"   # Same as 150mm
```

### Example 3: Font Size (Points)

```python
class CharShape(ParameterSet):
    """Character formatting."""

    # Font size in points (standard for fonts)
    fontsize = UnitProperty("Height", "Font size",
                           default_unit="pt", output_unit="pt",
                           min_value=1, max_value=200)

# Usage
char = CharShape(action.CreateSet())

# Set font size - all work!
char.fontsize = 12       # 12pt (bare number)
char.fontsize = "12pt"   # 12pt (with unit)
char.fontsize = "16pt"   # 16pt

# Can even use mm if needed
char.fontsize = "4.23mm" # Converts to points internally

# Get value (returns in points)
print(f"Font size: {char.fontsize}pt")
# Output: Font size: 12.0pt
```

### Example 4: Mixed Units

```python
class BorderFill(ParameterSet):
    """Border and fill settings."""

    # Border widths can use different units
    border_left = UnitProperty("BorderLeft", "Left border width", default_unit="mm")
    border_right = UnitProperty("BorderRight", "Right border width", default_unit="mm")
    border_top = UnitProperty("BorderTop", "Top border width", default_unit="mm")
    border_bottom = UnitProperty("BorderBottom", "Bottom border width", default_unit="mm")

# Usage - mix units freely!
border = BorderFill(action.CreateSet())
border.border_left = "2mm"     # 2mm
border.border_right = "0.2cm"  # 2mm (converted)
border.border_top = 2          # 2mm (bare number)
border.border_bottom = "5.67pt" # ~2mm (converted from points)
```

---

## Conversion Reference

### Standard Conversions

| Unit | HWPUNIT Multiplier | Example |
|------|-------------------|---------|
| **mm** (millimeter) | 283 | 1mm = 283 HWPUNIT |
| **cm** (centimeter) | 2830 | 1cm = 2830 HWPUNIT |
| **in** (inch) | 7200 | 1in = 7200 HWPUNIT |
| **pt** (point) | 100 | 1pt = 100 HWPUNIT |

### Common Paper Sizes

| Paper | Width × Height (mm) | Width × Height (in) |
|-------|---------------------|---------------------|
| **A4** | 210mm × 297mm | 8.27in × 11.69in |
| **A3** | 297mm × 420mm | 11.69in × 16.54in |
| **Letter** | 215.9mm × 279.4mm | 8.5in × 11in |
| **Legal** | 215.9mm × 355.6mm | 8.5in × 14in |

### Common Font Sizes

| Size Name | Points | Millimeters |
|-----------|--------|-------------|
| Small | 9pt | 3.18mm |
| Normal | 12pt | 4.23mm |
| Large | 14pt | 4.94mm |
| Title | 18pt | 6.35mm |
| Heading | 24pt | 8.47mm |

---

## Benefits

### User Experience

**Before (confusing):**
```python
page.width = 59430  # What unit is this?? (HWPUNIT)
```

**After (intuitive):**
```python
page.width = "210mm"  # Clear! A4 width
page.width = "21cm"   # Also clear!
page.width = 210      # Assumes mm (documented)
```

### Flexibility

- ✅ **Multiple units** - Use mm, cm, in, or pt
- ✅ **String or number** - Both `"210mm"` and `210` work
- ✅ **Auto-conversion** - Internally uses HWPUNIT
- ✅ **Validation** - Optional min/max in user units
- ✅ **Type-safe** - Clear error messages

### Backward Compatibility

```python
# Old UnitProperty (numeric only) still works
class OldPage(ParameterSet):
    width = UnitProperty("Width", "mm", "Width")  # Old signature

page.width = 210  # Still works!

# New UnitProperty (string support) is better
class NewPage(ParameterSet):
    width = UnitProperty("Width", "Width", default_unit="mm")

page.width = "210mm"  # Now also works!
page.width = "21cm"   # And this!
```

---

## Default Unit Recommendations

### By Property Type

| Property | Recommended Default Unit | Reason |
|----------|-------------------------|---------|
| **Page width/height** | `mm` | International standard for paper |
| **Margins** | `mm` | Easier to visualize than cm for small values |
| **Table width/height** | `mm` | Consistent with page dimensions |
| **Font size** | `pt` | Standard unit for typography |
| **Line width** | `mm` | More precise than pt for borders |
| **Spacing** | `pt` or `mm` | pt for text, mm for layout |

---

## Implementation Checklist

**Phase 1: Update Conversion Functions**
- [ ] Enhance `parse_unit_string()` in functions.py
- [ ] Update `to_hwpunit()` to accept strings
- [ ] Update `from_hwpunit()` for output unit selection
- [ ] Add unit constants (HWPUNIT_PER_MM, etc.)
- [ ] Add validation and error messages

**Phase 2: Update UnitProperty Class**
- [ ] Add `default_unit` and `output_unit` parameters
- [ ] Update `__get__` to use output unit
- [ ] Update `__set__` to parse unit strings
- [ ] Add range validation
- [ ] Update docstring with examples

**Phase 3: Update Existing ParameterSet Classes**
- [ ] `PageDef` - Use mm for dimensions
- [ ] `CharShape` - Use pt for font size
- [ ] `Table` - Use mm for dimensions
- [ ] `BorderFill` - Use mm for widths
- [ ] Document default units in each class

**Phase 4: Testing**
- [ ] Test string parsing ("210mm", "21cm", etc.)
- [ ] Test unit conversions (mm ↔ cm ↔ in ↔ pt)
- [ ] Test validation (min/max values)
- [ ] Test backward compatibility
- [ ] Integration test with real HWP

---

## Migration Guide

### For Users

**Old code (still works):**
```python
page.width = 59430  # HWPUNIT value
```

**New code (recommended):**
```python
page.width = "210mm"  # Clear and intuitive!
```

### For Developers

**Update ParameterSet definitions:**
```python
# Before
class PageDef(ParameterSet):
    width = UnitProperty("Width", "mm", "Page width")

# After
class PageDef(ParameterSet):
    width = UnitProperty("Width", "Page width", default_unit="mm")
```

**Note:** Old signature `UnitProperty(key, unit, doc)` vs new `UnitProperty(key, doc, default_unit=...)` - may need shim for compatibility.

---

## Conclusion

Enhanced UnitProperty makes hwpapi more intuitive by:

1. **Accepting familiar units** - mm, cm, in, pt
2. **Smart string parsing** - "210mm" just works
3. **Automatic conversion** - No manual HWPUNIT calculation
4. **Flexible and clear** - Use whatever unit makes sense
5. **Validated** - Prevents invalid values

**Result:** Users can work with measurements they understand, while hwpapi handles HWPUNIT conversion transparently.

---

**Status:** Design complete, ready for implementation
**Next:** Update AUTO_PROPERTY_DESIGN.md to include UnitProperty
