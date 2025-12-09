# Duplicate Class Fix Summary

**Date:** 2025-12-09
**Issue:** Duplicate ParameterSet class definition causing method override bug
**Impact:** Critical - FontSize and other properties displayed incorrectly
**Status:** ✅ Fixed

---

## Problem

The entire `ParameterSet` class was duplicated in `nbs/02_api/02_parameters.ipynb` cell 26, causing the second definition to override the first.

### Symptoms

- `FontSize` property displayed as raw `1200` instead of formatted `12.0pt`
- Recent fixes to `_format_int_value` method appeared to not work
- Generated `hwpapi/parametersets.py` had 2 definitions of `_format_int_value`

### Root Cause

```
Cell 26 content:
├── ParameterSet class (correct, 29,239 chars)
│   ├── __init__
│   ├── ... all methods ...
│   ├── _format_int_value (✅ CORRECT LOGIC)
│   └── attributes_names
├── ParameterSet class DUPLICATE (27,304 chars)  # ← THE PROBLEM
│   ├── __init__
│   ├── ... all methods again ...
│   ├── _format_int_value (❌ OLD LOGIC)
│   └── attributes_names
```

In Python, when a class is defined twice, the second definition completely overrides the first.

### The Culprit: _format_int_value Logic

**First definition (correct):**
```python
# Font size properties - convert HWPUNIT to pt
if 'Size' in prop_name or prop_name.endswith('Size'):
    if value > 0 and value < 100000:
        pt_value = value / 100
        return f"{pt_value}pt"
```

**Second definition (wrong - overrode first):**
```python
# Font size properties - convert HWPUNIT to pt
if prop_type_name == 'UnitProperty' or 'Size' in prop_name and 'Font' not in prop_name:
    if value > 0 and value < 100000:
        pt_value = value / 100
        return f"{pt_value}pt"
```

The condition `'Font' not in prop_name` excluded `FontSize`, causing it to display as raw `1200`.

---

## Detection

### Step 1: Count Definitions in Generated File

```bash
$ python -c "with open('hwpapi/parametersets.py', encoding='utf-8') as f: print(f.read().count('def _format_int_value'))"
2  # ← Should be 1!
```

### Step 2: Locate Duplicate in Notebook

```bash
$ python -c "
import json
nb = json.load(open('nbs/02_api/02_parameters.ipynb', encoding='utf-8'))
cell = nb['cells'][26]
source = ''.join(cell['source'])
count = source.count('def _format_int_value')
print(f'Cell 26 has {count} definitions')
"
Cell 26 has 2 definitions
```

### Step 3: Find All Duplicate Methods

```bash
$ python << 'EOF'
import re
from collections import Counter

with open('hwpapi/parametersets.py', encoding='utf-8') as f:
    content = f.read()

methods = re.findall(r'\n    def (\w+)', content)
method_counts = Counter(methods)
duplicates = {name: count for name, count in method_counts.items() if count > 1}

if duplicates:
    print("DUPLICATES FOUND:")
    for name, count in duplicates.items():
        print(f"  {name}: {count} times")
EOF

DUPLICATES FOUND:
  __init__: 2 times
  bind: 2 times
  _format_int_value: 2 times
  ... (entire class duplicated)
```

---

## Fix

### Step 1: Identify Duplicate Boundaries

```python
import json

with open('nbs/02_api/02_parameters.ipynb', encoding='utf-8') as f:
    nb = json.load(f)

cell = nb['cells'][26]
source = ''.join(cell['source'])

# The duplicate started at position 29239 (second __init__)
# and went to end of cell (56543)
```

### Step 2: Remove Duplicate Block

```python
# Remove everything from position 29239 to end
new_source = source[:29239].rstrip() + '\n'

# Update notebook
nb['cells'][26]['source'] = [line + '\n' for line in new_source.split('\n')[:-1]] + [new_source.split('\n')[-1]]

# Save
with open('nbs/02_api/02_parameters.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, indent=1, ensure_ascii=False)
```

### Step 3: Export and Verify

```bash
$ nbdev_export
$ python -c "with open('hwpapi/parametersets.py', encoding='utf-8') as f: print(f.read().count('def _format_int_value'))"
1  # ✅ Fixed!
```

### Step 4: Test Display Formatting

```python
from hwpapi.parametersets import ParameterSet, IntProperty, ColorProperty

class TestDisplay(ParameterSet):
    FontSize = IntProperty("FontSize", "Font size")
    TextColor = ColorProperty("TextColor", "Text color")

pset = TestDisplay()
pset.FontSize = 1200
pset.TextColor = 0x0000FF

print(pset)
# Output:
# TestDisplay(
#   FontSize=12.0pt     # ✅ Now shows as pt!
#   TextColor="#ff0000" # ✅ Color formatted correctly!
#   [staged changes: 2]
# )
```

---

## Results

### Before

```
DisplayTest(
  FontSize=1200               # ❌ Raw HWPUNIT
  TextColor=255               # ❌ Raw BBGGRR integer
  Width=59430                 # ❌ Raw HWPUNIT
  [staged changes: 3]
)
```

### After

```
DisplayTest(
  FontSize=12.0pt             # ✅ Formatted in points
  TextColor="#ff0000"         # ✅ Formatted as hex color
  Width=210.0mm               # ✅ Formatted in millimeters
  [staged changes: 3]
)
```

---

## Human-Readable Display Feature

As part of this fix, we also completed the human-readable display formatting for ParameterSet:

### Implementation

Added to `ParameterSet` base class:
1. **`__repr__()`** - Human-readable representation showing all properties
2. **`_format_repr()`** - Recursive formatting with indentation
3. **`_format_int_value()`** - Smart value formatting based on property type

### Formatting Rules

| Property Type | Input | Output | Logic |
|--------------|-------|--------|-------|
| **Colors** | `0x0000FF` (BBGGRR) | `#FF0000` | Convert from BBGGRR to #RRGGBB hex |
| **Font Sizes** | `1200` (HWPUNIT) | `12.0pt` | `'Size' in prop_name` → divide by 100 |
| **Dimensions** | `59430` (HWPUNIT) | `210.0mm` | UnitProperty or Width/Height → convert via `from_hwpunit()` |
| **Booleans** | `True`/`False` | `True`/`False` | Direct string conversion |
| **Nested** | `CharShape(...)` | Recursive repr | Call `_format_repr()` recursively |
| **Arrays** | `HArrayWrapper` | `[1, 2, 3]` | Call `.to_list()` |

### Example Output

```python
FindReplace(
  FindString="Python"
  ReplaceString="Python 3.11"
  MatchCase=True
  WholeWordOnly=False
  Direction="down"
  FindCharShape=
    CharShape(
      Bold=True
      Italic=False
      Size=12.0pt
      TextColor="#ff0000"
    )
  FindParaShape=
    ParaShape(
      Align="center"
      LineSpacing=160
    )
  [staged changes: 7]
)
```

---

## Prevention

### 1. Always Check Generated File After Refactoring

```bash
# Count definitions before and after
grep -c "def method_name" hwpapi/parametersets.py
```

### 2. Use Duplicate Detection Script

Save this as `check_duplicates.py`:
```python
import re
from collections import Counter

with open('hwpapi/parametersets.py', encoding='utf-8') as f:
    content = f.read()

methods = re.findall(r'\n    def (\w+)', content)
method_counts = Counter(methods)
duplicates = {name: count for name, count in method_counts.items() if count > 1}

if duplicates:
    print("❌ DUPLICATES FOUND:")
    for name, count in duplicates.items():
        print(f"  {name}: {count} times")
    exit(1)
else:
    print("✅ No duplicates found.")
    exit(0)
```

### 3. Add to Verification Workflow

```bash
#!/bin/bash
echo "=== Running nbdev_export ==="
nbdev_export

echo -e "\n=== Checking for duplicates ==="
python check_duplicates.py || exit 1

echo -e "\n=== Running tests ==="
python -m pytest tests/ -v

echo -e "\n✅ All checks passed!"
```

### 4. Be Careful with Copy-Paste in Notebooks

- Always verify cell content after major refactoring
- Use `git diff` to review notebook changes before committing
- Check line count: sudden doubling indicates duplication

---

## Lessons Learned

1. **Duplicate definitions silently override** - Python doesn't warn about class redefinition
2. **Generated files are the source of truth** - If generated `.py` looks wrong, notebook has the problem
3. **Always verify after major edits** - Especially after copy-paste operations
4. **Count method definitions** - Simple `grep -c` catches duplicates immediately
5. **Human-readable display is valuable** - Raw HWPUNIT/BBGGRR values are hard to interpret
6. **Smart formatting helps debugging** - Seeing `12.0pt` instead of `1200` makes property values obvious

---

## Files Modified

- `nbs/02_api/02_parameters.ipynb` - Removed duplicate ParameterSet class (cell 26)
- `hwpapi/parametersets.py` - Regenerated with correct single definition
- `CLAUDE.md` - Updated with Issue 5, debugging tips, and lessons learned
- `DUPLICATE_FIX_SUMMARY.md` - This document

---

## Commit

```bash
git add nbs/02_api/02_parameters.ipynb hwpapi/parametersets.py CLAUDE.md DUPLICATE_FIX_SUMMARY.md
git commit -m "Fix duplicate ParameterSet class and FontSize formatting"
```

**Commit ID:** `8be15b3` (on branch `fix-parameterset-formatting`)

---

**Created:** 2025-12-09
**Author:** Claude Sonnet 4.5 via Claude Code
**Related:** CLAUDE.md, PSET_MIGRATION_SUMMARY.md, REFACTORING_SUMMARY.md
