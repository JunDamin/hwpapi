# Version History & Lessons Learned

## Lessons Learned

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

5. **Human-readable display is valuable**
   - Raw HWPUNIT/BBGGRR values are not intuitive
   - Smart formatting (`_format_int_value`) makes debugging easier
   - `__repr__` showing properties helps users understand ParameterSet state
   - Context-aware formatting: colors as hex, sizes as pt, dimensions as mm

6. **Duplicate detection is critical after refactoring**
   - Count method definitions: `grep -c "def method_name" file.py`
   - Duplicates can silently override correct implementations
   - Second definition always wins in Python class definitions

### Pitfalls Encountered

1. ❌ Setting `self.attributes_names` → AttributeError (property has no setter)
2. ❌ Missing `_is_com` definition → NameError
3. ❌ Not checking for None backend → AttributeError
4. ❌ Manual attribute lists out of sync → Runtime errors
5. ❌ Duplicate class/method definitions → Second definition overrides first, causing bugs
6. ❌ Copy-paste refactoring without checking for duplicates → Hard-to-debug issues

## Recent Changes

### 2025-12-09 - Logging System Improvements
- **Default Log Level Changed**: Changed from `INFO` to `WARNING`
- Production-friendly: Normal users only see warnings, errors, and critical messages
- `HWPAPI_LOG_LEVEL` environment variable now defaults to `WARNING`
- Files: `hwpapi/logging.py`

### 2025-12-09 - Complete Display Enhancement Suite
- **Critical Bug Fix**: Removed duplicate ParameterSet class definition (27,304 chars duplicated)
- **Enhancement 1**: Human-Readable Value Formatting (colors, sizes, dimensions, booleans)
- **Enhancement 2**: Enum Display for MappedProperty (`Direction=0 (down)`)
- **Enhancement 3**: Property Description Comments (inline `# description`)
- Files: `hwpapi/parametersets.py`
- Examples: `nested_property_demo.py`, `mapped_property_display_demo.py`, `property_description_display_demo.py`

### 2025-01-08 - Auto-Creating Properties Design
- Designed `NestedProperty` for auto-creating nested ParameterSets
- Designed `ArrayProperty` for Pythonic HArray interface
- Enhanced `UnitProperty` for smart unit conversion (mm, cm, in, pt ↔ HWPUNIT)
- Created design documents: AUTO_PROPERTY_DESIGN.md, UNIT_PROPERTY_ENHANCEMENT.md
- Status: Design complete, ready for implementation

### 2025-01-08 - Architecture Analysis & Restructuring Plan
- Analyzed official HWP Automation Object Model (HwpAutomation_2504.pdf)
- Identified 4 major gaps: Collection objects, monolithic parametersets.py, form controls, organization
- Designed 3-phase restructuring plan (see `architecture-roadmap.md`)

### 2024 - Auto-Generated attributes_names
- Removed manual `self.attributes_names = [...]` from 9+ classes
- Added `@property attributes_names` to ParameterSet
- Fixed `_del_value` to handle None backend
- Result: -500 lines, eliminated sync issues

### 2024 - Added _is_com Function
- Fixed NameError in `_looks_like_pset`
- Added proper COM object detection

### Earlier - Pset Migration
- See PSET_MIGRATION_SUMMARY.md
- Migrated from HSet-based to pset-based approach
- Maintained backward compatibility
