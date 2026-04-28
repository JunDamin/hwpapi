# Architecture Deep Dive

## Core Design Patterns

### 1. Backend Abstraction Pattern

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

### 2. Property Descriptor System

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
- **`UnitProperty`** - Smart unit conversion (mm, cm, in, pt ↔ HWPUNIT)
- `MappedProperty` - String ↔ Integer via mapping dict
- `TypedProperty` - Nested ParameterSet (manual)
- `ListProperty` - List of values (basic Python lists)
- **`NestedProperty`** - Auto-creating nested ParameterSet with tab completion
- **`ArrayProperty`** - Auto-creating HArray with list-like interface

**Auto-Generated Attributes:**
- `attributes_names` property returns `list(self._property_registry.keys())`
- ParameterSetMeta metaclass auto-populates `_property_registry`
- **NEVER** manually set `self.attributes_names = [...]` in subclasses

**Auto-Creating Properties (New Pattern):**
- `NestedProperty` and `ArrayProperty` automatically create underlying COM objects
- No manual `create_itemset()` or array initialization needed
- Full IDE tab completion and type hints
- Lazy creation on first access

### 3. Staging vs Immediate Mode

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

## Understanding the Domain

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
