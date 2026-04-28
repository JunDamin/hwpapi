# Best Practices

## DO ✅

1. **Edit .py files directly** (this is a standard Python project)
2. **Test after every change** (at least run imports)
3. **Use property descriptors** for new ParameterSet attributes
4. **Check for None backend** in methods that access it
5. **Trust the backend factory** (make_backend)
6. **Follow existing patterns** in similar code
7. **Use type hints** (already set up with `from __future__ import annotations`)
8. **Keep backward compatibility** when refactoring

## DON'T ❌

1. **DON'T set `self.attributes_names = [...]` in subclasses**
2. **DON'T assume backend is always present** (can be None)
3. **DON'T mix staging modes** without understanding
4. **DON'T add features without tests**
5. **DON'T break existing API** without migration path
6. **DON'T use `isinstance` checks** unless necessary
7. **DON'T create new COM detection logic** (use `_is_com`)

## Development Workflow

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

## Simplification Strategies

### Successfully Applied: Auto-Generate attributes_names

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

## Checklist for New Contributors

Before making changes:
- [ ] Read CLAUDE.md and relevant claude_docs/ files
- [ ] Set up development environment
- [ ] Can run tests

Before committing:
- [ ] Edited .py files directly
- [ ] Ran tests (`python -m pytest tests/ -v`)
- [ ] Verified imports work
- [ ] Added/updated tests if needed
- [ ] Commit message describes what and why

## Remember

1. **Standard Python project**: .py files are the source of truth
2. **Backend abstraction works**: Trust `make_backend()` factory
3. **Properties are auto-registered**: No manual `attributes_names` needed
4. **Always check for None**: Backend might not be initialized
5. **Test your changes**: Don't break existing functionality
6. **Keep it simple**: Prefer simplification over clever abstractions
7. **Document decisions**: Update CLAUDE.md / claude_docs/ when you learn something new
