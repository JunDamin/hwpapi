# Common Issues and Solutions

## Issue 1: Missing Function Definitions

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

## Issue 2: AttributeError with attributes_names

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

## Issue 3: Backend is None

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

## Issue 4: Import Errors Between Modules

**Symptom:** Circular import or missing imports

**Solution:** Check module imports and definition order:
- Functions must be defined/imported before use
- Use `from typing import TYPE_CHECKING` for type-only imports

## Issue 5: Duplicate Class/Method Definitions

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
