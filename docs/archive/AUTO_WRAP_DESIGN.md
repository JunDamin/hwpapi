# Auto-Wrapping Design for Nested ParameterSets

## Problem Statement

When accessing nested ParameterSet properties, the current implementation returns raw COM objects instead of wrapped ParameterSet classes:

```python
pset = action.pset  # CharShape - wrapped ✓
nested = pset.BorderFill  # Raw COM object - not wrapped ✗
```

This breaks the abstraction and forces users to work with raw COM objects.

## Goals

1. **Automatic Wrapping**: Nested psets should be automatically wrapped
2. **Type Safety**: Known types should return their proper ParameterSet classes
3. **Robustness**: Handle new/unknown parameter sets from Hancom updates
4. **No Breaking Changes**: Existing code should continue to work
5. **Extensibility**: Easy to add new ParameterSet classes

## Proposed Solution: Three-Tier Auto-Wrapping System

### Tier 1: Known Types Registry

**Concept**: Maintain a global registry mapping pset IDs/names to ParameterSet classes.

```python
# Auto-populated registry
PARAMETERSET_REGISTRY = {
    "CharShape": CharShape,
    "ParaShape": ParaShape,
    "BorderFill": BorderFill,
    "ShapeObject": ShapeObject,
    # ... all defined ParameterSet classes
}
```

**Benefits**:
- Full type safety for known types
- All typed properties available
- Documentation and descriptions work
- IDE autocomplete works

**Implementation**:
- Use metaclass to auto-register all ParameterSet subclasses
- Registry is populated at import time
- No manual registration needed

### Tier 2: Generic ParameterSet Wrapper

**Concept**: For unknown pset types, create a generic wrapper that provides basic functionality.

```python
class GenericParameterSet(ParameterSet):
    """
    Generic wrapper for unknown ParameterSet types.
    Provides basic get/set functionality without typed properties.
    """

    def __init__(self, parameterset, pset_id=None):
        super().__init__(parameterset)
        self._pset_id = pset_id or "Unknown"

    def __getattr__(self, name):
        # Dynamic attribute access
        try:
            return self._backend.get(name)
        except:
            raise AttributeError(f"'{name}' not found in {self._pset_id}")

    def __setattr__(self, name, value):
        # Dynamic attribute setting
        if name.startswith('_'):
            super().__setattr__(name, value)
        else:
            self._backend.set(name, value)
```

**Benefits**:
- Doesn't break when Hancom adds new parameter sets
- Provides consistent interface
- Better than raw COM object
- Still allows get/set operations

**Trade-offs**:
- No typed properties
- No validation
- No documentation
- No IDE autocomplete

### Tier 3: Raw Access Fallback

**Concept**: If even generic wrapping fails, return raw object with helper methods.

```python
# Add helper methods to raw pset objects
def _wrap_unknown_pset(pset):
    """Add helper methods to completely unknown psets."""
    if not hasattr(pset, '_hwpapi_enhanced'):
        # Add convenience methods
        pset.get_item = lambda key: pset.Item(key)
        pset.set_item = lambda key, val: pset.SetItem(key, val)
        pset._hwpapi_enhanced = True
    return pset
```

**Benefits**:
- Never breaks, even for completely unknown types
- Adds convenience methods
- Users can still access everything

## Implementation Details

### 1. ParameterSet Metaclass Enhancement

```python
class ParameterSetMeta(type):
    """Enhanced metaclass that auto-registers ParameterSet classes."""

    def __new__(mcs, name, bases, namespace, **kwargs):
        cls = super().__new__(mcs, name, bases, namespace)

        # Auto-register in global registry
        if name != 'ParameterSet' and name != 'GenericParameterSet':
            # Get pset_id from class or use class name
            pset_id = namespace.get('_pset_id', name)
            PARAMETERSET_REGISTRY[pset_id] = cls

            # Also register by common variations
            PARAMETERSET_REGISTRY[name] = cls
            PARAMETERSET_REGISTRY[name.lower()] = cls

        return cls
```

### 2. Auto-Wrapping Function

```python
def wrap_parameterset(pset_obj, pset_id=None):
    """
    Automatically wrap a pset object in the appropriate class.

    Parameters
    ----------
    pset_obj : COM object
        The raw pset object to wrap
    pset_id : str, optional
        The pset ID/name if known

    Returns
    -------
    ParameterSet
        Wrapped pset object (specific class, generic, or enhanced raw)
    """
    # Already wrapped?
    if isinstance(pset_obj, ParameterSet):
        return pset_obj

    # Not a pset object?
    if not _looks_like_pset(pset_obj):
        return pset_obj

    # Try to get pset ID
    if pset_id is None:
        try:
            pset_id = pset_obj.GetIDStr() if hasattr(pset_obj, 'GetIDStr') else None
        except:
            pset_id = None

    # Tier 1: Look up in registry (known types)
    if pset_id and pset_id in PARAMETERSET_REGISTRY:
        cls = PARAMETERSET_REGISTRY[pset_id]
        logger.debug(f"Wrapping {pset_id} with {cls.__name__}")
        return cls(pset_obj)

    # Tier 2: Use generic wrapper (unknown types)
    logger.debug(f"Wrapping unknown pset {pset_id} with GenericParameterSet")
    return GenericParameterSet(pset_obj, pset_id=pset_id)
```

### 3. Property Descriptor Updates

Update all property descriptors to auto-wrap returned psets:

```python
class TypedProperty(PropertyDescriptor):
    def __get__(self, instance, owner):
        if instance is None:
            return self

        value = self._get_value(instance)
        if value is None:
            return self.default

        # Auto-wrap if it's a pset object
        if _looks_like_pset(value):
            return wrap_parameterset(value, self.property_type)

        return value
```

### 4. ItemSet Auto-Wrapping

When iterating ItemSets, auto-wrap each item:

```python
class ItemSetAccessor:
    def __getitem__(self, index):
        item = self._itemset.Item(index)
        # Auto-wrap the item
        return wrap_parameterset(item)

    def __iter__(self):
        for i in range(self._itemset.Count):
            yield wrap_parameterset(self._itemset.Item(i))
```

## Advantages of This Approach

### ✅ Robustness
- Handles known types perfectly
- Handles unknown types gracefully
- Never breaks on new Hancom updates

### ✅ Usability
- Consistent interface across all psets
- Natural Python attribute access
- No manual wrapping needed

### ✅ Extensibility
- New ParameterSet classes auto-register
- No code changes needed for new types
- Generic wrapper handles everything else

### ✅ Performance
- Registry lookup is O(1)
- Wrapping only happens once per access
- Negligible overhead

### ✅ Backward Compatibility
- Existing code continues to work
- No breaking changes
- Opt-in for advanced features

## Migration Path

### Phase 1: Foundation (Week 1)
1. Implement `wrap_parameterset()` function
2. Create `GenericParameterSet` class
3. Set up `PARAMETERSET_REGISTRY`
4. Enhance `ParameterSetMeta` with auto-registration

### Phase 2: Property Updates (Week 1-2)
1. Update `TypedProperty.__get__()` to auto-wrap
2. Update `NestedProperty.__get__()` to auto-wrap
3. Update `ListProperty` for ItemSet wrapping
4. Add helper methods to backends

### Phase 3: Testing & Refinement (Week 2)
1. Test with all known ParameterSet types
2. Test with unknown/hypothetical types
3. Performance benchmarking
4. Documentation updates

### Phase 4: Enhanced Features (Week 3)
1. Add `pset.to_dict()` for all types
2. Add `pset.from_dict()` for all types
3. Enhanced repr for generic psets
4. Logging for wrapped types

## Example Usage After Implementation

```python
# Known type - full functionality
char_shape = action.pset  # CharShape instance
char_shape.FaceNameHangul = "맑은 고딕"  # Typed property ✓
char_shape.Bold = True  # Validation ✓

# Nested known type - auto-wrapped
border = char_shape.BorderFill  # BorderFill instance (not raw COM) ✓
border.BorderType = "solid"  # Typed property ✓

# Unknown type (hypothetical new Hancom parameter) - generic wrapper
new_param = some_action.pset.NewParameter  # GenericParameterSet instance ✓
new_param.SomeSetting = 123  # Dynamic access ✓
value = new_param.SomeSetting  # Still works ✓

# Completely unknown - enhanced raw access
weird_param = some_other.WeirdThing  # Raw COM with helpers ✓
weird_param.get_item("key")  # Helper method added ✓
```

## Potential Issues and Solutions

### Issue 1: Performance Overhead
**Solution**: Wrap lazily, cache wrapped instances

### Issue 2: Circular References
**Solution**: Use weak references in registry

### Issue 3: ID Collisions
**Solution**: Use multiple registry keys (ID, name, class name)

### Issue 4: COM Object Lifecycle
**Solution**: Let Python GC handle it, no special cleanup needed

### Issue 5: Type Hints
**Solution**: Use `TypeVar` and `Generic` for proper typing

## Conclusion

This three-tier approach provides:
- **Best experience** for known types (Tier 1)
- **Good experience** for unknown types (Tier 2)
- **Working experience** for edge cases (Tier 3)

It's robust, extensible, and future-proof against Hancom updates while maintaining backward compatibility.

## Next Steps

1. Review this design with stakeholders
2. Implement Phase 1 (foundation)
3. Test with real HWP documents
4. Iterate based on feedback
5. Document for users
