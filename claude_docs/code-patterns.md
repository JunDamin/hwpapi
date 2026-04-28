# Code Patterns to Follow

## Pattern 1: Adding a New ParameterSet Class

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

## Pattern 2: Auto-Creating Nested ParameterSets

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

## Pattern 3: Adding a Custom Property Type

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

## Pattern 4: Checking COM Objects

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

## Pattern 5: Handling Optional Backend

```python
def my_method(self):
    """Method that accesses backend."""
    if self._backend is None:
        # Handle unbound case
        return default_value

    # Proceed with backend operations
    return self._backend.get(self.key)
```
