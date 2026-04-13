"""
Comprehensive tests for all ParameterSet classes.
These tests do NOT require HWP to be installed.
"""
import pytest
import inspect
from hwpapi.parametersets import (
    ParameterSet, GenericParameterSet, PARAMETERSET_REGISTRY,
    PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty,
    NestedProperty, ArrayProperty, ListProperty,
)
import hwpapi.parametersets as ps_mod


# ── Collect all ParameterSet subclasses ──────────────────────────────────

ALL_PS_CLASSES = []
for _name, _obj in inspect.getmembers(ps_mod, inspect.isclass):
    if issubclass(_obj, ParameterSet) and _obj is not ParameterSet:
        # Skip GenericParameterSet (requires positional arg)
        if _name == "GenericParameterSet":
            continue
        ALL_PS_CLASSES.append((_name, _obj))
ALL_PS_CLASSES.sort(key=lambda x: x[0])

# Classes that have properties defined
PS_WITH_PROPS = [(n, c) for n, c in ALL_PS_CLASSES if len(c._property_registry) > 0]


# ── A. Instantiation tests ──────────────────────────────────────────────

class TestParameterSetInstantiation:
    """Test that every ParameterSet subclass can be instantiated without error."""

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_instantiate(self, name, cls):
        """Each ParameterSet subclass should instantiate without arguments."""
        ps = cls()
        assert ps is not None

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_attributes_names_is_list(self, name, cls):
        """attributes_names should return a list."""
        ps = cls()
        names = ps.attributes_names
        assert isinstance(names, list)

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_property_registry_consistent(self, name, cls):
        """_property_registry keys should match attributes_names."""
        ps = cls()
        registry_keys = list(cls._property_registry.keys())
        attr_names = ps.attributes_names
        assert registry_keys == attr_names

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_repr_no_error(self, name, cls):
        """repr() should not raise."""
        ps = cls()
        r = repr(ps)
        assert isinstance(r, str)
        assert name in r or "ParameterSet" in r


# ── B. Property access tests ────────────────────────────────────────────

class TestPropertyAccess:
    """Test that all properties on all ParameterSet classes can be accessed."""

    @pytest.mark.parametrize("name,cls", PS_WITH_PROPS, ids=[n for n, _ in PS_WITH_PROPS])
    def test_get_all_properties(self, name, cls):
        """Reading every property should not raise (returns None for unbound)."""
        ps = cls()
        for prop_name in ps.attributes_names:
            descriptor = cls._property_registry.get(prop_name)
            # Skip NestedProperty/ArrayProperty - they require bound COM objects
            if isinstance(descriptor, (NestedProperty, ArrayProperty)):
                continue
            val = getattr(ps, prop_name)

    @pytest.mark.parametrize("name,cls", PS_WITH_PROPS, ids=[n for n, _ in PS_WITH_PROPS])
    def test_property_descriptors_valid(self, name, cls):
        """Every entry in _property_registry should be a PropertyDescriptor."""
        for prop_key, descriptor in cls._property_registry.items():
            assert isinstance(descriptor, PropertyDescriptor), (
                f"{name}.{prop_key} is {type(descriptor)}, not PropertyDescriptor"
            )


# ── C. Registry tests ───────────────────────────────────────────────────

class TestParameterSetRegistry:
    """Test PARAMETERSET_REGISTRY completeness."""

    def test_registry_not_empty(self):
        assert len(PARAMETERSET_REGISTRY) > 0

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_class_in_registry(self, name, cls):
        """Every subclass (except GenericParameterSet) should be registered."""
        if name == "GenericParameterSet":
            pytest.skip("GenericParameterSet is excluded from registry")
        assert name in PARAMETERSET_REGISTRY, f"{name} not in registry"
        assert PARAMETERSET_REGISTRY[name] is cls

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_lowercase_in_registry(self, name, cls):
        """Lowercase version should also be registered."""
        if name == "GenericParameterSet":
            pytest.skip("GenericParameterSet is excluded from registry")
        lower = name.lower()
        assert lower in PARAMETERSET_REGISTRY, f"{lower} not in registry"
        assert PARAMETERSET_REGISTRY[lower] is cls


# ── D. Property descriptor type tests ────────────────────────────────────

class TestPropertyDescriptorTypes:
    """Test that each property descriptor type works correctly."""

    def _make_ps_with(self, prop):
        """Create a ParameterSet subclass with a single property."""
        cls = type("TestPS", (ParameterSet,), {"test_prop": prop})
        return cls()

    def test_int_property(self):
        ps = self._make_ps_with(IntProperty("TestProp", "test int"))
        ps.test_prop = 42
        assert ps.test_prop == 42

    def test_bool_property(self):
        ps = self._make_ps_with(BoolProperty("TestProp", "test bool"))
        ps.test_prop = True
        assert ps.test_prop is True
        ps.test_prop = False
        assert ps.test_prop is False

    def test_string_property(self):
        ps = self._make_ps_with(StringProperty("TestProp", "test str"))
        ps.test_prop = "hello"
        assert ps.test_prop == "hello"

    def test_color_property(self):
        ps = self._make_ps_with(ColorProperty("TestProp", "test color"))
        ps.test_prop = "#FF0000"
        # ColorProperty stores as HWP color format

    def test_mapped_property(self):
        mapping = {"left": 0, "center": 1, "right": 2}
        ps = self._make_ps_with(MappedProperty("TestProp", mapping, "test mapped"))
        ps.test_prop = "center"
        assert ps.test_prop == "center"

    def test_list_property(self):
        ps = self._make_ps_with(ListProperty("TestProp", "test list"))
        ps.test_prop = [1, 2, 3]
        assert ps.test_prop == [1, 2, 3]

    def test_int_property_none_clear(self):
        ps = self._make_ps_with(IntProperty("TestProp", "test int"))
        ps.test_prop = 42
        ps.test_prop = None
        assert ps.test_prop is None

    def test_string_property_none_clear(self):
        ps = self._make_ps_with(StringProperty("TestProp", "test str"))
        ps.test_prop = "hello"
        ps.test_prop = None
        assert ps.test_prop is None


# ── E. GenericParameterSet tests ─────────────────────────────────────────

class TestGenericParameterSet:
    """Test GenericParameterSet special behavior."""

    def test_instantiate_with_none(self):
        ps = GenericParameterSet(parameterset=None)
        assert ps is not None

    def test_not_in_registry(self):
        assert "GenericParameterSet" not in PARAMETERSET_REGISTRY

    def test_attributes_names_empty(self):
        ps = GenericParameterSet(parameterset=None)
        assert ps.attributes_names == []


# ── F. Cross-class consistency tests ─────────────────────────────────────

class TestCrossClassConsistency:
    """Test consistency across all ParameterSet classes."""

    def test_total_class_count(self):
        """Verify we have the expected number of classes (excluding GenericParameterSet)."""
        assert len(ALL_PS_CLASSES) == 142

    def test_no_duplicate_class_names(self):
        names = [n for n, _ in ALL_PS_CLASSES]
        assert len(names) == len(set(names))

    def test_all_have_metaclass(self):
        """All ParameterSet subclasses should use ParameterSetMeta."""
        from hwpapi.parametersets import ParameterSetMeta
        for name, cls in ALL_PS_CLASSES:
            assert isinstance(cls, ParameterSetMeta), f"{name} doesn't use ParameterSetMeta"
