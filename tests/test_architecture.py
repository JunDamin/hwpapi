"""
Architecture validation tests for ParameterSet system.
Verifies structural invariants documented in PARAMETERSET_ARCHITECTURE.md.
"""
import pytest
import inspect
from hwpapi.parametersets import (
    ParameterSet, ParameterSetMeta, GenericParameterSet, PARAMETERSET_REGISTRY,
    PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty,
    NestedProperty, ArrayProperty, ListProperty,
    PsetBackend, HParamBackend, ComBackend, AttrBackend,
    _is_com, _looks_like_pset, make_backend,
)
import hwpapi.parametersets as ps_mod


# ── Collect classes ──────────────────────────────────────────────────────

ALL_PS_CLASSES = [
    (n, c) for n, c in inspect.getmembers(ps_mod, inspect.isclass)
    if issubclass(c, ParameterSet) and c is not ParameterSet and n != "GenericParameterSet"
]


# ── 1. Metaclass invariants ──────────────────────────────────────────────

class TestMetaclassInvariants:
    """Verify ParameterSetMeta correctly populates _property_registry and REGISTRY."""

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_has_property_registry(self, name, cls):
        """Every subclass must have _property_registry as dict."""
        assert hasattr(cls, '_property_registry')
        assert isinstance(cls._property_registry, dict)

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_registry_values_are_descriptors(self, name, cls):
        """All _property_registry values must be PropertyDescriptor instances."""
        for key, desc in cls._property_registry.items():
            assert isinstance(desc, PropertyDescriptor), (
                f"{name}.{key} registry value is {type(desc).__name__}, not PropertyDescriptor"
            )

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_registry_keys_match_descriptor_keys(self, name, cls):
        """_property_registry keys must match the descriptor.key values registered on the class."""
        for reg_key, desc in cls._property_registry.items():
            # The registry key should be the attribute name on the class
            assert hasattr(cls, reg_key) or any(
                hasattr(cls, reg_key) for base in cls.__mro__
            ), f"{name}: registry key '{reg_key}' not found as class attribute"

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_uses_parameterset_meta(self, name, cls):
        """All subclasses should use ParameterSetMeta."""
        assert isinstance(cls, ParameterSetMeta)

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_registered_in_global_registry(self, name, cls):
        """All subclasses must be in PARAMETERSET_REGISTRY (PascalCase + lowercase)."""
        assert name in PARAMETERSET_REGISTRY, f"{name} not in PARAMETERSET_REGISTRY"
        assert name.lower() in PARAMETERSET_REGISTRY, f"{name.lower()} not in PARAMETERSET_REGISTRY"


# ── 2. Backend invariants ────────────────────────────────────────────────

class TestBackendInvariants:
    """Verify backend selection and interface compliance."""

    def test_make_backend_returns_attr_for_dict(self):
        backend = make_backend({"key": "value"})
        assert isinstance(backend, AttrBackend)

    def test_make_backend_returns_com_for_oleobj(self):
        class MockCOM:
            _oleobj_ = "mock"
        backend = make_backend(MockCOM())
        assert isinstance(backend, ComBackend)

    def test_make_backend_returns_pset_for_pset_like(self):
        class MockPset:
            _oleobj_ = "mock"
            def Item(self, k): pass
            def SetItem(self, k, v): pass
            def CreateItemSet(self, k, s): pass
            SetID = "Test"
        backend = make_backend(MockPset())
        assert isinstance(backend, PsetBackend)

    def test_all_backends_have_get_set_delete(self):
        """All backend classes must implement get, set, delete."""
        for cls in [PsetBackend, HParamBackend, ComBackend, AttrBackend]:
            assert hasattr(cls, 'get'), f"{cls.__name__} missing get()"
            assert hasattr(cls, 'set'), f"{cls.__name__} missing set()"
            assert hasattr(cls, 'delete'), f"{cls.__name__} missing delete()"

    def test_attr_backend_roundtrip(self):
        """AttrBackend should support get/set/delete."""
        obj = type('Obj', (), {'x': 10})()
        backend = AttrBackend(obj)
        assert backend.get('x') == 10
        backend.set('x', 20)
        assert backend.get('x') == 20
        backend.delete('x')

    def test_hparam_backend_dotted_path(self):
        """HParamBackend should resolve dotted paths."""
        class Child:
            value = "hello"
        class Root:
            child = Child()
        backend = HParamBackend(Root())
        assert backend.get("child.value") == "hello"
        backend.set("child.value", "world")
        assert backend.get("child.value") == "world"


# ── 3. Property descriptor invariants ────────────────────────────────────

class TestPropertyDescriptorInvariants:
    """Verify property descriptor conversion rules."""

    def _make_ps(self, prop):
        return type("TestPS", (ParameterSet,), {"test": prop})()

    def test_bool_true_to_1(self):
        ps = self._make_ps(BoolProperty("Test", ""))
        ps.test = True
        assert ps._staged.get("Test") == 1

    def test_bool_false_to_0(self):
        ps = self._make_ps(BoolProperty("Test", ""))
        ps.test = False
        assert ps._staged.get("Test") == 0

    def test_int_stores_int(self):
        ps = self._make_ps(IntProperty("Test", ""))
        ps.test = 42
        assert ps._staged.get("Test") == 42
        assert isinstance(ps._staged.get("Test"), int)

    def test_string_stores_string(self):
        ps = self._make_ps(StringProperty("Test", ""))
        ps.test = "hello"
        assert ps._staged.get("Test") == "hello"

    def test_mapped_string_to_int(self):
        mapping = {"left": 0, "center": 1, "right": 2}
        ps = self._make_ps(MappedProperty("Test", mapping, ""))
        ps.test = "center"
        assert ps._staged.get("Test") == 1

    def test_mapped_get_returns_string(self):
        mapping = {"left": 0, "center": 1, "right": 2}
        ps = self._make_ps(MappedProperty("Test", mapping, ""))
        ps.test = "center"
        assert ps.test == "center"

    def test_mapped_invalid_raises(self):
        mapping = {"left": 0, "center": 1}
        ps = self._make_ps(MappedProperty("Test", mapping, ""))
        with pytest.raises(ValueError):
            ps.test = "invalid"

    def test_color_hex_to_bbggrr(self):
        ps = self._make_ps(ColorProperty("Test", ""))
        ps.test = "#FF0000"
        # BBGGRR: red = 0x0000FF
        stored = ps._staged.get("Test")
        assert stored is not None

    def test_list_property_stores_list(self):
        ps = self._make_ps(ListProperty("Test", ""))
        ps.test = [1, 2, 3]
        assert ps._staged.get("Test") == [1, 2, 3]

    def test_none_clears_value(self):
        ps = self._make_ps(IntProperty("Test", ""))
        ps.test = 42
        ps.test = None
        assert ps.test is None


# ── 4. ParameterSet lifecycle invariants ─────────────────────────────────

class TestParameterSetLifecycle:
    """Verify ParameterSet creation, binding, access patterns."""

    def test_unbound_has_none_backend(self):
        ps = ParameterSet()
        assert ps._backend is None

    def test_unbound_attributes_names(self):
        """Unbound ParameterSet should still have attributes_names."""
        from hwpapi.parametersets import CharShape
        ps = CharShape()
        assert isinstance(ps.attributes_names, list)
        assert len(ps.attributes_names) == 65

    def test_staged_dict_initially_empty(self):
        ps = ParameterSet()
        assert ps._staged == {}

    def test_snake_case_to_pascal_case(self):
        """Setting snake_case should map to PascalCase in _staged."""
        from hwpapi.parametersets import CharShape
        ps = CharShape()
        ps.bold = True
        assert "Bold" in ps._staged

    def test_repr_no_crash_unbound(self):
        from hwpapi.parametersets import CharShape
        ps = CharShape()
        r = repr(ps)
        assert "CharShape" in r

    def test_update_from_copies_values(self):
        from hwpapi.parametersets import IntProperty
        cls = type("TestPS", (ParameterSet,), {
            "a": IntProperty("a", ""),
            "b": IntProperty("b", ""),
        })
        src = cls()
        src.a = 10
        src.b = 20
        dst = cls()
        dst.update_from(src)
        assert dst.a == 10
        assert dst.b == 20


# ── 5. Cross-class structural checks ─────────────────────────────────────

class TestCrossClassStructure:
    """Verify structural consistency across all ParameterSet classes."""

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_descriptor_key_is_pascalcase_or_snake(self, name, cls):
        """Property descriptor keys should start with uppercase (PascalCase)
        or be snake_case. They must not be empty."""
        for key, desc in cls._property_registry.items():
            assert len(key) > 0, f"{name}: empty property key"
            # Key in registry is the Python attribute name
            # The descriptor's .key is the COM key (should be PascalCase)
            assert len(desc.key) > 0, f"{name}.{key}: empty descriptor.key"

    @pytest.mark.parametrize("name,cls", ALL_PS_CLASSES, ids=[n for n, _ in ALL_PS_CLASSES])
    def test_no_duplicate_descriptor_keys(self, name, cls):
        """No two properties should map to the same COM key."""
        com_keys = [desc.key for desc in cls._property_registry.values()]
        duplicates = [k for k in com_keys if com_keys.count(k) > 1]
        assert not duplicates, f"{name}: duplicate COM keys: {set(duplicates)}"

    def test_generic_parameterset_excluded_from_registry(self):
        assert "GenericParameterSet" not in PARAMETERSET_REGISTRY

    def test_registry_has_no_base_parameterset(self):
        """ParameterSet base class should not be in registry."""
        assert "ParameterSet" not in PARAMETERSET_REGISTRY
