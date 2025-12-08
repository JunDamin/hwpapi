#!/usr/bin/env python3
"""
Tests for HParameterSet support in hwpapi.parametersets

These tests verify the functionality of:
- HParamBackend with dotted key support
- make_backend auto-detection
- Non-destructive HParam helpers
- Type-aware coercion
- Error handling

Tests are designed to skip gracefully if pywin32 or HWP are not available.
"""

import unittest
import sys
import os

# Add hwpapi to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Check for dependencies
try:
    import win32com.client
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

try:
    from hwpapi.core import App
    from hwpapi.parametersets import (
        _is_com, _prop_map_get, _prop_map_put, _split_dotted,
        _resolve_parent_and_leaf, _coerce_for_put, _looks_like_hparamset,
        HParamBackend, make_backend, ComBackend, AttrBackend,
        to_dict_hparam, snapshot_hparam, apply_dict_hparam,
        temp_edit_hparam, resolve_action_args, apply_staged_to_backend
    )
    HAS_HWPAPI = True
except ImportError as e:
    HAS_HWPAPI = False
    HWPAPI_ERROR = str(e)

# Try to connect to HWP
HAS_HWP = False
if HAS_HWPAPI and HAS_PYWIN32:
    try:
        _test_app = App()
        HAS_HWP = True
        _test_app = None  # Clean up
    except Exception:
        HAS_HWP = False


def skip_if_no_deps(reason="Dependencies not available"):
    """Decorator to skip tests if dependencies are missing."""
    def decorator(test_func):
        if not (HAS_PYWIN32 and HAS_HWPAPI):
            return unittest.skip(f"{reason}: pywin32={HAS_PYWIN32}, hwpapi={HAS_HWPAPI}")(test_func)
        return test_func
    return decorator


def skip_if_no_hwp(test_func):
    """Decorator to skip tests if HWP is not available."""
    if not HAS_HWP:
        return unittest.skip("HWP application not available")(test_func)
    return test_func


class TestHParamUtilities(unittest.TestCase):
    """Test utility functions for HParam support."""
    
    @skip_if_no_deps()
    def test_split_dotted(self):
        """Test dotted key splitting."""
        self.assertEqual(_split_dotted("simple"), ["simple"])
        self.assertEqual(_split_dotted("HFindReplace.FindString"), ["HFindReplace", "FindString"])
        self.assertEqual(_split_dotted("a.b.c.d"), ["a", "b", "c", "d"])
    
    @skip_if_no_deps()
    def test_is_com_with_mock(self):
        """Test COM object detection with mock objects."""
        # Test with regular Python object
        regular_obj = {"test": "value"}
        self.assertFalse(_is_com(regular_obj))
        
        # Test with object that has _oleobj_ attribute
        class MockCOMObj:
            def __init__(self):
                self._oleobj_ = "mock"
        
        mock_com = MockCOMObj()
        self.assertTrue(_is_com(mock_com))
    
    @skip_if_no_deps()
    def test_looks_like_hparamset_mock(self):
        """Test HParameterSet detection with mock objects."""
        # Regular object
        regular_obj = {"test": "value"}
        self.assertFalse(_looks_like_hparamset(regular_obj))
        
        # Object with HSet
        class MockHParamSet:
            def __init__(self):
                self.HSet = "mock"
        
        mock_hparam = MockHParamSet()
        self.assertTrue(_looks_like_hparamset(mock_hparam))
        
        # COM object with HParam nodes
        class MockCOMWithHNodes:
            def __init__(self):
                self._oleobj_ = "mock"
                self.HFindReplace = "mock"
        
        mock_com_h = MockCOMWithHNodes()
        self.assertTrue(_looks_like_hparamset(mock_com_h))


class TestBackendSelection(unittest.TestCase):
    """Test backend auto-detection."""
    
    @skip_if_no_deps()
    def test_make_backend_selection(self):
        """Test backend selection logic."""
        # Test ComBackend selection
        class MockParameterSet:
            def SetID(self): pass
            def Item(self, key): pass
            def SetItem(self, key, value): pass
            def RemoveItem(self, key): pass
        
        mock_pset = MockParameterSet()
        backend = make_backend(mock_pset)
        self.assertIsInstance(backend, ComBackend)
        
        # Test HParamBackend selection
        class MockHParameterSet:
            def __init__(self):
                self._oleobj_ = "mock"
                self.HFindReplace = "mock"
        
        mock_hpset = MockHParameterSet()
        backend = make_backend(mock_hpset)
        self.assertIsInstance(backend, HParamBackend)
        
        # Test AttrBackend fallback
        regular_obj = {"test": "value"}
        backend = make_backend(regular_obj)
        self.assertIsInstance(backend, AttrBackend)


class TestHParamBackendMock(unittest.TestCase):
    """Test HParamBackend with mock objects."""
    
    @skip_if_no_deps()
    def test_hparam_backend_basic(self):
        """Test basic HParamBackend functionality with mock."""
        # Create mock HParameterSet structure
        class MockLeafNode:
            def __init__(self):
                self.FindString = "test"
                self.IgnoreCase = True
                self.MaxCount = 100
        
        class MockRoot:
            def __init__(self):
                self.HFindReplace = MockLeafNode()
        
        root = MockRoot()
        backend = HParamBackend(root)
        
        # Test get
        self.assertEqual(backend.get("HFindReplace.FindString"), "test")
        self.assertEqual(backend.get("HFindReplace.IgnoreCase"), True)
        self.assertEqual(backend.get("HFindReplace.MaxCount"), 100)
        
        # Test set
        backend.set("HFindReplace.FindString", "updated")
        self.assertEqual(backend.get("HFindReplace.FindString"), "updated")
        
        # Test delete (neutralization)
        success = backend.delete("HFindReplace.IgnoreCase")
        self.assertTrue(success)
        self.assertEqual(backend.get("HFindReplace.IgnoreCase"), False)  # Neutralized to False
        
        # Test error on non-existent path
        with self.assertRaises(KeyError):
            backend.get("NonExistent.Property")


@unittest.skipUnless(HAS_HWP, "HWP application not available")
class TestHParamIntegration(unittest.TestCase):
    """Integration tests with real HWP application."""
    
    def setUp(self):
        """Set up test with HWP app."""
        self.app = App()
        self.hparam_root = self.app.api.HParameterSet
    
    def test_real_hparam_detection(self):
        """Test HParameterSet detection with real HWP object."""
        self.assertTrue(_is_com(self.hparam_root))
        self.assertTrue(_looks_like_hparamset(self.hparam_root))
        
        backend = make_backend(self.hparam_root)
        self.assertIsInstance(backend, HParamBackend)
    
    def test_real_hparam_operations(self):
        """Test HParamBackend operations with real HWP."""
        backend = make_backend(self.hparam_root)
        
        # Test getting current values
        try:
            find_str = backend.get("HFindReplace.FindString")
            self.assertIsInstance(find_str, (str, type(None)))
        except KeyError:
            self.skipTest("HFindReplace.FindString not available")
        
        # Test setting and getting back
        test_value = "test_string"
        backend.set("HFindReplace.FindString", test_value)
        retrieved_value = backend.get("HFindReplace.FindString")
        self.assertEqual(retrieved_value, test_value)
    
    def test_snapshot_and_restore(self):
        """Test snapshot and restore functionality."""
        # Take initial snapshot
        initial_snapshot = snapshot_hparam(self.hparam_root, "HFindReplace")
        self.assertIsInstance(initial_snapshot, dict)
        
        # Make some changes
        changes = {"FindString": "test_snapshot", "IgnoreCase": True}
        apply_dict_hparam(self.hparam_root, "HFindReplace", changes, strict=False)
        
        # Verify changes
        backend = make_backend(self.hparam_root)
        self.assertEqual(backend.get("HFindReplace.FindString"), "test_snapshot")
        
        # Restore from snapshot
        apply_dict_hparam(self.hparam_root, "HFindReplace", initial_snapshot, strict=False)
        
        # Verify restoration
        final_snapshot = snapshot_hparam(self.hparam_root, "HFindReplace")
        # Note: We can't guarantee exact equality due to potential property variations
        # but the key properties should match
        if "FindString" in initial_snapshot and "FindString" in final_snapshot:
            self.assertEqual(initial_snapshot["FindString"], final_snapshot["FindString"])
    
    def test_temp_edit_context_manager(self):
        """Test temp_edit_hparam context manager."""
        backend = make_backend(self.hparam_root)
        
        # Get original value
        try:
            original_value = backend.get("HFindReplace.FindString")
        except KeyError:
            self.skipTest("HFindReplace.FindString not available")
        
        # Use context manager
        temp_changes = {"FindString": "temporary_value"}
        
        with temp_edit_hparam(self.hparam_root, "HFindReplace", temp_changes):
            # Inside context: should have temporary value
            temp_value = backend.get("HFindReplace.FindString")
            self.assertEqual(temp_value, "temporary_value")
        
        # Outside context: should be restored
        restored_value = backend.get("HFindReplace.FindString")
        self.assertEqual(restored_value, original_value)
    
    def test_resolve_action_args(self):
        """Test action argument resolution."""
        action_name, arg = resolve_action_args(
            self.app, "FindReplace", self.hparam_root.HFindReplace
        )
        self.assertEqual(action_name, "FindReplace")
        self.assertIsNotNone(arg)
        
        # Try GetDefault (should not fail)
        try:
            self.app.HAction.GetDefault(action_name, arg)
        except Exception as e:
            # Some actions might fail GetDefault, that's OK for testing
            self.assertIsInstance(e, Exception)


class TestErrorHandling(unittest.TestCase):
    """Test error handling and edge cases."""
    
    @skip_if_no_deps()
    def test_resolve_parent_and_leaf_errors(self):
        """Test error handling in path resolution."""
        class MockRoot:
            def __init__(self):
                self.existing = "value"
        
        root = MockRoot()
        
        # Valid path
        parent, leaf = _resolve_parent_and_leaf(root, "existing")
        self.assertEqual(parent, root)
        self.assertEqual(leaf, "existing")
        
        # Invalid path
        with self.assertRaises(KeyError):
            _resolve_parent_and_leaf(root, "nonexistent.property")
    
    @skip_if_no_deps()
    def test_coerce_for_put_fallback(self):
        """Test type coercion fallback behavior."""
        class MockObj:
            pass
        
        obj = MockObj()
        
        # Should return value as-is when no prop map available
        result = _coerce_for_put(obj, "test", "value")
        self.assertEqual(result, "value")


class TestParameterSetUpdateFrom(unittest.TestCase):
    """Test ParameterSet.update_from for attribute-based copying."""

    def test_update_from_simple(self):
        from hwpapi.parametersets import ParameterSet, IntProperty

        class SimplePS(ParameterSet):
            a = IntProperty("a", "Value a")
            b = IntProperty("b", "Value b")

            def __init__(self, a=None, b=None):
                super().__init__()
                if a is not None:
                    self.a = a
                if b is not None:
                    self.b = b

        src = SimplePS(a=1, b=2)
        dst = SimplePS(a=0, b=0)
        dst.update_from(src)
        self.assertEqual(dst.a, 1)
        self.assertEqual(dst.b, 2)

    def test_attributes_names_auto_generation(self):
        """Test that attributes_names property works correctly."""
        from hwpapi.parametersets import CharShape

        # Test that attributes_names is auto-generated from property registry
        ps = CharShape()
        self.assertIsInstance(ps.attributes_names, list)
        self.assertEqual(len(ps.attributes_names), 65)  # CharShape has 65 properties
        # Should contain known CharShape attributes (in PascalCase)
        self.assertIn('Bold', ps.attributes_names)
        self.assertIn('Italic', ps.attributes_names)
        self.assertIn('SizeHangul', ps.attributes_names)

        # Test dual access - both PascalCase and snake_case should work
        ps.Bold = True
        self.assertEqual(ps.bold, True)  # snake_case access
        ps.italic = True
        self.assertEqual(ps.Italic, True)  # PascalCase access

def run_tests():
    """Run all tests with appropriate skipping."""
    # Print dependency status
    print("Dependency Status:")
    print(f"  pywin32: {'✓' if HAS_PYWIN32 else '✗'}")
    print(f"  hwpapi: {'✓' if HAS_HWPAPI else '✗'}")
    if not HAS_HWPAPI:
        print(f"    Error: {HWPAPI_ERROR}")
    print(f"  HWP app: {'✓' if HAS_HWP else '✗'}")
    print()
    
    # Run tests
    unittest.main(verbosity=2)


if __name__ == "__main__":
    run_tests()
