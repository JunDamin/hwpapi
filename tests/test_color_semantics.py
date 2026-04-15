"""P0-2 regression tests: Color wrapper and UNSET sentinel semantics.

These tests pin down the intended behaviour of ``ColorProperty``:

- reading a color property always returns a ``Color`` instance (never raw int)
- ``prop = UNSET`` is a no-op (leaves the underlying value unchanged)
- ``prop = None`` removes the field
- writing a string/int/Color all work and normalize through ``convert_to_hwp_color``
- ``Color`` equality accepts string / int / None for backwards compatibility
"""
import unittest

from hwpapi.parametersets import Color, UNSET, CharShape


class TestColorClass(unittest.TestCase):
    def test_init_from_hex_string(self):
        c = Color("#FF0000")
        self.assertFalse(c.is_unset)
        self.assertIsNotNone(c.raw)

    def test_init_from_none_is_unset(self):
        c = Color(None)
        self.assertTrue(c.is_unset)
        self.assertIsNone(c.hex)
        self.assertIsNone(c.raw)

    def test_init_from_unset_is_unset(self):
        c = Color(UNSET)
        self.assertTrue(c.is_unset)

    def test_init_from_another_color(self):
        c1 = Color("#00FF00")
        c2 = Color(c1)
        self.assertEqual(c1, c2)
        self.assertEqual(c1.raw, c2.raw)

    def test_equality_with_string(self):
        c = Color("#FF0000")
        # Case-insensitive round-trip via convert_to_hwp_color
        self.assertTrue(c == "#FF0000" or c == "#ff0000")

    def test_equality_with_none(self):
        self.assertEqual(Color(None), None)
        self.assertNotEqual(Color("#000000"), None)

    def test_str_is_hex_or_empty(self):
        self.assertTrue(str(Color("#FF0000")).startswith("#"))
        self.assertEqual(str(Color(None)), "")

    def test_bool_semantics(self):
        self.assertTrue(bool(Color("#FF0000")))
        self.assertFalse(bool(Color(None)))

    def test_repr_indicates_unset(self):
        self.assertIn("UNSET", repr(Color(None)))

    def test_hashable(self):
        s = {Color("#FF0000"), Color("#FF0000"), Color(None)}
        self.assertEqual(len(s), 2)


class TestUnsetSentinel(unittest.TestCase):
    def test_unset_is_falsy(self):
        self.assertFalse(bool(UNSET))

    def test_unset_is_singleton(self):
        from hwpapi.parametersets.properties import _Unset
        a = _Unset()
        b = _Unset()
        self.assertIs(a, b)

    def test_unset_repr(self):
        self.assertEqual(repr(UNSET), "UNSET")

    def test_unset_is_not_none(self):
        self.assertIsNotNone(UNSET)


class TestColorPropertyOnCharShape(unittest.TestCase):
    """CharShape is a ParameterSet with ColorProperty fields (text_color, shade_color, ...)."""

    def test_read_returns_color_instance(self):
        cs = CharShape()
        cs.text_color = "#FF0000"
        val = cs.text_color
        self.assertIsInstance(val, Color)
        # Should NOT be a raw int or hex string
        self.assertFalse(isinstance(val, (int, str)))

    def test_set_with_string(self):
        cs = CharShape()
        cs.text_color = "#FF0000"
        self.assertFalse(cs.text_color.is_unset)

    def test_set_with_color_instance(self):
        cs = CharShape()
        cs.text_color = Color("#00FF00")
        self.assertFalse(cs.text_color.is_unset)

    def test_set_with_unset_is_noop(self):
        cs = CharShape()
        cs.text_color = "#FF0000"
        before = cs.text_color.raw
        cs.text_color = UNSET
        # UNSET should leave value unchanged
        self.assertEqual(cs.text_color.raw, before)

    def test_set_with_none_removes_field(self):
        cs = CharShape()
        cs.text_color = "#FF0000"
        cs.text_color = None
        # After None, the property should behave as if never set
        self.assertTrue(cs.text_color.is_unset)

    def test_roundtrip_preserves_value(self):
        cs = CharShape()
        cs.shade_color = "#FFFF00"
        # Now read back — should still be that color (structurally equal)
        self.assertEqual(cs.shade_color, "#FFFF00")

    def test_string_comparison_still_works(self):
        cs = CharShape()
        cs.text_color = "#FF0000"
        # Backwards-compatible: old code like `if cs.text_color == "#FF0000":`
        self.assertTrue(cs.text_color == "#FF0000" or cs.text_color == "#ff0000")


if __name__ == "__main__":
    unittest.main()
