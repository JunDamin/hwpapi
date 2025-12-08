"""
Demonstration of Auto-Creating Properties in hwpapi

This example shows how to use the new NestedProperty, ArrayProperty,
and enhanced UnitProperty features added to hwpapi.
"""

from hwpapi.parametersets import FindReplace, TabDef
from hwpapi.functions import parse_unit_string, to_hwpunit, from_hwpunit

print("="*70)
print("Auto-Creating Properties Demo")
print("="*70)

# ============================================================================
# 1. Enhanced Unit Conversion
# ============================================================================
print("\n1. Enhanced Unit Conversion")
print("-" * 70)

print("\nParsing unit strings:")
print(f"  parse_unit_string('210mm') = {parse_unit_string('210mm')}")
print(f"  parse_unit_string('21cm') = {parse_unit_string('21cm')}")
print(f"  parse_unit_string('8.27in') = {parse_unit_string('8.27in')}")
print(f"  parse_unit_string('12pt') = {parse_unit_string('12pt')}")

print("\nConverting to HWPUNIT:")
print(f"  to_hwpunit('210mm') = {to_hwpunit('210mm')}")
print(f"  to_hwpunit('21cm') = {to_hwpunit('21cm')}")
print(f"  to_hwpunit(210) = {to_hwpunit(210)} (defaults to mm)")

print("\nConverting from HWPUNIT:")
hwpunit_value = 59430  # A4 width
print(f"  from_hwpunit({hwpunit_value}, 'mm') = {from_hwpunit(hwpunit_value, 'mm')} mm")
print(f"  from_hwpunit({hwpunit_value}, 'cm') = {from_hwpunit(hwpunit_value, 'cm')} cm")
print(f"  from_hwpunit({hwpunit_value}, 'in') = {from_hwpunit(hwpunit_value, 'in')} in")

# ============================================================================
# 2. FindReplace with Type-Safe Properties
# ============================================================================
print("\n\n2. FindReplace with Type-Safe Properties")
print("-" * 70)

# Create unbound FindReplace (for demo - normally you'd get from action.CreateSet())
find_replace = FindReplace()

print("\nSetting string properties:")
find_replace.find_string = "Python"
find_replace.replace_string = "Python 3.11"
print(f"  find_string = '{find_replace.find_string}'")
print(f"  replace_string = '{find_replace.replace_string}'")

print("\nSetting boolean properties:")
find_replace.match_case = True
find_replace.whole_word_only = True
print(f"  match_case = {find_replace.match_case}")
print(f"  whole_word_only = {find_replace.whole_word_only}")

print("\nSetting mapped property (direction):")
find_replace.direction = "down"
print(f"  direction = '{find_replace.direction}' (maps to integer internally)")

print("\nType-safe properties:")
print("  - find_char_shape: TypedProperty (use create_itemset for nested access)")
print("  - find_para_shape: TypedProperty (use create_itemset for nested access)")
print("\n  NOTE: To use auto-creating NestedProperty, define CharShape/ParaShape")
print("        classes first, then use:")
print("        find_char_shape = NestedProperty('FindCharShape', 'CharShape', CharShape)")

# ============================================================================
# 3. TabDef with ArrayProperty
# ============================================================================
print("\n\n3. TabDef with ArrayProperty")
print("-" * 70)

# Create unbound TabDef (for demo - normally you'd get from action.CreateSet())
tab_def = TabDef()

print("\nSetting boolean properties:")
tab_def.auto_tab_left = True
tab_def.auto_tab_right = False
print(f"  auto_tab_left = {tab_def.auto_tab_left}")
print(f"  auto_tab_right = {tab_def.auto_tab_right}")

print("\nUsing ArrayProperty (Pythonic list interface):")
print("  Creating tab stops: [1000, 0, 0, 2000, 0, 0, 3000, 0, 0]")
print("  (Each tab is 3 integers: position, fill type, tab type)")

# tab_item is an HArrayWrapper with list interface
tab_def.tab_item = [1000, 0, 0, 2000, 0, 0, 3000, 0, 0]

print(f"\n  tab_item = {tab_def.tab_item}")
print(f"  Length: {len(tab_def.tab_item)} elements")

print("\nList operations:")
print("  Appending fourth tab stop...")
tab_def.tab_item.append(4000)  # Position
tab_def.tab_item.append(0)     # Fill type
tab_def.tab_item.append(0)     # Tab type
print(f"  New length: {len(tab_def.tab_item)}")
print(f"  tab_item = {tab_def.tab_item}")

print("\n  Accessing by index:")
print(f"  tab_item[0] (first tab position) = {tab_def.tab_item[0]}")
print(f"  tab_item[3] (second tab position) = {tab_def.tab_item[3]}")

print("\n  Converting to plain list:")
plain_list = tab_def.tab_item.to_list()
print(f"  tab_item.to_list() = {plain_list}")

# ============================================================================
# Summary
# ============================================================================
print("\n\n" + "="*70)
print("Summary")
print("="*70)
print("""
New Features Demonstrated:

1. Enhanced Unit Conversion
   - parse_unit_string(): Parse "210mm", "21cm", "8.27in", "12pt"
   - to_hwpunit(): Accept both numbers and unit strings
   - from_hwpunit(): Convert to any target unit (mm, cm, in, pt)

2. Type-Safe Property Descriptors
   - StringProperty: For text values
   - BoolProperty: For boolean flags
   - MappedProperty: String <-> Integer mapping with named values
   - TypedProperty: For nested ParameterSets (manual creation)
   - NestedProperty: Auto-creating nested ParameterSets (when classes defined)

3. Pythonic Array Interface
   - ArrayProperty: Wraps HArray with list interface
   - HArrayWrapper: Full list operations (append, insert, pop, etc.)
   - Type-safe element conversion
   - Index access and iteration

Benefits:
   - Full IDE tab completion support
   - Type safety and validation
   - Intuitive Pythonic API
   - Automatic unit conversions
   - Eliminates manual create_itemset() calls (with NestedProperty)

Next Steps:
   - Define CharShape, ParaShape classes to use NestedProperty
   - Update more ParameterSet classes with new property types
   - Use enhanced UnitProperty for all dimension properties
""")

print("="*70)
print("Demo complete!")
print("="*70)
