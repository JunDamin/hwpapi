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

print("\nSetting string properties (PascalCase):")
find_replace.FindString = "Python"
find_replace.ReplaceString = "Python 3.11"
print(f"  FindString = '{find_replace.FindString}'")
print(f"  ReplaceString = '{find_replace.ReplaceString}'")

print("\nSetting boolean properties (PascalCase):")
find_replace.MatchCase = True
find_replace.WholeWordOnly = True
print(f"  MatchCase = {find_replace.MatchCase}")
print(f"  WholeWordOnly = {find_replace.WholeWordOnly}")

print("\nSetting mapped property (Direction - PascalCase):")
find_replace.Direction = "down"
print(f"  Direction = '{find_replace.Direction}' (maps to integer internally)")

print("\nType-safe properties (PascalCase):")
print("  - FindCharShape: TypedProperty (use create_itemset for nested access)")
print("  - FindParaShape: TypedProperty (use create_itemset for nested access)")
print("\n  NOTE: To use auto-creating NestedProperty, define CharShape/ParaShape")
print("        classes first, then use:")
print("        FindCharShape = NestedProperty('FindCharShape', 'CharShape', CharShape)")

# ============================================================================
# 3. TabDef with ArrayProperty
# ============================================================================
print("\n\n3. TabDef with ArrayProperty")
print("-" * 70)

# Create unbound TabDef (for demo - normally you'd get from action.CreateSet())
tab_def = TabDef()

print("\nSetting boolean properties (PascalCase):")
tab_def.AutoTabLeft = True
tab_def.AutoTabRight = False
print(f"  AutoTabLeft = {tab_def.AutoTabLeft}")
print(f"  AutoTabRight = {tab_def.AutoTabRight}")

print("\nUsing ArrayProperty (PascalCase - Pythonic list interface):")
print("  Creating tab stops: [1000, 0, 0, 2000, 0, 0, 3000, 0, 0]")
print("  (Each tab is 3 integers: position, fill type, tab type)")

# TabItem is an HArrayWrapper with list interface
tab_def.TabItem = [1000, 0, 0, 2000, 0, 0, 3000, 0, 0]

print(f"\n  TabItem = {tab_def.TabItem}")
print(f"  Length: {len(tab_def.TabItem)} elements")

print("\nList operations:")
print("  Appending fourth tab stop...")
tab_def.TabItem.append(4000)  # Position
tab_def.TabItem.append(0)     # Fill type
tab_def.TabItem.append(0)     # Tab type
print(f"  New length: {len(tab_def.TabItem)}")
print(f"  TabItem = {tab_def.TabItem}")

print("\n  Accessing by index:")
print(f"  TabItem[0] (first tab position) = {tab_def.TabItem[0]}")
print(f"  TabItem[3] (second tab position) = {tab_def.TabItem[3]}")

print("\n  Converting to plain list:")
plain_list = tab_def.TabItem.to_list()
print(f"  TabItem.to_list() = {plain_list}")

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
