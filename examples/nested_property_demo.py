"""
Demonstration of NestedProperty Auto-Creation

This example shows how NestedProperty automatically creates nested
ParameterSet instances when accessed, eliminating manual create_itemset() calls.
"""

from hwpapi.parametersets import FindReplace, CharShape, ParaShape

print("="*70)
print("NestedProperty Auto-Creation Demo")
print("="*70)

print("""
Overview:
---------
NestedProperty automatically calls backend.create_itemset() on first access,
wraps the result in the specified ParameterSet class, and caches it for
subsequent access.

This eliminates the need for manual create_itemset() calls and provides
full IDE tab completion support!
""")

# ============================================================================
# 1. The Old Way (Manual create_itemset)
# ============================================================================
print("\n" + "="*70)
print("1. OLD WAY - Manual create_itemset (Verbose)")
print("="*70)
print("""
# Before: Manual, verbose, no tab completion
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")  # Manual
char_shape = CharShape(char_com)  # Manual wrapping
char_shape.Bold = True  # Finally set the value!
char_shape.Size = 1200

# Problems:
# - 4 steps to set a single property
# - No tab completion after create_itemset()
# - Easy to make mistakes with SetID names
# - Verbose and repetitive
""")

# ============================================================================
# 2. The New Way (NestedProperty Auto-Creation)
# ============================================================================
print("\n" + "="*70)
print("2. NEW WAY - NestedProperty Auto-Creation (Simple!)")
print("="*70)
print("""
# After: Auto-creating, clean, tab completion works!
pset = FindReplace(action.CreateSet())

# Direct access - auto-creates on first access!
pset.FindCharShape.Bold = True  # Auto-creates CharShape!
pset.FindCharShape.Size = 1200  # Uses cached instance

# Benefits:
# - Single line to set properties
# - Full IDE tab completion
# - Type-safe (CharShape class)
# - Automatic caching (same instance on subsequent access)
# - No manual create_itemset() needed!
""")

# ============================================================================
# 3. How It Works
# ============================================================================
print("\n" + "="*70)
print("3. How NestedProperty Works")
print("="*70)
print("""
Class Definition:
-----------------
class FindReplace(ParameterSet):
    # Auto-creating nested properties
    FindCharShape = NestedProperty("FindCharShape", "CharShape", CharShape)
    FindParaShape = NestedProperty("FindParaShape", "ParaShape", ParaShape)
    ReplaceCharShape = NestedProperty("ReplaceCharShape", "CharShape", CharShape)
    ReplaceParaShape = NestedProperty("ReplaceParaShape", "ParaShape", ParaShape)

First Access:
-------------
>>> pset.FindCharShape.Bold = True

What happens behind the scenes:
1. NestedProperty.__get__() is called
2. Checks cache - not found (first access)
3. Calls backend.create_itemset("FindCharShape", "CharShape")
4. Wraps result in CharShape class
5. Caches wrapped instance for future access
6. Returns CharShape instance
7. Sets Bold property on the CharShape instance

Second Access:
--------------
>>> pset.FindCharShape.Size = 1200

What happens:
1. NestedProperty.__get__() is called
2. Finds cached instance
3. Returns cached CharShape instance immediately (no create_itemset!)
4. Sets Size property on the cached instance
""")

# ============================================================================
# 4. Demo with Mock Backend
# ============================================================================
print("\n" + "="*70)
print("4. Testing (Unbound Mode)")
print("="*70)

# In unbound mode, NestedProperty correctly requires a backend
fr = FindReplace()
print("\nTrying to access nested property without backend:")
print("  >>> fr = FindReplace()")
print("  >>> fr.FindCharShape")
try:
    char_shape = fr.FindCharShape
    print(f"  ERROR: Should have raised RuntimeError!")
except RuntimeError as e:
    print(f"  [OK] Correctly raised RuntimeError:")
    print(f"       '{e}'")

print("\nThis is expected behavior!")
print("NestedProperty requires a backend (from action.CreateSet())")
print("With a real backend, it auto-creates the nested ParameterSet!")

# ============================================================================
# 5. Real Usage Pattern
# ============================================================================
print("\n" + "="*70)
print("5. Real Usage Pattern")
print("="*70)
print("""
Complete example with real HWP application:

from hwpapi import App

app = App()

# Create FindReplace action
pset = app.actions.repeat_find.create_set()

# Set simple properties
pset.FindString = "Python"
pset.ReplaceString = "Python 3.11"
pset.MatchCase = True
pset.WholeWordOnly = False
pset.Direction = "down"

# Set nested character shape properties - auto-creates!
pset.FindCharShape.Bold = True
pset.FindCharShape.Italic = False
pset.FindCharShape.Size = 1200  # 12pt
pset.FindCharShape.TextColor = "#FF0000"  # Red

# Set nested paragraph shape properties - auto-creates!
pset.FindParaShape.Align = "center"
pset.FindParaShape.LineSpacing = 160  # 160%

# Execute the action
result = pset.run()
print(f"Replaced {result} occurrences")
""")

# ============================================================================
# 6. Comparison
# ============================================================================
print("\n" + "="*70)
print("6. Side-by-Side Comparison")
print("="*70)
print("""
Task: Set Bold=True on FindCharShape

OLD WAY (TypedProperty):
------------------------
pset = FindReplace(action.CreateSet())
char_com = pset.create_itemset("FindCharShape", "CharShape")
char_shape = CharShape(char_com)
char_shape.Bold = True

NEW WAY (NestedProperty):
--------------------------
pset = FindReplace(action.CreateSet())
pset.FindCharShape.Bold = True

Lines of code: 4 -> 1
Tab completion: NO -> YES
Type safety: Manual -> Automatic
Caching: Manual -> Automatic
""")

# ============================================================================
# Summary
# ============================================================================
print("\n" + "="*70)
print("Summary")
print("="*70)
print("""
NestedProperty Benefits:
------------------------
[+] Auto-creates nested ParameterSets on first access
[+] Automatic caching (same instance on subsequent access)
[+] Full IDE tab completion support
[+] Type-safe property access
[+] Eliminates manual create_itemset() calls
[+] Cleaner, more Pythonic code
[+] Less error-prone

How to Use:
-----------
1. Define ParameterSet class with NestedProperty
2. Create instance from action.CreateSet()
3. Access nested properties directly - they auto-create!
4. No manual create_itemset() needed!

Currently Supported:
--------------------
FindReplace.FindCharShape    → CharShape (auto-creates)
FindReplace.FindParaShape    → ParaShape (auto-creates)
FindReplace.ReplaceCharShape → CharShape (auto-creates)
FindReplace.ReplaceParaShape → ParaShape (auto-creates)

You can add NestedProperty to any ParameterSet class that has
nested parameter sets!
""")

print("="*70)
print("Demo complete!")
print("="*70)
