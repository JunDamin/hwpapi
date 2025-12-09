"""
Demonstration of MappedProperty Enhanced Display

This example shows how MappedProperty (enum-like) values now display
with both their numeric value and human-readable name.
"""

from hwpapi.parametersets import ParameterSet, MappedProperty, StringProperty, IntProperty

print("="*70)
print("MappedProperty Enhanced Display Demo")
print("="*70)

# Example 1: Direction enum (English)
print("\n" + "="*70)
print("1. Direction Enum (English)")
print("="*70)

class FindReplace(ParameterSet):
    """Find/Replace with direction enum."""

    Direction = MappedProperty("Direction", {
        "down": 0,
        "up": 1,
        "all": 2
    }, "Search direction")

    FindString = StringProperty("FindString", "Text to find")
    ReplaceString = StringProperty("ReplaceString", "Replacement text")

fr = FindReplace()
fr.Direction = "down"  # Can set using string name
fr.FindString = "Python"
fr.ReplaceString = "Python 3.11"

print("\nFindReplace ParameterSet:")
print(fr)
print("\nNotice: Direction shows as '0 (down)' - both value and name!")

# Example 2: Alignment enum
print("\n" + "="*70)
print("2. Alignment Enum")
print("="*70)

class ParaFormat(ParameterSet):
    """Paragraph formatting."""

    Align = MappedProperty("Align", {
        "left": 0,
        "center": 1,
        "right": 2,
        "justify": 3
    }, "Text alignment")

    LineSpacing = IntProperty("LineSpacing", "Line spacing percentage")

para = ParaFormat()
para.Align = "center"
para.LineSpacing = 160

print("\nParaFormat ParameterSet:")
print(para)
print("\nNotice: Align shows as '1 (center)'")

# Example 3: Setting with numeric value
print("\n" + "="*70)
print("3. Setting with Numeric Value")
print("="*70)

para2 = ParaFormat()
para2.Align = 2  # Set using numeric value
para2.LineSpacing = 140

print("\nParaFormat with numeric value (2):")
print(para2)
print("\nNotice: Still shows as '2 (right)' - auto-maps to name!")

# Example 4: All enum values
print("\n" + "="*70)
print("4. All Direction Values")
print("="*70)

directions = [
    ("down", 0),
    ("up", 1),
    ("all", 2)
]

print("\nIterating through all direction values:")
for name, value in directions:
    fr = FindReplace()
    fr.Direction = name
    fr.FindString = "test"

    # Show in repr
    lines = str(fr).split('\n')
    direction_line = [l for l in lines if 'Direction=' in l][0]
    print(f"  {direction_line.strip()}")

# Example 5: Benefits
print("\n" + "="*70)
print("5. Benefits of Enhanced Display")
print("="*70)

print("""
Before Enhancement:
-------------------
FindReplace(
  Direction="down"    # Only shows the name
  ...
)

After Enhancement:
------------------
FindReplace(
  Direction=0 (down)  # Shows BOTH value and name!
  ...
)

Why This Matters:
-----------------
[+] Debugging: Immediately see the numeric value used by HWP
[+] Understanding: See both the internal value and human meaning
[+] API Learning: Understand the mapping without checking docs
[+] Verification: Confirm the correct numeric value is being set
[+] Multilingual: Works with Korean, English, any language

Use Cases:
----------
- Bookmark types: 0=일반책갈피, 1=블록책갈피
- Text alignment: 0=left, 1=center, 2=right, 3=justify
- Search direction: 0=down, 1=up, 2=all
- Border styles: 0=none, 1=solid, 2=dashed, etc.
- Any parameter that maps strings to integers!
""")

print("="*70)
print("Demo complete!")
print("="*70)
