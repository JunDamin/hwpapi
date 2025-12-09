"""
Demonstration of Property Description Display

This example shows how property descriptions are now shown as inline
comments in the ParameterSet representation, making it easy to understand
what each property means without checking documentation.
"""

from hwpapi.parametersets import ParameterSet, StringProperty, IntProperty, MappedProperty, ColorProperty, BoolProperty

print("="*70)
print("Property Description Display Demo")
print("="*70)

# Example 1: Video Parameters with Korean descriptions
print("\n" + "="*70)
print("1. Video Parameters (Korean Descriptions)")
print("="*70)

class VideoInsert(ParameterSet):
    """동영상 삽입 파라미터"""

    Base = StringProperty("Base", "동영상 파일의 경로")
    Format = MappedProperty("Format", {
        "mp4": 0,
        "avi": 1,
        "wmv": 2
    }, "동영상 형식")
    Width = IntProperty("Width", "동영상 너비 (HWPUNIT)")
    Height = IntProperty("Height", "동영상 높이 (HWPUNIT)")
    AutoPlay = BoolProperty("AutoPlay", "자동 재생 여부")

video = VideoInsert()
video.Base = "C:/Videos/presentation.mp4"
video.Format = "mp4"
video.Width = 59430  # 210mm
video.Height = 42050  # 149mm
video.AutoPlay = True

print("\nVideoInsert ParameterSet:")
print(video)
print("\nNotice: Each property shows its Korean description!")

# Example 2: Character Formatting with English descriptions
print("\n" + "="*70)
print("2. Character Formatting (English Descriptions)")
print("="*70)

class CharFormat(ParameterSet):
    """Character formatting parameters."""

    FontName = StringProperty("FontName", "Font family name")
    FontSize = IntProperty("FontSize", "Font size in HWPUNIT (100 = 1pt)")
    TextColor = ColorProperty("TextColor", "Text color in BBGGRR format")
    Bold = BoolProperty("Bold", "Bold formatting")
    Italic = BoolProperty("Italic", "Italic formatting")
    Underline = MappedProperty("Underline", {
        "none": 0,
        "single": 1,
        "double": 2,
        "dotted": 3
    }, "Underline style")

char = CharFormat()
char.FontName = "Arial"
char.FontSize = 1200  # 12pt
char.TextColor = 0x0000FF  # Red
char.Bold = True
char.Italic = False
char.Underline = "single"

print("\nCharFormat ParameterSet:")
print(char)
print("\nNotice: Descriptions explain units, format, and meaning!")

# Example 3: Table Parameters with detailed descriptions
print("\n" + "="*70)
print("3. Table Parameters (Detailed Descriptions)")
print("="*70)

class TableFormat(ParameterSet):
    """Table creation parameters."""

    Rows = IntProperty("Rows", "Number of rows (1-500)")
    Cols = IntProperty("Cols", "Number of columns (1-100)")
    Width = IntProperty("Width", "Table width in HWPUNIT (283 = 1mm)")
    CellSpacing = IntProperty("CellSpacing", "Space between cells in HWPUNIT")
    BorderWidth = IntProperty("BorderWidth", "Border line width in HWPUNIT")
    BorderColor = ColorProperty("BorderColor", "Border color (BBGGRR format)")
    Align = MappedProperty("Align", {
        "left": 0,
        "center": 1,
        "right": 2
    }, "Table alignment on page")

table = TableFormat()
table.Rows = 5
table.Cols = 3
table.Width = 42495  # 150mm
table.CellSpacing = 283  # 1mm
table.BorderWidth = 100  # ~0.35mm
table.BorderColor = 0x000000  # Black
table.Align = "center"

print("\nTableFormat ParameterSet:")
print(table)
print("\nNotice: Descriptions include units and valid ranges!")

# Example 4: Benefits
print("\n" + "="*70)
print("4. Benefits of Property Descriptions")
print("="*70)

print("""
Before Enhancement:
-------------------
CharFormat(
  FontSize=12.0pt
  TextColor="#ff0000"
  Bold=True
  ...
)
# Need to check docs: What is FontSize unit? What format is TextColor?

After Enhancement:
------------------
CharFormat(
  FontSize=12.0pt  # Font size in HWPUNIT (100 = 1pt)
  TextColor="#ff0000"  # Text color in BBGGRR format
  Bold=True  # Bold formatting
  ...
)
# Everything explained right there!

Why This Matters:
-----------------
[+] Self-Documenting: Understand parameters without checking docs
[+] Units Clarified: Know if it's HWPUNIT, pt, mm, etc.
[+] Format Explained: Understand BBGGRR, enum values, ranges
[+] Quick Reference: See valid values and meanings inline
[+] Learning: Discover what each parameter does while debugging
[+] Multilingual: Works with Korean, English, any language
[+] Context: Descriptions provide hints and constraints

Use Cases:
----------
- Video parameters: 동영상 파일의 경로, 동영상 형식
- Text formatting: Font size in HWPUNIT (100 = 1pt)
- Colors: Text color in BBGGRR format
- Enums: Underline style, Table alignment on page
- Ranges: Number of rows (1-500)
- Any property where context helps!
""")

# Example 5: Combined with Enum Display
print("\n" + "="*70)
print("5. Combined Enhancement: Enum Values + Descriptions")
print("="*70)

class FindOptions(ParameterSet):
    """Find and replace options."""

    FindString = StringProperty("FindString", "Text pattern to search for")
    ReplaceString = StringProperty("ReplaceString", "Replacement text")
    Direction = MappedProperty("Direction", {
        "down": 0,
        "up": 1,
        "all": 2
    }, "Search direction (down=forward, up=backward, all=entire doc)")
    MatchCase = BoolProperty("MatchCase", "Case-sensitive matching")
    WholeWordOnly = BoolProperty("WholeWordOnly", "Match whole words only")

find = FindOptions()
find.FindString = "Python"
find.ReplaceString = "Python 3.11"
find.Direction = "down"
find.MatchCase = True
find.WholeWordOnly = False

print("\nFindOptions ParameterSet:")
print(find)
print("\nNotice: Combines enum display (0 (down)) with descriptions!")
print("You see: Direction=0 (down)  # Search direction ...")

print("\n" + "="*70)
print("Summary")
print("="*70)
print("""
Three Display Enhancements Working Together:
---------------------------------------------
1. Human-readable values: 12.0pt, #ff0000, 210.0mm
2. Enum display: 0 (down), 1 (center), 2 (right)
3. Property descriptions: # Font size in HWPUNIT (100 = 1pt)

Result: Complete, self-documenting parameter display!

Example:
--------
FontSize=12.0pt  # Font size in HWPUNIT (100 = 1pt)
         ^^^^^     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         Value     Description explains it!

Direction=0 (down)  # Search direction (down=forward, up=backward)
          ^^^^^^^^    ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
          Enum        Description provides context!

All three enhancements make debugging and learning much easier!
""")

print("="*70)
print("Demo complete!")
print("="*70)
