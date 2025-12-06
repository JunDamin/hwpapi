# ParameterSet Property Verification Report

## Summary

- **Total Classes**: 128
- **Classes with Mismatches**: 11
- **Classes Verified OK**: 117 (91%)

## Detailed Findings

### 1. BorderFill (25 missing properties)
**Status**: Manually created class is incomplete

Missing properties that exist in COM interface:
- Border properties: `BorderTypeLeft`, `BorderTypeRight`, `BorderTypeTop`, `BorderTypeBottom`
- Border widths: `BorderWidthLeft`, `BorderWidthRight`, `BorderWidthTop`, `BorderWidthBottom`
- Border colors: `BorderCorlorLeft` (note typo in docs), `BorderColorRight`, `BorderColorTop`, `BorderColorBottom`
- Diagonal properties: `SlashFlag`, `BackSlashFlag`, `DiagonalType`, `DiagonalWidth`, `DiagonalColor`
- Advanced: `CrookedSlashFlag`, `CrookedSlashFlag1`, `CrookedSlashFlag2`, `CounterSlashFlag`, `CounterBackSlashFlag`
- Other: `BorderFill3D`, `CenterLineFlag`, `BreakCellSeparateLine`

### 2. CharShape (35 missing, 2 extra properties)
**Status**: Manually created class is incomplete

Missing COM properties:
- Font types for all languages: `FontTypeHangul`, `FontTypeLatin`, `FontTypeHanja`, etc.
- Size ratios: `SizeHangul`, `SizeJapanese`, `SizeSymbol`
- Width ratios: `RatioHangul`, `RatioLatin`, `RatioHanja`, etc.
- Spacing: `SpacingHangul`, `SpacingLatin`, `SpacingHanja`, etc.
- Offsets: `OffsetHangul`, `OffsetLatin`, `OffsetHanja`, etc.
- Text effects: `UnderlineType`, `OutlineType`, `ShadowType`, `StrikeOutType`, `DiacSymMark`
- Shadow positioning: `ShadowOffsetX`, `ShadowOffsetY`
- Font features: `UseFontSpace`, `UseKerning`

Extra properties (not in docs):
- `StrikeOutColor`, `StrikeOutShape`

### 3. FindReplace (10 missing properties)
**Status**: Manually created class is incomplete

Missing COM properties:
- Search options: `Direction`, `WholeWordOnly`, `HanjaFromHangul`
- Ignore flags: `IgnoreFindString`, `IgnoreReplaceString`, `IgnoreMessage`
- Nested parameter sets: `FindCharShape`, `FindParaShape`, `ReplaceCharShape`, `ReplaceParaShape`

### 4. ParaShape (20 missing properties)
**Status**: Manually created class is incomplete

Missing COM properties:
- Alignment: `AlignType`, `TextAlignment`
- Line breaking: `BreakLatinWord`, `BreakNonLatinWord`
- Layout: `SnapToGrid`, `FontLineHeight`, `LineSpacingType`
- Paragraph protection: `KeepLinesTogether`, `KeepWithNext`, `PagebreakBefore`
- Border: `BorderConnect`, `BorderText`, `BorderOffsetLeft/Right/Top/Bottom`
- Heading: `HeadingType`, `TailType`
- Nested: `BorderFill`, `Numbering`

### 5. BulletShape (7 missing properties)
Missing: `Alignment`, `BulletImage`, `HasCharShape`, `HasImage`, `TextOffset`, `TextOffsetType`, `UseInstWidth`

### 6. Caption (2 missing properties)
Missing: `CapFullSize`, `Side`

### 7. Cell (2 missing properties)
Missing: `CellCtrlData`, `HasMargin`

### 8. DrawImageAttr (1 extra property)
Extra: `ImageAlpha` (not documented but exists in code)

### 9. NumberingShape
**Status**: Properties exist, just differently named in docs
- Docs show: `AlignmentLevel0~ AlignmentLevel6`
- Code has: `AlignmentLevel0`, `AlignmentLevel1`, ..., `AlignmentLevel6`
- This is actually correct - just different notation

### 10. ShapeObject (3 missing, 18 extra properties)
Missing: `NumberingType`, `OutsideMarginBottom`, `OutsideMarginTop`
Extra: Various `Shape*` prefixed properties like `ShapeCaption`, `ShapeDrawArcType`, etc.

### 11. Table (1 missing property)
Missing: `TableCharInfo`

## Root Cause Analysis

The mismatches fall into two categories:

1. **Pre-existing manually created classes** (BorderFill, CharShape, FindReplace, ParaShape, etc.)
   - These were created before the auto-generation process
   - They contain only the most commonly used properties
   - They are functional but incomplete

2. **Auto-generated classes** (98 new classes)
   - These match the documentation exactly
   - No mismatches found in any of the 98 newly added classes

## Recommendations

1. **High Priority**: Update the 11 manually created classes to include all documented properties
   - This ensures complete COM interface coverage
   - Maintains consistency with auto-generated classes

2. **Medium Priority**: Verify the "extra" properties in code
   - Check if `StrikeOutColor`, `StrikeOutShape`, `ImageAlpha` are real COM properties
   - Update documentation if they exist

3. **Low Priority**: Standardize property naming
   - Fix the typo `BorderCorlorLeft` → `BorderColorLeft` (if it exists in COM)
