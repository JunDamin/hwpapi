"""
Style / border / formatting ParameterSet classes (3 classes).

- BorderFillExt: Extended border/fill properties
- StyleDelete: Style removal operations
- StyleTemplate: Style template operations

Note: ``BorderFill`` and ``Style`` are in ``primitives.py`` (used widely).
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class BorderFillExt(ParameterSet):
    """BorderFillExt ParameterSet."""
    TypeHorz = PropertyDescriptor("TypeHorz", r"""중앙선 종류 : 가로 \[선 종류]""")
    TypeVert = PropertyDescriptor("TypeVert", r"""중앙선 종류 : 세로""")
    WidthHorz = PropertyDescriptor("WidthHorz", r"""중앙선 두께 : 가로 \[선 굵기]""")
    WidthVert = PropertyDescriptor("WidthVert", r"""중앙선 두께 : 세로""")
    ColorHorz = ColorProperty("ColorHorz", r"""중앙선 색깔 : 가로RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    ColorVert = ColorProperty("ColorVert", r"""중앙선 색깔 : 세로RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")


class StyleDelete(ParameterSet):
    """StyleDelete ParameterSet."""
    Target = PropertyDescriptor("Target", r"""지워야할 스타일 인덱스""")
    Alternation = PropertyDescriptor("Alternation", r"""대체할 스타일 인덱스""")


class StyleTemplate(ParameterSet):
    """StyleTemplate ParameterSet."""
    FileName = PropertyDescriptor("FileName", r"""파일 이름""")

