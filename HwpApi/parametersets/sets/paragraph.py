"""
Paragraph-level ParameterSet classes (4 classes).

- TabDef: Tab stop definitions
- ListProperties: List-level formatting
- NumberingShape: Paragraph numbering (72 properties — Level0~Level6 series)
- ListParaPos: List paragraph positioning
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)
from hwpapi.parametersets.mappings import (
    LINE_WRAP_MAP, TEXT_DIRECTION_MAP, VERT_ALIGN_MAP,
)


class ListProperties(ParameterSet):
    """
    ### ListProperties

    76) ListProperties : 셀 리스트의 속성

    | Item ID       | Type     | SubType | Description                                      |
    |---------------|----------|---------|--------------------------------------------------|
    | TextDirection | PIT_UI1  |         | 글자 방향 (세로 쓰기 여부를 미정)                |
    | LineWrap      | PIT_UI1  |         | 강제에서 줄 바꿈 0 = 줄바꿈 없는 기본값           |
    | VertAlign     | PIT_UI1  |         | 세로 정렬 0 = 위 정렬                             |
    | MarginLeft    | PIT_I4   |         | 왼쪽 여백                                       |
    | MarginRight   | PIT_I4   |         | 오른쪽 여백                                     |
    | MarginTop     | PIT_I4   |         | 위쪽 여백                                       |
    | MarginBottom  | PIT_I4   |         | 아래쪽 여백                                     |
    """
    text_direction = ParameterSet._mapped_prop("TextDirection", TEXT_DIRECTION_MAP,
                                               doc="글자 방향: 0 = 기본값, 1 = 세로 쓰기")
    line_wrap = ParameterSet._mapped_prop("LineWrap", LINE_WRAP_MAP,
                                          doc="줄 바꿈 옵션: 0 = 기본값, 1 = 줄 바꿈 없음, 2 = 강제 줄 바꿈")
    vert_align = ParameterSet._mapped_prop("VertAlign", VERT_ALIGN_MAP,
                                           doc="세로 정렬: 0 = 위, 1 = 가운데, 2 = 아래")
    margin_left = ParameterSet._int_prop("MarginLeft", "왼쪽 여백: 정수 값을 입력하세요.")
    margin_right = ParameterSet._int_prop("MarginRight", "오른쪽 여백: 정수 값을 입력하세요.")
    margin_top = ParameterSet._int_prop("MarginTop", "위쪽 여백: 정수 값을 입력하세요.")
    margin_bottom = ParameterSet._int_prop("MarginBottom", "아래쪽 여백: 정수 값을 입력하세요.")



class NumberingShape(ParameterSet):
    """NumberingShape ParameterSet - 문단 번호 모양 정의. Level0~Level6 (7개 수준)."""
    StartNumber = PropertyDescriptor("StartNumber", r"""시작 번호 (0=앞 구역에 이어, n=지정한 번호로 시작)""")
    NewList = PropertyDescriptor("NewList", r"""새로운 번호 목록을 시작할지 여부""")
    # 7개 수준별 속성 (Level0 ~ Level6)
    HasCharShapeLevel0 = PropertyDescriptor("HasCharShapeLevel0", r"""수준0: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel1 = PropertyDescriptor("HasCharShapeLevel1", r"""수준1: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel2 = PropertyDescriptor("HasCharShapeLevel2", r"""수준2: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel3 = PropertyDescriptor("HasCharShapeLevel3", r"""수준3: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel4 = PropertyDescriptor("HasCharShapeLevel4", r"""수준4: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel5 = PropertyDescriptor("HasCharShapeLevel5", r"""수준5: 자체 글자 모양 사용 여부""")
    HasCharShapeLevel6 = PropertyDescriptor("HasCharShapeLevel6", r"""수준6: 자체 글자 모양 사용 여부""")
    CharShapeLevel0 = PropertyDescriptor("CharShapeLevel0", r"""수준0: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel1 = PropertyDescriptor("CharShapeLevel1", r"""수준1: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel2 = PropertyDescriptor("CharShapeLevel2", r"""수준2: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel3 = PropertyDescriptor("CharShapeLevel3", r"""수준3: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel4 = PropertyDescriptor("CharShapeLevel4", r"""수준4: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel5 = PropertyDescriptor("CharShapeLevel5", r"""수준5: 글자 모양 정의 (PIT_SET: CharShape)""")
    CharShapeLevel6 = PropertyDescriptor("CharShapeLevel6", r"""수준6: 글자 모양 정의 (PIT_SET: CharShape)""")
    WidthAdjustLevel0 = PropertyDescriptor("WidthAdjustLevel0", r"""수준0: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel1 = PropertyDescriptor("WidthAdjustLevel1", r"""수준1: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel2 = PropertyDescriptor("WidthAdjustLevel2", r"""수준2: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel3 = PropertyDescriptor("WidthAdjustLevel3", r"""수준3: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel4 = PropertyDescriptor("WidthAdjustLevel4", r"""수준4: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel5 = PropertyDescriptor("WidthAdjustLevel5", r"""수준5: 번호 너비 보정값 (HWPUNIT)""")
    WidthAdjustLevel6 = PropertyDescriptor("WidthAdjustLevel6", r"""수준6: 번호 너비 보정값 (HWPUNIT)""")
    TextOffsetLevel0 = PropertyDescriptor("TextOffsetLevel0", r"""수준0: 본문과의 거리""")
    TextOffsetLevel1 = PropertyDescriptor("TextOffsetLevel1", r"""수준1: 본문과의 거리""")
    TextOffsetLevel2 = PropertyDescriptor("TextOffsetLevel2", r"""수준2: 본문과의 거리""")
    TextOffsetLevel3 = PropertyDescriptor("TextOffsetLevel3", r"""수준3: 본문과의 거리""")
    TextOffsetLevel4 = PropertyDescriptor("TextOffsetLevel4", r"""수준4: 본문과의 거리""")
    TextOffsetLevel5 = PropertyDescriptor("TextOffsetLevel5", r"""수준5: 본문과의 거리""")
    TextOffsetLevel6 = PropertyDescriptor("TextOffsetLevel6", r"""수준6: 본문과의 거리""")
    AlignmentLevel0 = PropertyDescriptor("AlignmentLevel0", r"""수준0: 번호 정렬 0=왼쪽, 1=가운데, 2=오른쪽""")
    AlignmentLevel1 = PropertyDescriptor("AlignmentLevel1", r"""수준1: 번호 정렬""")
    AlignmentLevel2 = PropertyDescriptor("AlignmentLevel2", r"""수준2: 번호 정렬""")
    AlignmentLevel3 = PropertyDescriptor("AlignmentLevel3", r"""수준3: 번호 정렬""")
    AlignmentLevel4 = PropertyDescriptor("AlignmentLevel4", r"""수준4: 번호 정렬""")
    AlignmentLevel5 = PropertyDescriptor("AlignmentLevel5", r"""수준5: 번호 정렬""")
    AlignmentLevel6 = PropertyDescriptor("AlignmentLevel6", r"""수준6: 번호 정렬""")
    UseInstWidthLevel0 = PropertyDescriptor("UseInstWidthLevel0", r"""수준0: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel1 = PropertyDescriptor("UseInstWidthLevel1", r"""수준1: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel2 = PropertyDescriptor("UseInstWidthLevel2", r"""수준2: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel3 = PropertyDescriptor("UseInstWidthLevel3", r"""수준3: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel4 = PropertyDescriptor("UseInstWidthLevel4", r"""수준4: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel5 = PropertyDescriptor("UseInstWidthLevel5", r"""수준5: 번호 너비를 문자열 너비에 따를지 여부""")
    UseInstWidthLevel6 = PropertyDescriptor("UseInstWidthLevel6", r"""수준6: 번호 너비를 문자열 너비에 따를지 여부""")
    AutoIndentLevel0 = PropertyDescriptor("AutoIndentLevel0", r"""수준0: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel1 = PropertyDescriptor("AutoIndentLevel1", r"""수준1: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel2 = PropertyDescriptor("AutoIndentLevel2", r"""수준2: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel3 = PropertyDescriptor("AutoIndentLevel3", r"""수준3: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel4 = PropertyDescriptor("AutoIndentLevel4", r"""수준4: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel5 = PropertyDescriptor("AutoIndentLevel5", r"""수준5: 번호 너비 자동 들여쓰기 여부""")
    AutoIndentLevel6 = PropertyDescriptor("AutoIndentLevel6", r"""수준6: 번호 너비 자동 들여쓰기 여부""")
    TextOffsetTypeLevel0 = PropertyDescriptor("TextOffsetTypeLevel0", r"""수준0: 본문과의 거리 종류 0=퍼센트, 1=HWPUNIT""")
    TextOffsetTypeLevel1 = PropertyDescriptor("TextOffsetTypeLevel1", r"""수준1: 본문과의 거리 종류""")
    TextOffsetTypeLevel2 = PropertyDescriptor("TextOffsetTypeLevel2", r"""수준2: 본문과의 거리 종류""")
    TextOffsetTypeLevel3 = PropertyDescriptor("TextOffsetTypeLevel3", r"""수준3: 본문과의 거리 종류""")
    TextOffsetTypeLevel4 = PropertyDescriptor("TextOffsetTypeLevel4", r"""수준4: 본문과의 거리 종류""")
    TextOffsetTypeLevel5 = PropertyDescriptor("TextOffsetTypeLevel5", r"""수준5: 본문과의 거리 종류""")
    TextOffsetTypeLevel6 = PropertyDescriptor("TextOffsetTypeLevel6", r"""수준6: 본문과의 거리 종류""")
    StrFormatLevel0 = PropertyDescriptor("StrFormatLevel0", r"""수준0: 포맷 문자열""")
    StrFormatLevel1 = PropertyDescriptor("StrFormatLevel1", r"""수준1: 포맷 문자열""")
    StrFormatLevel2 = PropertyDescriptor("StrFormatLevel2", r"""수준2: 포맷 문자열""")
    StrFormatLevel3 = PropertyDescriptor("StrFormatLevel3", r"""수준3: 포맷 문자열""")
    StrFormatLevel4 = PropertyDescriptor("StrFormatLevel4", r"""수준4: 포맷 문자열""")
    StrFormatLevel5 = PropertyDescriptor("StrFormatLevel5", r"""수준5: 포맷 문자열""")
    StrFormatLevel6 = PropertyDescriptor("StrFormatLevel6", r"""수준6: 포맷 문자열""")
    NumFormatLevel0 = PropertyDescriptor("NumFormatLevel0", r"""수준0: 번호 모양""")
    NumFormatLevel1 = PropertyDescriptor("NumFormatLevel1", r"""수준1: 번호 모양""")
    NumFormatLevel2 = PropertyDescriptor("NumFormatLevel2", r"""수준2: 번호 모양""")
    NumFormatLevel3 = PropertyDescriptor("NumFormatLevel3", r"""수준3: 번호 모양""")
    NumFormatLevel4 = PropertyDescriptor("NumFormatLevel4", r"""수준4: 번호 모양""")
    NumFormatLevel5 = PropertyDescriptor("NumFormatLevel5", r"""수준5: 번호 모양""")
    NumFormatLevel6 = PropertyDescriptor("NumFormatLevel6", r"""수준6: 번호 모양""")


class TabDef(ParameterSet):
    """
    ### TabDef

    113) TabDef : Tab settings

    Tab definition with auto tab settings and tab stop positions.
    Uses ArrayProperty for Pythonic list interface to tab stops.

    | Item ID    | Type      | SubType | Description |
    |------------|-----------|---------|-------------|
    | AutoTabLeft| PIT_UI1   |         | Auto left tab (on / off) |
    | AutoTabRight| PIT_UI1  |         | Auto right tab (on / off) |
    | TabItem    | PIT_ARRAY | PIT_I   | Tab stop positions array (n*3+0: position, n*3+1: fill type, n*3+2: tab type) |

    Example:
        >>> tab_def = TabDef(action.CreateSet())
        >>> tab_def.AutoTabLeft = True
        >>> tab_def.TabItem = [1000, 0, 0, 2000, 0, 0, 3000, 0, 0]  # Three tab stops
        >>> tab_def.TabItem.append(4000)  # Add position for fourth tab
        >>> tab_def.TabItem.append(0)     # Fill type
        >>> tab_def.TabItem.append(0)     # Tab type
    """
    AutoTabLeft = BoolProperty("AutoTabLeft", "Auto left tab")
    AutoTabRight = BoolProperty("AutoTabRight", "Auto right tab")

    # Array property for tab stops (Pythonic list interface)
    # Each tab stop is represented by 3 integers: position, fill type, tab type
    TabItem = ArrayProperty("TabItem", int, "Tab stop data (position, fill, type triplets)")



class ListParaPos(ParameterSet):
    """ListParaPos ParameterSet."""
    List = PropertyDescriptor("List", r"""현재 위치한 리스트""")
    Para = PropertyDescriptor("Para", r"""현재 위치한 문단""")
    Pos = PropertyDescriptor("Pos", r"""현재 위치한 글자""")

