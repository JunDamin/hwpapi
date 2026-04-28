"""
Text / character-level ParameterSet classes (20 classes).

Covers character manipulation operations:
- BulletShape, CodeTable, InsertText
- ConvertCase / ConvertFullHalf / ConvertHiraToGata / ConvertJianFan / ConvertToHangul
- InputHanja variants, ChCompose, ChangeRome, Dutmal
- SpellingCheck, QCorrect, TextVertical, MarkpenShape
- TextCtrl, AddHanjaWord
"""
from __future__ import annotations
from hwpapi.low.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class BulletShape(ParameterSet):
    """BulletShape ParameterSet."""
    HasCharShape = PropertyDescriptor("HasCharShape", r"""자체적인 글자 모양을 사용할지 여부 :0 \= 스타일을 따라감, 1 \= 자체 모양을 가짐""")
    CharShape = PropertyDescriptor("CharShape", r"""글자 모양 (HasCharShape가 1일 경우에만 사용)""")
    WidthAdjust = PropertyDescriptor("WidthAdjust", r"""번호 너비 보정 값 (HWPUNIT)""")
    TextOffset = PropertyDescriptor("TextOffset", r"""본문과의 거리 (percent or HWPUNIT)""")
    Alignment = PropertyDescriptor("Alignment", r"""번호 정렬 :0 \= 왼쪽 정렬, 1 \= 가운데 정렬, 2 \= 오른쪽 정렬""")
    UseInstWidth = PropertyDescriptor("UseInstWidth", r"""번호 너비를 문서 내 문자열의 너비에 따를지 여부 on / off""")
    AutoIndent = PropertyDescriptor("AutoIndent", r"""번호 너비 자동 들여쓰기 여부 : 0 \= 들여쓰기 안함, 1 \= 들여쓰기""")
    TextOffsetType = PropertyDescriptor("TextOffsetType", r"""본문과의 거리 종류 : 0 \= percent, 1 \= HWPUNIT""")
    BulletChar = PropertyDescriptor("BulletChar", r"""불릿 문자 코드""")
    HasImage = PropertyDescriptor("HasImage", r"""그림 글머리표 여부 : 0 \= 일반 글머리표, 1 \= 그림 글머리표""")
    BulletImage = PropertyDescriptor("BulletImage", r"""글머리표 이미지""")


class InsertText(ParameterSet):
    """

    ### InsertText 
| Item ID | Type | SubType | Description |
| --- | --- | --- | --- |
| Text | PIT_BSTR |  | 삽입할 텍스트 |

"""

    text = StringProperty("Text", "삽입할 텍스트")


class ChCompose(ParameterSet):
    """ChCompose ParameterSet."""
    Chars = PropertyDescriptor("Chars", r"""겹쳐질 글자들""")


class ChangeRome(ParameterSet):
    """ChangeRome ParameterSet."""
    Option = PropertyDescriptor("Option", r"""변환옵션 :0 \= 일반1 \= 주소2 \= 사람이름3 \= 한글복원""")
    HanString = PropertyDescriptor("HanString", r"""변경시킬 또는 변경된 한글문자""")


class CodeTable(ParameterSet):
    """CodeTable ParameterSet."""
    Text = PropertyDescriptor("Text", r"""삽입될 스트링""")


class ConvertCase(ParameterSet):
    """ConvertCase ParameterSet."""
    Type = PropertyDescriptor("Type", r"""공통사용.""")


class ConvertFullHalf(ParameterSet):
    """ConvertFullHalf ParameterSet."""
    Type = PropertyDescriptor("Type", r"""전각으로 변경할지, 반각으로 변경할지 여부 : 0 \= 반각, 1 \= 전각""")
    Number = PropertyDescriptor("Number", r"""변경 대상에 숫자를 추가할지 여부 : 0 \= off, 1 \= on""")
    Alpha = PropertyDescriptor("Alpha", r"""변경 대상에 영문자를 추가할지 여부 : 0 \= off, 1 \= on""")
    Symbol = PropertyDescriptor("Symbol", r"""변경 대상에 기호를 추가할지 여부 : 0 \= off, 1 \= on""")
    Gata = PropertyDescriptor("Gata", r"""변경 대상에 가타가나를 추가할지 여부 : 0 \= off, 1 \= on""")
    HGJamo = PropertyDescriptor("HGJamo", r"""변경 대상에 한글자모를 추가할지 여부 : 0 \= off, 1 \= on""")


class ConvertHiraToGata(ParameterSet):
    """ConvertHiraToGata ParameterSet."""
    Type = PropertyDescriptor("Type", r"""히라가나로 변경할지, 가타가나로 변경할지 여부0 \= 가타가나로 변경, 1 \= 히라가나로 변경""")


class ConvertJianFan(ParameterSet):
    """ConvertJianFan ParameterSet."""
    Type = PropertyDescriptor("Type", r"""간체로 변경할지, 번체로 변경할지 여부0 \= 번체로 변경, 1 \= 간체로 변경""")


class ConvertToHangul(ParameterSet):
    """ConvertToHangul ParameterSet."""
    Type = PropertyDescriptor("Type", r"""변경할 유형. 확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨.""")
    Hanja = PropertyDescriptor("Hanja", r"""한자를 한글로 변경할지 여부 : 0 \= off, 1 \= on""")
    Hira = PropertyDescriptor("Hira", r"""히라가나를 한글로 변경할지 여부 : 0 \= off, 1 \= on확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨.""")
    Gata = PropertyDescriptor("Gata", r"""가타가나를 한글로 변경할지 여부 : 0 \= off, 1 \= on확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨.""")
    Gu = PropertyDescriptor("Gu", r"""구결을 한글로 변경할지 여부 : 0 \= off, 1 \= on""")


class Dutmal(ParameterSet):
    """Dutmal ParameterSet."""
    ResultText = PropertyDescriptor("ResultText", r"""Parameter property""")
    SubText = PropertyDescriptor("SubText", r"""Parameter property""")
    PosType = PropertyDescriptor("PosType", r"""덧말 위치. 0 \= 위, 1 \= 아래.""")
    FsizeRatio = PropertyDescriptor("FsizeRatio", r"""덧말 크기 Percent. 0\=이면 기본 50%.""")
    Option = PropertyDescriptor("Option", r"""Parameter property""")
    StyleNo = PropertyDescriptor("StyleNo", r"""스타일번호 (옵션이 켜있을 때)""")
    Align = PropertyDescriptor("Align", r"""덧말의 좌우 Align. 0 \= 양쪽 정렬1 \= 왼쪽 정렬2 \= 오른쪽 정렬3 \= 가운데 정렬4 \= 배분 정렬5 \= 나눔 정렬기본은 가운데 정렬이며 옵션이 켜있을 때만 유효""")
    Delete = PropertyDescriptor("Delete", r"""덧말 지움 (한글2007에 새로 추가)""")
    Modify = PropertyDescriptor("Modify", r"""덧말 편집 모드 여부 (한글2007에 새로 추가)""")


class MarkpenShape(ParameterSet):
    """MarkpenShape ParameterSet."""
    Color = ColorProperty("Color", r"""형광펜색 (COLORREF)""")


class QCorrect(ParameterSet):
    """QCorrect ParameterSet."""
    LauncherKey = PropertyDescriptor("LauncherKey", r"""빠른 교정을 실행한 키 정보 (한글2007에 새로 추가)""")
    HyperLinkRunKey = PropertyDescriptor("HyperLinkRunKey", r"""URL 또는 email 하이퍼링크 작성 키 정보 (한글2007에 새로 추가)""")


class TextCtrl(ParameterSet):
    """TextCtrl ParameterSet."""
    CtrlData = PropertyDescriptor("CtrlData", r"""컨트롤 이름 저장을 위한 영역""")


class TextVertical(ParameterSet):
    """TextVertical ParameterSet."""
    Landscope = PropertyDescriptor("Landscope", r"""용지 방향. 0 \= 좁게, 1 \= 넓게""")
    TextDirection = PropertyDescriptor("TextDirection", r"""글자 방향.0 \= 보통 (왼쪽에서 오른쪽)1 \= 세로쓰기 (라틴 문자 회전)2 \= 세로쓰기 (라틴 문자 포함)""")
    TextVerticalWidthHead = PropertyDescriptor("TextVerticalWidthHead", r"""머리말/꼬리말 세로쓰기 여부""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용 대상0 \= 선택된 구역1 \= 선택된 문자열2 \= 현재 구역3 \= 문서전체4 \= 새 구역 : 현재 위치부터 새로5 \= no items (적용대상 없음)""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용 대상 분류.적용 대상 분류는 현재 캐럿의 상태에 따라 ApplyTo에 적용 가능한 대상을 한정짓는 역할을 한다. 내부적으로 값이 계산되므로, 값을 참조하는 용도로만 사용하도록 한다. 다음의 값의 조합으로 구성된다.0x0001 \= 선택된 구역0x0002 \= 선택된 문자열0x0004 \= 현재 구역0x0008 \= 문서 전체0x0010 \= 새 구역 : 현재 위치부터 새로""")


class AddHanjaWord(ParameterSet):
    """AddHanjaWord ParameterSet - 한자단어 등록."""


class InputHanja(ParameterSet):
    """InputHanja ParameterSet - 한자 변환."""


class InputHanjaBusu(ParameterSet):
    """InputHanjaBusu ParameterSet - 부수로 입력."""


class InputHanjaMean(ParameterSet):
    """InputHanjaMean ParameterSet - 새김으로 입력."""


class SpellingCheck(ParameterSet):
    """SpellingCheck ParameterSet - 맞춤법 검사."""


# ── Finalize attribute lookup tables ──────────────────────────────────────
# Populate snake_case entries in each class's _attr_lookup.
# This runs once at module import after all classes are defined.
