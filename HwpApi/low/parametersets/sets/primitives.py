"""
ParameterSet primitives — foundation classes used across multiple domains.

These 8 classes are the building blocks that other ParameterSet classes
reference (via docstrings, and in one case Python-level):

- CharShape (65 properties): Character formatting. Referenced by NumberingShape,
  BulletShape, FindReplace (Python-level import).
- ParaShape (32 properties): Paragraph formatting. Referenced by FindReplace.
- BorderFill (27 properties): Border/background settings.
- Cell (8 properties): Table cell info.
- Caption (4 properties): Caption settings.
- CtrlData (1 property): Generic control data.
- Password (5 properties): Document security.
- Style (1 property): Style reference.

All classes here auto-register to ``PARAMETERSET_REGISTRY`` via
``ParameterSetMeta``.
"""
from __future__ import annotations
from hwpapi.low.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class BorderFill(ParameterSet):
    """BorderFill ParameterSet."""
    BorderTypeLeft = PropertyDescriptor("BorderTypeLeft", r"""4방향 테두리 종류 : 왼쪽 \[선 종류]""")
    BorderTypeRight = PropertyDescriptor("BorderTypeRight", r"""4방향 테두리 종류 : 오른쪽""")
    BorderTypeTop = PropertyDescriptor("BorderTypeTop", r"""4방향 테두리 종류 : 위""")
    BorderTypeBottom = PropertyDescriptor("BorderTypeBottom", r"""4방향 테두리 종류 : 아래""")
    BorderWidthLeft = PropertyDescriptor("BorderWidthLeft", r"""4방향 테두리 두께 : 왼쪽 \[선 굵기]""")
    BorderWidthRight = PropertyDescriptor("BorderWidthRight", r"""4방향 테두리 두께 : 오른쪽""")
    BorderWidthTop = PropertyDescriptor("BorderWidthTop", r"""4방향 테두리 두께 : 위""")
    BorderWidthBottom = PropertyDescriptor("BorderWidthBottom", r"""4방향 테두리 두께 : 아래""")
    BorderCorlorLeft = ColorProperty("BorderCorlorLeft", r"""4방향 테두리 색깔 : 왼쪽RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    BorderColorRight = ColorProperty("BorderColorRight", r"""4방향 테두리 색깔 : 오른쪽RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    BorderColorTop = ColorProperty("BorderColorTop", r"""4방향 테두리 색깔 : 위RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    BorderColorBottom = ColorProperty("BorderColorBottom", r"""4방향 테두리 색깔 :아래RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    SlashFlag = PropertyDescriptor("SlashFlag", r"""슬래쉬 대각선 플래그 : 비트 플래그의 조합으로 표현되며 각 위치의 비트는 다음을 나타낸다.bit 0 \= 하단 대각선bit 1 \= 중앙 대각선bit 2 \= 상단 대각선더 자세한 내용은 하단의 ※ 대각선의 형태를 참고한다.""")
    BackSlashFlag = PropertyDescriptor("BackSlashFlag", r"""백슬래쉬 대각선 플래그 : 비트 플래그의 조합으로 표현되며 각 위치의 비트는 다음을 나타낸다.bit 0 \= 하단 대각선bit 1 \= 중앙 대각선bit 2 \= 상단 대각선더 자세한 내용은 하단의 ※ 대각선의 형태를 참고한다.""")
    DiagonalType = PropertyDescriptor("DiagonalType", r"""선 종류 셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용""")
    DiagonalWidth = PropertyDescriptor("DiagonalWidth", r"""선 굵기 셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용""")
    DiagonalColor = ColorProperty("DiagonalColor", r"""선 색상RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용""")
    BorderFill3D = PropertyDescriptor("BorderFill3D", r"""3차원 효과 : 0 \= off, 1 \= on""")
    Shadow = PropertyDescriptor("Shadow", r"""그림자 효과 : 0 \= off, 1 \= on""")
    FillAttr = PropertyDescriptor("FillAttr", r"""배경 채우기 속성""")
    CrookedSlashFlag = PropertyDescriptor("CrookedSlashFlag", r"""꺾인 대각선 플래그 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺾인 대각선임을 나타낸다.)""")
    BreakCellSeparateLine = PropertyDescriptor("BreakCellSeparateLine", r"""자동으로 나뉜 표의 경계선 설정 :0 \= 경계선 설정을 기본 값에 따름, 1 \= 사용자가 경계선 설정""")
    CounterSlashFlag = PropertyDescriptor("CounterSlashFlag", r"""슬래쉬 대각선의 역방향 플래그(우상향 대각선) :0 \= 순방향, 1 \= 역방향""")
    CounterBackSlashFlag = PropertyDescriptor("CounterBackSlashFlag", r"""역슬래쉬 대각선의 역방향 플래그(좌상향 대각선) :0 \= 순방향, 1 \= 역방향""")
    CenterLineFlag = PropertyDescriptor("CenterLineFlag", r"""중심선 : ( 0 \= 중심선 없음, 1 \= 중심선 있음)더 자세한 내용은 하단의 ※ 중심선의 형태를 참고한다.""")
    CrookedSlashFlag1 = PropertyDescriptor("CrookedSlashFlag1", r"""Low Byte CrookedSlashFlag (슬래쉬 대각선의 꺾임 여부)(CrookedSlashFlag를 쓰기 편하도록 CrookedSlashFlag1,2로 나눔)""")
    CrookedSlashFlag2 = PropertyDescriptor("CrookedSlashFlag2", r"""High Byte CrookedSlashFlag (역슬래쉬 대각선의 꺾임 여부)(CrookedSlashFlag를 쓰기 편하도록 CrookedSlashFlag1,2로 나눔)""")


class Caption(ParameterSet):
    """Caption ParameterSet."""
    Side = PropertyDescriptor("Side", r"""방향. 0 \= 왼쪽, 1 \= 오른쪽, 2 \= 위, 3 \= 아래""")
    Width = PropertyDescriptor("Width", r"""캡션 폭 (가로 방향일 때만 사용됨. 단위 HWPUNIT)""")
    Gap = PropertyDescriptor("Gap", r"""캡션과 틀 사이 간격(HWPUNIT)""")
    CapFullSize = PropertyDescriptor("CapFullSize", r"""캡션 폭에 여백을 포함할지 여부 (세로 방향일 때만 사용됨)""")

# Forward declarations for commonly referenced parameter sets
# These are minimal implementations - full implementations would be added as needed


class CharShape(ParameterSet):
    """CharShape ParameterSet."""
    FaceNameHangul = PropertyDescriptor("FaceNameHangul", r"""글꼴 이름 (한글)""")
    FaceNameLatin = PropertyDescriptor("FaceNameLatin", r"""글꼴 이름 (영문)""")
    FaceNameHanja = PropertyDescriptor("FaceNameHanja", r"""글꼴 이름 (한자)""")
    FaceNameJapanese = PropertyDescriptor("FaceNameJapanese", r"""글꼴 이름 (일본어)""")
    FaceNameOther = PropertyDescriptor("FaceNameOther", r"""글꼴 이름 (기타)""")
    FaceNameSymbol = PropertyDescriptor("FaceNameSymbol", r"""글꼴 이름 (심벌)""")
    FaceNameUser = PropertyDescriptor("FaceNameUser", r"""글꼴 이름 (사용자)""")
    FontTypeHangul = PropertyDescriptor("FontTypeHangul", r"""폰트 종류 (한글) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeLatin = PropertyDescriptor("FontTypeLatin", r"""폰트 종류 (영문) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeHanja = PropertyDescriptor("FontTypeHanja", r"""폰트 종류 (한자) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeJapanese = PropertyDescriptor("FontTypeJapanese", r"""폰트 종류 (일본어) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeOther = PropertyDescriptor("FontTypeOther", r"""폰트 종류 (기타) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeSymbol = PropertyDescriptor("FontTypeSymbol", r"""폰트 종류 (심벌) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontTypeUser = PropertyDescriptor("FontTypeUser", r"""폰트 종류 (사용자) : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    SizeHangul = PropertyDescriptor("SizeHangul", r"""각 언어별 크기 비율. (한글) 10% \- 250%""")
    SizeLatin = PropertyDescriptor("SizeLatin", r"""각 언어별 크기 비율. (영문) 10% \- 250%""")
    SizeHanja = PropertyDescriptor("SizeHanja", r"""각 언어별 크기 비율. (한자) 10% \- 250%""")
    SizeJapanese = PropertyDescriptor("SizeJapanese", r"""각 언어별 크기 비율. (일본어) 10% \- 250%""")
    SizeOther = PropertyDescriptor("SizeOther", r"""각 언어별 크기 비율. (기타) 10% \- 250%""")
    SizeSymbol = PropertyDescriptor("SizeSymbol", r"""각 언어별 크기 비율. (심벌) 10% \- 250%""")
    SizeUser = PropertyDescriptor("SizeUser", r"""각 언어별 크기 비율. (사용자) 10% \- 250%""")
    RatioHangul = PropertyDescriptor("RatioHangul", r"""각 언어별 장평 비율. (한글) 50% \- 200%""")
    RatioLatin = PropertyDescriptor("RatioLatin", r"""각 언어별 장평 비율. (영문) 50% \- 200%""")
    RatioHanja = PropertyDescriptor("RatioHanja", r"""각 언어별 장평 비율. (한자) 50% \- 200%""")
    RatioJapanese = PropertyDescriptor("RatioJapanese", r"""각 언어별 장평 비율. (일본어) 50% \- 200%""")
    RatioOther = PropertyDescriptor("RatioOther", r"""각 언어별 장평 비율. (기타) 50% \- 200%""")
    RatioSymbol = PropertyDescriptor("RatioSymbol", r"""각 언어별 장평 비율. (심벌) 50% \- 200%""")
    RatioUser = PropertyDescriptor("RatioUser", r"""각 언어별 장평 비율. (사용자) 50% \- 200%""")
    SpacingHangul = PropertyDescriptor("SpacingHangul", r"""각 언어별 자간. (한글) \-50% \- 50%""")
    SpacingLatin = PropertyDescriptor("SpacingLatin", r"""각 언어별 자간. (영문) \-50% \- 50%""")
    SpacingHanja = PropertyDescriptor("SpacingHanja", r"""각 언어별 자간. (한자) \-50% \- 50%""")
    SpacingJapanese = PropertyDescriptor("SpacingJapanese", r"""각 언어별 자간. (일본어) \-50% \- 50%""")
    SpacingOther = PropertyDescriptor("SpacingOther", r"""각 언어별 자간. (기타) \-50% \- 50%""")
    SpacingSymbol = PropertyDescriptor("SpacingSymbol", r"""각 언어별 자간. (심벌) \-50% \- 50%""")
    SpacingUser = PropertyDescriptor("SpacingUser", r"""각 언어별 자간. (사용자) \-50% \- 50%""")
    OffsetHangul = PropertyDescriptor("OffsetHangul", r"""각 언어별 오프셋. (한글) \-100% \- 100%""")
    OffsetLatin = PropertyDescriptor("OffsetLatin", r"""각 언어별 오프셋. (영문) \-100% \- 100%""")
    OffsetHanja = PropertyDescriptor("OffsetHanja", r"""각 언어별 오프셋. (한자) \-100% \- 100%""")
    OffsetJapanese = PropertyDescriptor("OffsetJapanese", r"""각 언어별 오프셋. (일본어) \-100% \- 100%""")
    OffsetOther = PropertyDescriptor("OffsetOther", r"""각 언어별 오프셋. (기타) \-100% \- 100%""")
    OffsetSymbol = PropertyDescriptor("OffsetSymbol", r"""각 언어별 오프셋. (심벌) \-100% \- 100%""")
    OffsetUser = PropertyDescriptor("OffsetUser", r"""각 언어별 오프셋. (사용자) \-100% \- 100%""")
    Bold = PropertyDescriptor("Bold", r"""Bold : 0 \= off, 1 \= on""")
    Italic = PropertyDescriptor("Italic", r"""Italic : 0 \= off, 1 \= on""")
    SmallCaps = PropertyDescriptor("SmallCaps", r"""Small Caps : 0 \= off, 1 \= on""")
    Emboss = PropertyDescriptor("Emboss", r"""Emboss : 0 \= off, 1 \= on""")
    Engrave = PropertyDescriptor("Engrave", r"""Engrave : 0 \= off, 1 \= on""")
    SuperScript = PropertyDescriptor("SuperScript", r"""Superscript : 0 \= off, 1 \= on""")
    SubScript = PropertyDescriptor("SubScript", r"""Subscript : 0 \= off, 1 \= on""")
    UnderlineType = PropertyDescriptor("UnderlineType", r"""밑줄 종류 : 0 \= none, 1 \= bottom, 2 \= center, 3 \= top""")
    OutlineType = PropertyDescriptor("OutlineType", r"""외곽선 종류 : 0 \= none, 1 \= solid, 2 \= dot, 3 \= thick, 4 \= dash, 5 \= dashdot, 6 \= dashdotdot""")
    ShadowType = PropertyDescriptor("ShadowType", r"""그림자 종류 : 0 \= none, 1 \= drop, 2 \= continuous""")
    TextColor = ColorProperty("TextColor", r"""글자색. (COLORREF)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    ShadeColor = ColorProperty("ShadeColor", r"""음영색. (COLORREF)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    UnderlineShape = PropertyDescriptor("UnderlineShape", r"""밑줄 모양 : 선 종류""")
    UnderlineColor = ColorProperty("UnderlineColor", r"""밑줄 색 (COLORREF)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    ShadowOffsetX = PropertyDescriptor("ShadowOffsetX", r"""그림자 간격 (X 방향) \-100% \- 100%""")
    ShadowOffsetY = PropertyDescriptor("ShadowOffsetY", r"""그림자 간격 (Y 방향) \-100% \- 100%""")
    ShadowColor = ColorProperty("ShadowColor", r"""그림자 색 (COLORREF)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    StrikeOutType = PropertyDescriptor("StrikeOutType", r"""취소선 종류 : 0 \= none, 1 \= red single, 2 \= red double, 3 \= text single, 4 \= text double""")
    DiacSymMark = PropertyDescriptor("DiacSymMark", r"""강조점 종류 : 0 \= none, 1 \= 검정 동그라미, 2 \= 속 빈 동그라미""")
    UseFontSpace = PropertyDescriptor("UseFontSpace", r"""글꼴에 어울리는 빈칸 : 0 \= off, 1 \= on""")
    UseKerning = PropertyDescriptor("UseKerning", r"""커닝 : 0 \= off, 1 \= on""")
    Height = PropertyDescriptor("Height", r"""글자 크기 (HWPUNIT)""")
    BorderFill = PropertyDescriptor("BorderFill", r"""테두리/배경 (한글2007에 새로 추가)""")


class ParaShape(ParameterSet):
    """ParaShape ParameterSet."""
    LeftMargin = PropertyDescriptor("LeftMargin", r"""왼쪽 여백 (URC)""")
    RightMargin = PropertyDescriptor("RightMargin", r"""오른쪽 여백 (URC)""")
    Indentation = PropertyDescriptor("Indentation", r"""들여쓰기/내어 쓰기 (URC)""")
    PrevSpacing = PropertyDescriptor("PrevSpacing", r"""문단 간격 위 (URC)""")
    NextSpacing = PropertyDescriptor("NextSpacing", r"""문단 간격 아래 (URC)""")
    LineSpacingType = PropertyDescriptor("LineSpacingType", r"""줄 간격 종류 (HWPUNIT)0 \= 글자에 따라1 \= 고정 값2 \= 여백만 지정""")
    LineSpacing = PropertyDescriptor("LineSpacing", r"""줄 간격 값. 줄 간격 종류(LineSpacingType)에 따라 :\- 글자에 따라일 경우(0 \- 500%)\- "고정 값"일 경우(URC)\- "여백만 지정"일 경우(URC)""")
    AlignType = PropertyDescriptor("AlignType", r"""정렬 방식0 \= 양쪽 정렬1 \= 왼쪽 정렬2 \= 오른쪽 정렬3 \= 가운데 정렬4 \= 배분 정렬5 \= 나눔 정렬 (공백에만 배분)""")
    BreakLatinWord = PropertyDescriptor("BreakLatinWord", r"""줄 나눔 단위 (라틴 문자)0 \= 단어1 \= 하이픈2 \= 글자""")
    BreakNonLatinWord = PropertyDescriptor("BreakNonLatinWord", r"""단위 (비 라틴 문자) TRUE \= 글자, FALSE \= 어절""")
    SnapToGrid = PropertyDescriptor("SnapToGrid", r"""편집 용지의 줄 격자 사용 (on / off)""")
    Condense = PropertyDescriptor("Condense", r"""공백 최소값 (0 \- 75%)""")
    WidowOrphan = PropertyDescriptor("WidowOrphan", r"""외톨이줄 보호 (on / off)""")
    KeepWithNext = PropertyDescriptor("KeepWithNext", r"""다음 문단과 함께 (on / off)""")
    KeepLinesTogether = PropertyDescriptor("KeepLinesTogether", r"""문단 보호 (on / off)""")
    PagebreakBefore = PropertyDescriptor("PagebreakBefore", r"""문단 앞에서 항상 쪽 나눔 (on / off)""")
    TextAlignment = PropertyDescriptor("TextAlignment", r"""세로 정렬0 \= 글꼴기준1 \= 위2 \= 가운데3 \= 아래""")
    FontLineHeight = PropertyDescriptor("FontLineHeight", r"""글꼴에 어울리는 줄 높이 (on / off)""")
    HeadingType = PropertyDescriptor("HeadingType", r"""문단 머리 모양0 \= 없음1 \= 개요2 \= 번호3 \= 불릿""")
    Level = PropertyDescriptor("Level", r"""단계 (0 \- 6\)""")
    BorderConnect = PropertyDescriptor("BorderConnect", r"""문단 테두리/배경 \- 테두리 연결 (on / off)""")
    BorderText = PropertyDescriptor("BorderText", r"""문단 테두리/배경 \- 여백 무시 (0 \= 단, 1 \= 텍스트)""")
    BorderOffsetLeft = PropertyDescriptor("BorderOffsetLeft", r"""문단 테두리/배경 \- 4방향 간격 (HWPUNIT) : 왼쪽""")
    BorderOffsetRight = PropertyDescriptor("BorderOffsetRight", r"""문단 테두리/배경 \- 4방향 간격 (HWPUNIT) : 오른쪽""")
    BorderOffsetTop = PropertyDescriptor("BorderOffsetTop", r"""문단 테두리/배경 \- 4방향 간격 (HWPUNIT) : 위""")
    BorderOffsetBottom = PropertyDescriptor("BorderOffsetBottom", r"""문단 테두리/배경 \- 4방향 간격 (HWPUNIT) : 아래""")
    TailType = PropertyDescriptor("TailType", r"""문단 꼬리 모양 (마지막 꼬리 줄 적용) on/off""")
    LineWrap = PropertyDescriptor("LineWrap", r"""글꼴에 어울리는 줄 높이 (on/off)""")
    TabDef = PropertyDescriptor("TabDef", r"""탭 정의""")
    Numbering = PropertyDescriptor("Numbering", r"""문단 번호문단 머리 모양(HeadingType)이 ‘개요’, ‘번호’ 일 때 사용""")
    Bullet = PropertyDescriptor("Bullet", r"""불릿 모양문단 머리 모양(HeadingType)이 ‘불릿’(글머리표) 일 때 사용""")
    BorderFill = PropertyDescriptor("BorderFill", r"""테두리/배경""")


class Cell(ParameterSet):
    """Cell ParameterSet."""
    HasMargin = PropertyDescriptor("HasMargin", r"""테이블의 기본 셀 여백 대신 자체 셀 여백을 적용할지 여부. (on / off)""")
    Protected = PropertyDescriptor("Protected", r"""사용자 편집을 막을지 여부 : 0 \= off, 1 \= on""")
    Header = PropertyDescriptor("Header", r"""제목 셀인지 여부 : 0 \= off, 1 \= on""")
    Width = PropertyDescriptor("Width", r"""셀의 폭 (HWPUNIT)""")
    Height = PropertyDescriptor("Height", r"""셀의 높이 (HWPUNIT)""")
    Editable = PropertyDescriptor("Editable", r"""양식모드에서 편집 가능 여부 : 0 \= off, 1 \= on""")
    Dirty = PropertyDescriptor("Dirty", r"""초기화 상태인지 수정된 상태인지 여부 : 0 \= off, 1 \= on(한글2007에 새로 추가)""")
    CellCtrlData = PropertyDescriptor("CellCtrlData", r"""셀 데이터 (한글2007에 새로 추가)""")


class CtrlData(ParameterSet):
    """
    ### CtrlData

    22) CtrlData : 제어 데이터

    | Item ID | Type      | SubType | Description          |
    |---------|-----------|---------|----------------------|
    | Name    | PIT_BSTR  |         | 제어 데이터의 이름.  |
    """
    name = ParameterSet._str_prop("Name", "제어 데이터의 이름.")



class Password(ParameterSet):
    """Password ParameterSet."""
    DialogResult = PropertyDescriptor("DialogResult", r"""암호해제 버튼 클릭 여부. 1=암호해제""")
    String = PropertyDescriptor("String", r"""암호 문자열""")
    FullRange = PropertyDescriptor("FullRange", r"""TRUE=유니코드/모든 문자, FALSE=영문자만""")
    Ask = PropertyDescriptor("Ask", r"""TRUE=문서 암호 확인, FALSE=문서 암호 설정""")
    Level = PropertyDescriptor("Level", r"""보안 수준: 0=낮음, 1=높음 (한글2007)""")


class Style(ParameterSet):
    """Style ParameterSet."""
    Apply = PropertyDescriptor("Apply", r"""적용할 스타일 인덱스""")

