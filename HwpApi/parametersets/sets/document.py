"""
Document / page / section ParameterSet classes (17 classes).

Covers document-level and section-level configuration:
- DocumentInfo, SummaryInfo, VersionInfo, FileInfo
- PageDef, PageBorderFill, PageHiding, PageNumCtrl, PageNumPos
- SecDef (32 properties), SectionApply, ColDef
- HeaderFooter, FootnoteShape, EndnoteShape
- GridInfo, MasterPage
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class ColDef(ParameterSet):
    """ColDef ParameterSet."""
    Type = PropertyDescriptor("Type", r"""단 종류 : 0 \= 보통 다단, 1 \= 배분 다단, 2 \= 평행 다단""")
    Count = PropertyDescriptor("Count", r"""단 개수. 1\-255까지.""")
    SameSize = PropertyDescriptor("SameSize", r"""단의 너비를 같도록 할지 여부 :0 \= 단 너비 각자 지정, 1 \= 단 너비 동일""")
    SameGap = PropertyDescriptor("SameGap", r"""단 사이 간격(HWPUNIT) SAME_SIZE가 1일 때만 사용된다.""")
    WidthGap = PropertyDescriptor("WidthGap", r"""각 단의 너비와 간격(HWPUNIT)col\*2 \= 단의 폭, col\*2 \+ 1 \= 단 사이 간격.영역 전체의 폭을 Column ratio base (\=32768\)로 보았을 때의 비율로 환산한다.SameSize가 0일 때만 사용된다.배열의 아이템의 개수는 Count\*2\-1과 같아야 한다.""")
    Layout = PropertyDescriptor("Layout", r"""단 방향 지정 :0 \= 왼쪽부터, 1 \= 오른쪽부터, 2 \= 맞쪽""")
    LineType = PropertyDescriptor("LineType", r"""선 종류.""")
    LineWidth = PropertyDescriptor("LineWidth", r"""선 굵기.""")
    LineColor = PropertyDescriptor("LineColor", r"""선 색깔. (COLORREF)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위 :0 \= 선택된 다단1 \= 선택된 문자열2 \= 현재 다단3 \= 개체 전체4 \= 선택된 셀5 \= 현재 구역6 \= 문서 전체7 \= 현재 셀8 \= 새 쪽으로9 \= 새 다단으로10 \= 모든 셀""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용범위의 분류. 아래 값의 조합이다.0x0001 \= 선택된 다단0x0002 \= 선택된 문자열0x0004 \= 현재 다단0x0008 \= 개체 전체0x0010 \= 선택된 셀0x0020 \= 현재 구역0x0040 \= 문서전체0x0080 \= 현재 셀0x0100 \= 새 쪽으로0x0200 \= 새 다단으로0x0400 \= 모든 셀""")


class DocumentInfo(ParameterSet):
    """DocumentInfo ParameterSet."""
    CurPara = PropertyDescriptor("CurPara", r"""현재 위치한 문단""")
    CurPos = PropertyDescriptor("CurPos", r"""현재 위치한 오프셋""")
    CurParaLen = PropertyDescriptor("CurParaLen", r"""현재 위치한 문단의 길이""")
    CurCtrl = PropertyDescriptor("CurCtrl", r"""현재 리스트의 종류0 \= 일반 텍스트1 \= 글상자기타 \= 컨트롤 ID""")
    CurParaCount = PropertyDescriptor("CurParaCount", r"""현재 리스트의 문단 수""")
    RootPara = PropertyDescriptor("RootPara", r"""루트 리스트의 현재 문단""")
    RootPos = PropertyDescriptor("RootPos", r"""루트 리스트의 현재 오프셋""")
    RootParaCout = PropertyDescriptor("RootParaCout", r"""루트 리스트의 문단 수""")
    DetailInfo = PropertyDescriptor("DetailInfo", r"""자세한 정보를 구할지 여부Detail\~ 로 시작하는 아이템의 정보를 얻기 위해서는 이 값을 1로 넣어준 후에 액션을 실행해준다.""")
    DetailCharCount = PropertyDescriptor("DetailCharCount", r"""문서에 포함된 글자 수""")
    DetailWordCount = PropertyDescriptor("DetailWordCount", r"""문서에 포함된 어절 수""")
    DetailLineCount = PropertyDescriptor("DetailLineCount", r"""문서에 포함된 줄 수""")
    DetailPageCount = PropertyDescriptor("DetailPageCount", r"""문서에 포함된 쪽 수""")
    DetailCurPage = PropertyDescriptor("DetailCurPage", r"""현재 쪽 번호""")
    DetailCurPrtPage = PropertyDescriptor("DetailCurPrtPage", r"""현재 쪽 번호 (인쇄 번호)""")
    SectionInfo = PropertyDescriptor("SectionInfo", r"""구역의 속성까지 구할지 여부SecDef 아이템은 이 값을 1로 넣어준 후 액션을 실행한 후에 얻을 수 있다.""")
    SecDef = PropertyDescriptor("SecDef", r"""구역의 속성 (한글2007에 새로 추가)""")


class FileInfo(ParameterSet):
    """FileInfo ParameterSet."""
    Format = PropertyDescriptor("Format", r"""파일의 형식.HWP : 한글 파일UNKNOWN : 알 수 없음.""")
    VersionStr = PropertyDescriptor("VersionStr", r"""파일의 버전 문자열ex)5\.0\.0\.3""")
    VersionNum = PropertyDescriptor("VersionNum", r"""파일의 버전ex) 0x05000003""")
    Encrypted = PropertyDescriptor("Encrypted", r"""암호 여부 (현재는 파일 버전 3\.0\.0\.0 이후 문서\-한글97, 한글 워디안, 한글 2002\-에 대해서만 판단한다.)\-1 : 판단 할 수 없음0 : 암호가 걸려 있지 않음양수: 암호가 걸려 있음.""")


class FootnoteShape(ParameterSet):
    """FootnoteShape ParameterSet."""
    NumberFormat = PropertyDescriptor("NumberFormat", r"""번호모양""")
    UserChar = PropertyDescriptor("UserChar", r"""사용자 기호 (WCHAR)""")
    PrefixChar = PropertyDescriptor("PrefixChar", r"""앞 장식 문자 (WCHAR)""")
    Suffix = PropertyDescriptor("Suffix", r"""뒤 장식 문자 (WCHAR)""")
    PlaceAt = PropertyDescriptor("PlaceAt", r"""위치\- 각주용 옵션 (한 페이지 내에서 각주를 다단에 어떻게 위치시킬지)0 \= 각 단마다 따로 배열1 \= 통단으로 배열2 \= 가장 오른쪽 단에 배열 \- 미주용 옵션 (문서 내에서 미주를 어디에 위치시킬지)0 \= 문서의 마지막1 \= 구역의 마지막""")
    Restart = PropertyDescriptor("Restart", r"""번호 매기기0 \= 앞 구역에 이어서1 \= 현재 구역부터 새로 시작2 \= 쪽마다 새로 시작 (각주 전용)""")
    NewNumber = PropertyDescriptor("NewNumber", r"""시작 번호 (1 .. n)번호 매기기 값이 ‘쪽마다 새로 시작’ 일 때만 사용된다.""")
    LineLength = PropertyDescriptor("LineLength", r"""구분선 길이 (HWPUNIT)""")
    LineType = PropertyDescriptor("LineType", r"""선 종류""")
    LineWidth = PropertyDescriptor("LineWidth", r"""선 굵기""")
    SpaceAboveLine = PropertyDescriptor("SpaceAboveLine", r"""구분선 위 여백 (HWPUNIT)""")
    SpaceBelowLine = PropertyDescriptor("SpaceBelowLine", r"""구분선 아래 여백 (HWPUNIT)""")
    SpaceBetweenNotes = PropertyDescriptor("SpaceBetweenNotes", r"""주석 사이 여백 (HWPUNIT)""")
    SuperScript = PropertyDescriptor("SuperScript", r"""각주 내용 중 번호 코드의 모양을 위첨자 형식으로 할지""")
    BeneathText = PropertyDescriptor("BeneathText", r"""텍스트에 이어 바로 출력할지 여부 (on / off)""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위0 \= 선택된 구역1 \= 선택된 문자열2 \= 현재 구역3 \= 문서전체4 \= 새 구역 : 현재 위치부터 새로""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용범위의 분류 (대화상자를 호출할 경우 사용)0x01 \= 선택된 구역0x02 \= 선택된 문자열0x04 \= 현재 구역0x08 \= 문서전체0x10 \= 새 구역 : 현재 위치부터 새로""")


class EndnoteShape(ParameterSet):
    """EndnoteShape ParameterSet - 미주 모양 (FootnoteShape와 동일 구조)."""
    NumberFormat = PropertyDescriptor("NumberFormat", r"""번호모양""")
    UserChar = PropertyDescriptor("UserChar", r"""사용자 기호 (WCHAR)""")
    PrefixChar = PropertyDescriptor("PrefixChar", r"""앞 장식 문자 (WCHAR)""")
    Suffix = PropertyDescriptor("Suffix", r"""뒤 장식 문자 (WCHAR)""")
    PlaceAt = PropertyDescriptor("PlaceAt", r"""위치: 0=문서의 마지막, 1=구역의 마지막""")
    Restart = PropertyDescriptor("Restart", r"""번호 매기기: 0=앞 구역에 이어서, 1=현재 구역부터 새로 시작""")
    NewNumber = PropertyDescriptor("NewNumber", r"""시작 번호 (1..n)""")
    LineLength = PropertyDescriptor("LineLength", r"""구분선 길이 (HWPUNIT)""")
    LineType = PropertyDescriptor("LineType", r"""선 종류""")
    LineWidth = PropertyDescriptor("LineWidth", r"""선 굵기""")
    SpaceAboveLine = PropertyDescriptor("SpaceAboveLine", r"""구분선 위 여백 (HWPUNIT)""")
    SpaceBelowLine = PropertyDescriptor("SpaceBelowLine", r"""구분선 아래 여백 (HWPUNIT)""")
    SpaceBetweenNotes = PropertyDescriptor("SpaceBetweenNotes", r"""주석 사이 여백 (HWPUNIT)""")
    SuperScript = PropertyDescriptor("SuperScript", r"""번호 코드 위첨자 형식 여부""")
    BeneathText = PropertyDescriptor("BeneathText", r"""텍스트에 이어 바로 출력할지 여부""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위: 0=선택된 구역, 1=선택된 문자열, 2=현재 구역, 3=문서전체, 4=새 구역""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용범위 분류 (비트마스크): 0x01=선택된 구역, 0x02=선택된 문자열, 0x04=현재 구역, 0x08=문서전체, 0x10=새 구역""")


class GridInfo(ParameterSet):
    """GridInfo ParameterSet."""
    Method = PropertyDescriptor("Method", r"""격자 방식0 \= 격자와 상관없이1 \= 격자 자석효과2 \= 격자에만 붙음""")
    Align = PropertyDescriptor("Align", r"""격자 기준(쪽 \= 0 / 종이 \= 1\)""")
    HorzAlign = PropertyDescriptor("HorzAlign", r"""격자 기준 가로 offset (단위 HWPUNIT)""")
    VertAlign = PropertyDescriptor("VertAlign", r"""격자 기준 세로 offset (단위 HWPUNIT)""")
    Type = PropertyDescriptor("Type", r"""격자 모양0 \= 점 격자1 \= 선 격자""")
    HorzSpan = PropertyDescriptor("HorzSpan", r"""가로 간격 (단위 HWPUNIT)""")
    VertSpan = PropertyDescriptor("VertSpan", r"""세로 간격 (단위 HWPUNIT)""")
    HorzRange = PropertyDescriptor("HorzRange", r"""가로 자석 범위 (단위 HWPUNIT)""")
    VertRange = PropertyDescriptor("VertRange", r"""세로 자석 범위 (단위 HWPUNIT)""")
    Show = PropertyDescriptor("Show", r"""격자 보이기 ( on / off )""")
    ZOrder = PropertyDescriptor("ZOrder", r"""격자 위치(글 위/글 아래) (ZOrder)0 \= 글 아래, 1 \= 글 위""")
    ViewLine = PropertyDescriptor("ViewLine", r"""선격자 보이기 종류 (한글2007에 새로 추가)0 \= 모두1 \= 수평격자만2 \= 수직격자만""")


class HeaderFooter(ParameterSet):
    """HeaderFooter ParameterSet."""
    DialogOption = PropertyDescriptor("DialogOption", r"""머리말/꼬리말 대화상자를 실행할 때 "편집"버튼을 보일 것인지 말 것인지 지정한다. 1=보이기, 그외=안보이기""")
    DialogResult = PropertyDescriptor("DialogResult", r"""대화상자 버튼 결과. 1=만들기(삽입), 2=편집, 그외=취소""")
    CtrlType = PropertyDescriptor("CtrlType", r"""머리말/꼬리말 여부 : 0=머리말, 1=꼬리말""")
    Type = PropertyDescriptor("Type", r"""머리말/꼬리말 위치 : 0=양쪽, 1=짝수쪽, 2=홀수쪽""")
    HeaderFooterCtrlType = PropertyDescriptor("HeaderFooterCtrlType", r"""머리말/꼬리말 종류 : 0=머리말, 1=꼬리말 (한글2007)""")
    HeaderFooterStyle = PropertyDescriptor("HeaderFooterStyle", r"""머리말/꼬리말 마당 스타일 (한글2007)""")


class MasterPage(ParameterSet):
    """MasterPage ParameterSet."""
    Type = PropertyDescriptor("Type", r"""바탕쪽 종류0 \= 양쪽1 \= 짝수쪽2 \= 홀수쪽""")
    Duplicate = PropertyDescriptor("Duplicate", r"""기존 바탕쪽과 겹침 (On/Off) (한글2007에 새로 추가)""")
    Front = PropertyDescriptor("Front", r"""바탕쪽과 앞으로 보내기 (On/Off) (한글2007에 새로 추가)""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용대상 (한글2007에 새로 추가)0 \= 현재구역1 \= 문서 전체""")


class PageBorderFill(ParameterSet):
    """PageBorderFill ParameterSet."""
    TextBorder = PropertyDescriptor("TextBorder", r"""TRUE \= 본문 기준, FALSE \= 종이 기준""")
    HeaderInside = PropertyDescriptor("HeaderInside", r"""머리말 포함 (on / off)""")
    FooterInside = PropertyDescriptor("FooterInside", r"""꼬리말 포함 (on / off)""")
    FillArea = PropertyDescriptor("FillArea", r"""채울 영역0 \= 종이1 \= 쪽2 \= 테두리""")
    OffsetLeft = PropertyDescriptor("OffsetLeft", r"""4방향 간격 (HWPUNIT) : 왼쪽""")
    OffsetRight = PropertyDescriptor("OffsetRight", r"""4방향 간격 (HWPUNIT) : 오른쪽""")
    OffsetTop = PropertyDescriptor("OffsetTop", r"""4방향 간격 (HWPUNIT) : 위""")
    OffsetBottom = PropertyDescriptor("OffsetBottom", r"""4방향 간격 (HWPUNIT) : 아래""")


class PageDef(ParameterSet):
    """PageDef ParameterSet."""
    PaperWidth = PropertyDescriptor("PaperWidth", r"""용지 가로 크기 (HWPUNIT)""")
    PaperHeight = PropertyDescriptor("PaperHeight", r"""용지 세로 크기 (HWPUNIT)""")
    Landscape = PropertyDescriptor("Landscape", r"""용지 방향. 0 \= 좁게, 1 \= 넓게""")
    LeftMargin = PropertyDescriptor("LeftMargin", r"""왼쪽 여백 (HWPUNIT)""")
    RightMargin = PropertyDescriptor("RightMargin", r"""오른쪽 여백 (HWPUNIT)""")
    TopMargin = PropertyDescriptor("TopMargin", r"""위 여백 (HWPUNIT)""")
    BottomMargin = PropertyDescriptor("BottomMargin", r"""아래 여백 (HWPUNIT)""")
    HeaderLen = PropertyDescriptor("HeaderLen", r"""머리말 길이 (HWPUNIT)""")
    FooterLen = PropertyDescriptor("FooterLen", r"""꼬리말 길이 (HWPUNIT)""")
    GutterLen = PropertyDescriptor("GutterLen", r"""제본 여백 (HWPUNIT)""")
    GutterType = PropertyDescriptor("GutterType", r"""편집 방법. 0 \= 한쪽 편집, 1 \= 맞쪽 편집, 2 \= 위로 넘기기""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위0 \= 선택된 구역1 \= 선택된 문자열2 \= 현재 구역3 \= 문서전체4 \= 새 구역 : 현재 위치부터 새로""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용범위의 분류 (대화상자를 호출할 경우 사용)0x01 \= 선택된 구역0x02 \= 선택된 문자열0x04 \= 현재 구역0x08 \= 문서전체0x10 \= 새 구역 : 현재 위치부터 새로""")


class PageHiding(ParameterSet):
    """PageHiding ParameterSet."""
    Fields = PropertyDescriptor("Fields", r"""감출 대상 비트 필드.0x01 \= 머리말0x02 \= 꼬리말0x04 \= 바탕쪽0x08 \= 테두리0x10 \= 배경0x20 \= 쪽번호 위치""")


class PageNumCtrl(ParameterSet):
    """PageNumCtrl ParameterSet."""
    PageStartsOn = PropertyDescriptor("PageStartsOn", r"""페이지 번호 적용 옵션. 0 \= 양쪽1 \= 짝수쪽2 \= 홀수쪽""")


class PageNumPos(ParameterSet):
    """PageNumPos ParameterSet."""
    NumberFormat = PropertyDescriptor("NumberFormat", r"""번호 모양 :0 \= 1, 2, 31 \= ①, ②, ③2 \= Ⅰ, Ⅱ, Ⅲ3 \= ⅰ, ⅱ, ⅲ4 \= A, B, C8 \= 가, 나, 다13 \= 一, 二, 三15 \= 갑, 을, 병16 \= 甲, 乙, 丙※ 중간에 빈 번호에도 문자포맷이 존재하나 이곳에서 사용하지 않아 생략함""")
    UserChar = PropertyDescriptor("UserChar", r"""사용자 기호(WCHAR). 한글2007에선 더 이상 사용하지 않는다.""")
    PrefixChar = PropertyDescriptor("PrefixChar", r"""앞 장식 문자(WCHAR). 한글2007에선 더 이상 사용하지 않는다.""")
    SuffixChar = PropertyDescriptor("SuffixChar", r"""뒤 장식 문자(WCHAR). 한글2007에선 더 이상 사용하지 않는다.""")
    SideChar = PropertyDescriptor("SideChar", r"""양쪽 옆 장식 문자(WCHAR). L\-만 사용할 수 있다.""")
    DrawPos = PropertyDescriptor("DrawPos", r"""번호 위치0 \= 쪽 번호 없음1 \= 왼쪽 위,2 \= 가운데 위,3 \= 오른쪽 위,4 \= 왼쪽 아래,5 \= 가운데 아래,6 \= 오른쪽 아래,7 \= 바깥쪽 위,8 \= 바깥쪽 아래,9 \= 안쪽 위,10 \= 안쪽 아래""")


class SecDef(ParameterSet):
    """SecDef ParameterSet."""
    TextDirection = PropertyDescriptor("TextDirection", r"""글자 방향""")
    StartsOn = PropertyDescriptor("StartsOn", r"""구역 나눔으로 새 페이지가 생길 때의 페이지 번호 적용 옵션 0 \= 이어서, 1 \= 홀수, 2 \= 짝수, 3 \= 임의 값""")
    LineGrid = PropertyDescriptor("LineGrid", r"""세로로 줄맞춤을 할지 여부. 0 \= off, 1 \- n \= 간격을 HWPUNIT 단위로 지정""")
    CharGrid = PropertyDescriptor("CharGrid", r"""가로로 줄맞춤을 할지 여부. 0 \= off, 1 \- n \= 간격을 HWPUNIT 단위로 지정""")
    PageDef = PropertyDescriptor("PageDef", r"""용지 설정 정보""")
    HideEmptyLine = PropertyDescriptor("HideEmptyLine", r"""빈 줄 감춤 on / off""")
    SpaceBetweenColumns = PropertyDescriptor("SpaceBetweenColumns", r"""동일한 페이지에서 서로 다른 단 사이의 간격""")
    TabStop = PropertyDescriptor("TabStop", r"""기본 탭 간격""")
    FootnoteShape = PropertyDescriptor("FootnoteShape", r"""각주 모양""")
    EndnoteShape = PropertyDescriptor("EndnoteShape", r"""미주 모양""")
    HideHeader = PropertyDescriptor("HideHeader", r"""구역의 첫 쪽에만 머리말 감추기 옵션 on / off""")
    HideFooter = PropertyDescriptor("HideFooter", r"""구역의 첫 쪽에만 꼬리말 감추기 옵션 on / off""")
    HideMasterPage = PropertyDescriptor("HideMasterPage", r"""구역의 첫 쪽에만 바탕쪽 감추기 옵션 on / off""")
    HideBorder = PropertyDescriptor("HideBorder", r"""구역의 첫 쪽에만 테두리 감추기 옵션 on / off""")
    HideFill = PropertyDescriptor("HideFill", r"""구역의 첫 쪽에만 배경 감추기 옵션 on / off""")
    HidePageNumPos = PropertyDescriptor("HidePageNumPos", r"""구역의 첫 쪽에만 쪽번호 감추기 옵션 on / off""")
    FirstBorder = PropertyDescriptor("FirstBorder", r"""구역의 첫 쪽에만 테두리 표시 옵션 on / off""")
    FirstFill = PropertyDescriptor("FirstFill", r"""구역의 첫 쪽에만 배경 표시 옵션 on / off""")
    OutlineShape = PropertyDescriptor("OutlineShape", r"""개요 번호""")
    PageBorderFillBoth = PropertyDescriptor("PageBorderFillBoth", r"""쪽 테두리/배경 (양쪽)""")
    PageBorderFillEven = PropertyDescriptor("PageBorderFillEven", r"""쪽 테두리/배경 (짝수)""")
    PageBorderFillOdd = PropertyDescriptor("PageBorderFillOdd", r"""쪽 테두리/배경 (홀수)""")
    PageNumber = PropertyDescriptor("PageNumber", r"""쪽 시작 번호 0 \= 앞 구역에 이어, n \= 새 번호로 시작""")
    FigureNumber = PropertyDescriptor("FigureNumber", r"""그림 시작 번호 0 \= 앞 구역에 이어, n \= 새 번호로 시작""")
    TableNumber = PropertyDescriptor("TableNumber", r"""표 시작 번호 0 \= 앞 구역에 이어, n \= 새 번호로 시작""")
    EquationNumber = PropertyDescriptor("EquationNumber", r"""수식 시작 번호 0 \= 앞 구역에 이어, n \= 새 번호로 시작""")
    WongojiFormat = PropertyDescriptor("WongojiFormat", r"""원고지 방식의 포맷팅. CHAR_GRID가 지정되어야 함.""")
    MemoShape = PropertyDescriptor("MemoShape", r"""메모 모양 (한글2007에 새로 추가)""")
    TextVerticalWidthHead = PropertyDescriptor("TextVerticalWidthHead", r"""머리말/꼬리말 세로쓰기 여부 (한글2007에 새로 추가)""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위0 \= 선택된 구역1 \= 선택된 문자열2 \= 현재 구역3 \= 문서전체4 \= 새 구역 : 현재 위치부터 새로""")
    ApplyClass = PropertyDescriptor("ApplyClass", r"""적용범위의 분류 (대화상자를 호출할 경우 사용)0x01 \= 선택된 구역0x02 \= 선택된 문자열0x04 \= 현재 구역0x08 \= 문서전체0x10 \= 새 구역 : 현재 위치부터 새로""")
    ApplyToPageBorderFill = PropertyDescriptor("ApplyToPageBorderFill", r"""채울 영역 분류 (PageBorder 액션에서 사용)0 \= 종이, 1 \= 쪽, 2 \= 테두리 (한글2007에 새로 추가)""")


class SectionApply(ParameterSet):
    """SectionApply ParameterSet."""
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용범위 분류(ApplyClass)에서 하나의 값""")
    String = PropertyDescriptor("String", r"""ApplyTo를 문자열로 변환한 값의 배열""")
    Index = PropertyDescriptor("Index", r"""ApplyTo를 변환한 ComboBox의 Index""")
    ConvAplly2Index = PropertyDescriptor("ConvAplly2Index", r"""ApplyTo값을 ComboBox의 Index로 변환할지 여부FALSE이면 IndexToApply로 변환이 이루어진다. (반대변환)""")
    Category = PropertyDescriptor("Category", r"""적용범위 분류(ApplyClass)를 사용자가 직접 설정할 때 사용.아이템이 없으면 한글이 현재 상태에 맞춰 적용범위 분류(ApplyClass)를 설정한다. (일반적으로 설정하지 않고 사용)""")


class SummaryInfo(ParameterSet):
    """SummaryInfo ParameterSet."""
    Title = PropertyDescriptor("Title", r"""Parameter property""")
    Subject = PropertyDescriptor("Subject", r"""Parameter property""")
    Author = PropertyDescriptor("Author", r"""지은이""")
    Date = PropertyDescriptor("Date", r"""Parameter property""")
    KeyWords = PropertyDescriptor("KeyWords", r"""키워드""")
    Comments = PropertyDescriptor("Comments", r"""Parameter property""")
    CreationTimeLow = PropertyDescriptor("CreationTimeLow", r"""작성한 날짜 (low)""")
    CreationTimeHigh = PropertyDescriptor("CreationTimeHigh", r"""작성한 날짜 (high)""")
    ModifiedTimeLow = PropertyDescriptor("ModifiedTimeLow", r"""마지막 수정한 날짜 (low)""")
    ModifiedTimeHigh = PropertyDescriptor("ModifiedTimeHigh", r"""마지막 수정한 날짜 (high)""")
    PrintedTimeLow = PropertyDescriptor("PrintedTimeLow", r"""마지막 인쇄한 날짜 (low)""")
    PrintedTimeHigh = PropertyDescriptor("PrintedTimeHigh", r"""마지막 인쇄한 날짜 (high)""")
    LastSavedBy = PropertyDescriptor("LastSavedBy", r"""마지막 저장한 사람""")
    Characters = PropertyDescriptor("Characters", r"""문서분량 (글자)""")
    Words = PropertyDescriptor("Words", r"""문서분량 (낱말)""")
    Lines = PropertyDescriptor("Lines", r"""문서분량 (줄)""")
    Paragraphs = PropertyDescriptor("Paragraphs", r"""문서분량 (문단)""")
    Pages = PropertyDescriptor("Pages", r"""문서분량 (쪽)""")
    CopyPapers = PropertyDescriptor("CopyPapers", r"""문서분량 (원고지)""")
    Etcetera = PropertyDescriptor("Etcetera", r"""문서분량 (표, 그림 등)""")
    DocVersion = PropertyDescriptor("DocVersion", r"""문서 파일 버전 (한글2007에 새로 추가)""")
    HwpVersion = PropertyDescriptor("HwpVersion", r"""문서를 생성한 한글 워드프로그램의 버전 (한글2007에 새로 추가)""")
    HanjaChar = PropertyDescriptor("HanjaChar", r"""문서분량 (한자 수) (한글2007에 새로 추가)""")


class VersionInfo(ParameterSet):
    """VersionInfo ParameterSet."""
    SourcePath = PropertyDescriptor("SourcePath", r"""버전 비교용 소스 패스""")
    TargetPath = PropertyDescriptor("TargetPath", r"""버전 비교용 타겟 패스""")
    ItemStartIndex = PropertyDescriptor("ItemStartIndex", r"""버전 비교를 보여줄 시작 히스토리 인덱스""")
    ItemEndIndex = PropertyDescriptor("ItemEndIndex", r"""버전 비교를 보여줄 마지막 히스토리 인덱스""")
    ItemOverWrite = PropertyDescriptor("ItemOverWrite", r"""히스토리 정보를 저장할 때 마지막 버전으로 덮어쓰는 플랙 (on/off)""")
    ItemSaveDescription = PropertyDescriptor("ItemSaveDescription", r"""히스토리 정보를 저장할 때 설명을 입력하는 대화상자를 띄우는 플랙 (on/off)""")
    TempFilePath = PropertyDescriptor("TempFilePath", r"""버전 비교용 임시파일 경로""")
    ItemInfoIndex = PropertyDescriptor("ItemInfoIndex", r"""버전 정보 얻어오기 및 삭제 시 사용될 인덱스""")
    SaveFilePath = PropertyDescriptor("SaveFilePath", r"""버전 저장 파일 경로(OCX 컨트롤용)""")
    ItemInfoWriter = PropertyDescriptor("ItemInfoWriter", r"""작성자 정보""")
    ItemInfoDescription = PropertyDescriptor("ItemInfoDescription", r"""해당 버전에 대한 설명""")
    ItemInfoTimeHi = PropertyDescriptor("ItemInfoTimeHi", r"""날짜 정보, FILETIME의 HIWORD""")
    ItemInfoTimeLo = PropertyDescriptor("ItemInfoTimeLo", r"""날짜 정보, FILETIME의 LOWORD""")
    ItemInfoLock = PropertyDescriptor("ItemInfoLock", r"""히스토리 정보 수정 플랙""")

