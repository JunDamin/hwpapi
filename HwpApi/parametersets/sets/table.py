"""
Table-related ParameterSet classes (12 classes).

Covers table creation, modification, and formatting:
- Table, TableCreation, TableTemplate
- CellBorderFill (cell-level), AutoFill
- TableSplitCell, TableSwap, TableStrToTbl, TableTblToStr
- TableDeleteLine, TableInsertLine, TableDrawPen

Note: ``Cell`` itself is in ``primitives.py`` as it's referenced
across multiple domains.
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class Table(ParameterSet):
    """Table ParameterSet."""
    PageBreak = PropertyDescriptor("PageBreak", r"""표가 페이지 경계에 걸렸을 때의 처리 방식0 \= 나누지 않는다.1 \= 테이블은 나누지만 셀은 나누지 않는다.2 \= 셀 내의 텍스트도 나눈다.""")
    RepeatHeader = PropertyDescriptor("RepeatHeader", r"""제목 행을 반복할지 여부. (on / off)""")
    CellSpacing = PropertyDescriptor("CellSpacing", r"""셀 간격(HTML의 셀 간격과 동일 의미. HWPUNIT)""")
    CellMarginLeft = PropertyDescriptor("CellMarginLeft", r"""기본 셀 안쪽 여백(왼쪽)""")
    CellMarginRight = PropertyDescriptor("CellMarginRight", r"""기본 셀 안쪽 여백(오른쪽)""")
    CellMarginTop = PropertyDescriptor("CellMarginTop", r"""기본 셀 안쪽 여백(위쪽)""")
    CellMarginBottom = PropertyDescriptor("CellMarginBottom", r"""기본 셀 안쪽 여백(아래쪽)""")
    BorderFill = PropertyDescriptor("BorderFill", r"""표에 적용되는 테두리/배경""")
    TableCharInfo = PropertyDescriptor("TableCharInfo", r"""표와 연결된 차트 정보 \- 차트 미완성""")
    TableBorderFill = PropertyDescriptor("TableBorderFill", r"""표에 적용되는 테두리/배경""")
    Cell = PropertyDescriptor("Cell", r"""셀 속성""")


class AutoFill(ParameterSet):
    """AutoFill ParameterSet."""
    AutoFillSection = PropertyDescriptor("AutoFillSection", r"""자동 채우기 섹션 : 0 \= 기본, 1 \= 사용자 정의""")
    AutoFillItem = PropertyDescriptor("AutoFillItem", r"""섹션의 아이템 인덱스 : 0 \~""")


class CellBorderFill(ParameterSet):
    """CellBorderFill ParameterSet."""
    ApplyTo = PropertyDescriptor("ApplyTo", r"""적용 대상 : 0 \= 선택된 셀, 1 \= 전체 셀, 2 \= 여러 셀에 걸쳐 적용""")
    NoNeighborCell = PropertyDescriptor("NoNeighborCell", r"""주변 셀에 선 모양을 적용하지 않을지 여부 (1이면 적용하지 않는다)""")
    TableBorderFill = PropertyDescriptor("TableBorderFill", r"""표 테두리/배경""")
    AllCellsBorderFill = PropertyDescriptor("AllCellsBorderFill", r"""전체 셀의 테두리/배경""")
    SelCellsBorderFill = PropertyDescriptor("SelCellsBorderFill", r"""선택된 셀의 테두리/배경""")


class TableCreation(ParameterSet):
    """TableCreation ParameterSet."""
    Rows = PropertyDescriptor("Rows", r"""행 수 (생략하면 5\)""")
    Cols = PropertyDescriptor("Cols", r"""칼럼 수 (생략하면 5\)""")
    RowHeight = PropertyDescriptor("RowHeight", r"""행의 디폴트 높이 (PIT_I4\)""")
    ColWidth = PropertyDescriptor("ColWidth", r"""칼럼의 디폴트 폭 (PIT_I4\)""")
    CellInfo = PropertyDescriptor("CellInfo", r"""정보가 없는 셀은 디폴트값을 따라가므로 모든 셀에 대해 정보를 줄 필요는 없다.""")
    WidthType = PropertyDescriptor("WidthType", r"""Parameter property""")
    HeightType = PropertyDescriptor("HeightType", r"""Parameter property""")
    WidthValue = PropertyDescriptor("WidthValue", r"""너비 값""")
    HeightValue = PropertyDescriptor("HeightValue", r"""높이 값""")
    TableTemplateValue = PropertyDescriptor("TableTemplateValue", r"""표 마당 적용 여부 (한글2007에 새로 추가)""")
    TableProperties = PropertyDescriptor("TableProperties", r"""초기 표 속성""")
    TableTemplate = PropertyDescriptor("TableTemplate", r"""표마당 적용 속성 (한글2007에 새로 추가)""")
    TableDrawProperties = PropertyDescriptor("TableDrawProperties", r"""마우스로 선을 그릴 때 속성 (한글2007에 새로 추가)""")


class TableDeleteLine(ParameterSet):
    """TableDeleteLine ParameterSet."""
    Type = PropertyDescriptor("Type", r"""0 \= 줄, 1 \= 칸""")


class TableDrawPen(ParameterSet):
    """TableDrawPen ParameterSet."""
    Style = PropertyDescriptor("Style", r"""Table을 그리는 연필(펜)의 선 모양""")
    Width = PropertyDescriptor("Width", r"""Table을 그리는 연필(펜)의 선 굵기""")
    Color = PropertyDescriptor("Color", r"""Table을 그리는 연필(펜)의 선 색깔RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")


class TableInsertLine(ParameterSet):
    """TableInsertLine ParameterSet."""
    Side = PropertyDescriptor("Side", r"""Parameter property""")
    Count = PropertyDescriptor("Count", r"""Parameter property""")


class TableSplitCell(ParameterSet):
    """TableSplitCell ParameterSet."""
    Cols = PropertyDescriptor("Cols", r"""칸 수""")
    Rows = PropertyDescriptor("Rows", r"""줄 수""")
    DistributeHeight = PropertyDescriptor("DistributeHeight", r"""줄 높이를 같게""")
    Merge = PropertyDescriptor("Merge", r"""나누기 전에 합치기""")
    Mode2 = PropertyDescriptor("Mode2", r"""셀 나누기 모드 2, 셀 나누기를 할 때, adjust를 생략하고 셀이 어긋나는 것을 방지한다. (한글2007에 새로 추가)""")


class TableStrToTbl(ParameterSet):
    """TableStrToTbl ParameterSet."""
    DelimiterType = PropertyDescriptor("DelimiterType", r"""분리 문자(탭, 쉼표, 공백)""")
    UserDefine = PropertyDescriptor("UserDefine", r"""사용자 정의 필드 구분 기호""")
    AutoOrDefine = PropertyDescriptor("AutoOrDefine", r"""자동으로 할 것인지 분리 문자를 지정 할 것인지를 결정""")
    KeepSeperator = PropertyDescriptor("KeepSeperator", r"""선택 사항 (구분자 유지)""")
    DelimiterEtc = PropertyDescriptor("DelimiterEtc", r"""기타 문자 필드 구분 기호""")


class TableSwap(ParameterSet):
    """TableSwap ParameterSet."""
    Type = PropertyDescriptor("Type", r"""표 뒤집기 형식0 \= 상하 뒤집기1 \= 좌우 뒤집기2 \= X와 Y를 바꿈3 \= 반시계 방향으로 90도 회전4 \= 180도 회전5 \= 시계 방향으로 90도 회전""")
    SwapMargin = PropertyDescriptor("SwapMargin", r"""여백 뒤집기 지원여부""")


class TableTblToStr(ParameterSet):
    """TableTblToStr ParameterSet."""
    DelimiterType = PropertyDescriptor("DelimiterType", r"""분리 문자(탭, 쉼표, 공백)""")
    UserDefine = PropertyDescriptor("UserDefine", r"""사용자 정의 필드 구분 기호""")


class TableTemplate(ParameterSet):
    """TableTemplate ParameterSet."""
    Format = PropertyDescriptor("Format", r"""적용할 서식. 다음 값의 조합으로 구성된다.0x0001 \= 테두리0x0002 \= 글자 모양과 문단 모양0x0004 \= 셀 배경0x0008 \= 그레이 스케일""")
    ApplyTarger = PropertyDescriptor("ApplyTarger", r"""적용 대상. 다음 값의 조합으로 구성된다.0x0001 \= 제목 줄0x0002 \= 마지막 줄0x0004 \= 첫째 칸0x0008 \= 마지막 칸""")
    FileName = PropertyDescriptor("FileName", r"""표 마당 파일 이름 ？ C:\\Program Files\\Hnc\\Shared80\\HwpTemplate\\Table\\Kor에 있는 파일명임""")
    CreateMode = PropertyDescriptor("CreateMode", r"""표 만들기 모드 (표 만들기에서 제목줄에 제목 속성 넣기 위해)""")

