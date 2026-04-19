"""
Drawing / shape ParameterSet classes (25 classes).

Contains the ``ShapeObject`` pset and all ``Draw*`` variants for
graphical objects (shapes, lines, fills, images, rotations, etc.).

ShapeObject is the common parent pset — it has 53 properties covering
layout, flow, margins, and references to nested Draw* psets (via string
descriptors, not Python imports).
"""
from __future__ import annotations
from hwpapi.low.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)
from hwpapi.low.parametersets.mappings import (
    FONTTYPE_MAP, ROTATION_SETTING_MAP, SHADOW_TYPE_MAP,
)


class DrawFillAttr(ParameterSet):
    """
    DrawFillAttr ParameterSet - 그리기 개체의 채우기 속성.
    """
    Type = PropertyDescriptor("Type", r"""배경 유형: 0=채우기 없음, 1=면색/무늬색, 2=그림, 3=그러데이션""")
    GradationType = PropertyDescriptor("GradationType", r"""그러데이션 형태: 1=줄무늬형, 2=원형, 3=원뿔형, 4=사각형""")
    GradationAngle = PropertyDescriptor("GradationAngle", r"""그러데이션 기울임(시작각)""")
    GradationCenterX = PropertyDescriptor("GradationCenterX", r"""그러데이션 가로 중심 X 좌표""")
    GradationCenterY = PropertyDescriptor("GradationCenterY", r"""그러데이션 세로 중심 Y 좌표""")
    GradationStep = PropertyDescriptor("GradationStep", r"""그러데이션 번짐 정도 (0~100)""")
    GradationColorNum = ColorProperty("GradationColorNum", r"""그러데이션 색 수""")
    GradationColor = ColorProperty("GradationColor", r"""그러데이션 색깔 배열 (PIT_ARRAY)""")
    GradationIndexPos = PropertyDescriptor("GradationIndexPos", r"""그러데이션 다음 색깔과의 거리 배열 (PIT_ARRAY)""")
    GradationStepCenter = PropertyDescriptor("GradationStepCenter", r"""그러데이션 번짐 정도의 중심 (0~100)""")
    GradationAlpha = PropertyDescriptor("GradationAlpha", r"""그러데이션 투명도 (한글2007)""")
    GradationBrush = PropertyDescriptor("GradationBrush", r"""현재 브러시가 그러데이션 브러시인지 여부""")
    WinBrushFaceColor = ColorProperty("WinBrushFaceColor", r"""면 색 (RGB 0x00BBGGRR)""")
    WinBrushHatchColor = ColorProperty("WinBrushHatchColor", r"""무늬 색 (RGB 0x00BBGGRR)""")
    WinBrushFaceStyle = PropertyDescriptor("WinBrushFaceStyle", r"""무늬 스타일""")
    WinBrushAlpha = PropertyDescriptor("WinBrushAlpha", r"""면/무늬 색 투명도""")
    WindowsBrush = PropertyDescriptor("WindowsBrush", r"""현재 브러시가 면/무늬 브러시인지 여부""")
    FileName = PropertyDescriptor("FileName", r"""그림 파일 경로""")
    Embedded = PropertyDescriptor("Embedded", r"""그림 삽입 방식: TRUE=문서에 삽입, FALSE=파일로 연결""")
    PicEffect = PropertyDescriptor("PicEffect", r"""그림 효과: 0=원본, 1=그레이스케일, 2=흑백""")
    Brightness = PropertyDescriptor("Brightness", r"""명도 (-100~100)""")
    Contrast = PropertyDescriptor("Contrast", r"""밝기 (-100~100)""")
    Reverse = PropertyDescriptor("Reverse", r"""반전 유무""")
    DrawFillImageType = PropertyDescriptor("DrawFillImageType", r"""배경 채우기 방식: 0=바둑판식, 5=크기에 맞추어, 6=가운데 등""")
    SkipLeft = PropertyDescriptor("SkipLeft", r"""왼쪽 자르기""")
    SkipRight = PropertyDescriptor("SkipRight", r"""오른쪽 자르기""")
    SkipTop = PropertyDescriptor("SkipTop", r"""위 자르기""")
    SkipBottom = PropertyDescriptor("SkipBottom", r"""아래 자르기""")
    OriginalSizeX = PropertyDescriptor("OriginalSizeX", r"""이미지 원본 크기 X""")
    OriginalSizeY = PropertyDescriptor("OriginalSizeY", r"""이미지 원본 크기 Y""")
    InsideMarginLeft = PropertyDescriptor("InsideMarginLeft", r"""이미지 안쪽 여백 왼쪽""")
    InsideMarginRight = PropertyDescriptor("InsideMarginRight", r"""이미지 안쪽 여백 오른쪽""")
    InsideMarginTop = PropertyDescriptor("InsideMarginTop", r"""이미지 안쪽 여백 위""")
    InsideMarginBottom = PropertyDescriptor("InsideMarginBottom", r"""이미지 안쪽 여백 아래""")
    ImageBrush = PropertyDescriptor("ImageBrush", r"""현재 브러시가 그림 브러시인지 여부""")
    ImageCreateOnDrag = PropertyDescriptor("ImageCreateOnDrag", r"""그림 개체 생성 시 마우스로 끌어 생성할지 여부 (한글2007)""")
    ImageAlpha = PropertyDescriptor("ImageAlpha", r"""그림 개체/배경 투명도 (한글2007)""")


class ShapeObject(ParameterSet):
    """ShapeObject ParameterSet."""
    TreatAsChar = PropertyDescriptor("TreatAsChar", r"""글자처럼 취급 on / off""")
    AffectsLine = PropertyDescriptor("AffectsLine", r"""줄 간격에 영향을 줄지 여부 on / off (TreatAsChar가 TRUE일 경우에만 사용된다)""")
    VertRelTo = PropertyDescriptor("VertRelTo", r"""세로 위치의 기준.0 \= 종이 영역(Paper)1 \= 쪽 영역(Page)2 \= 문단 영역(Paragraph)(TreatAsChar가 FALSE일 경우에만 사용된다)""")
    VertAlign = PropertyDescriptor("VertAlign", r"""VertRelTo값에 따른 상대적인 정렬 기준.VertRelTo값이 2(문단영역)일 경우 0 값만 사용할 수 있다.0 \= 위(Top)1 \= 가운데(Center)2 \= 아래(Bottom)""")
    VertOffset = PropertyDescriptor("VertOffset", r"""VertRelTo와 VertAlign을 기준으로 한 Y축 위치 오프셋 값. HWPUNIT 단위.""")
    HorzRelTo = PropertyDescriptor("HorzRelTo", r"""가로 위치의 기준.0 \= 종이 영역(Paper)1 \= 쪽 영역(Page)2 \= 다단 영역(Column)3 \= 문단 영역(Paragraph)(TreatAsChar가 FALSE일 경우에만 사용된다)""")
    HorzAlign = PropertyDescriptor("HorzAlign", r"""HorzRelTo값에 따른 상대적인 정렬 기준.HorzRelTo값이 3(문단영역)일 경우 0\~2 사이의 값만 사용할 수 있다.0 \= 왼쪽(Left)1 \= 가운데(Center)2 \= 오른쪽(Right)3 \= 안쪽(Inside)4 \= 바깥쪽(Outside)""")
    HorzOffset = PropertyDescriptor("HorzOffset", r"""HorzRelTo와 HorzAlign을 기준으로 한 X축 위치 오프셋 값. HWPUNIT 단위.""")
    FlowWithText = PropertyDescriptor("FlowWithText", r"""그리기 개체의 세로 위치를 쪽 영역 안쪽으로 제한할지 여부 on / offVertRelTo값이 2(문단영역)일 경우에만 의미가 있다.""")
    AllowOverlap = PropertyDescriptor("AllowOverlap", r"""다른 개체와 겹치는 것을 허용할지 여부 on / offTreatAsChar가 FALSE일 때만 의미가 있으며, FlowWithText가 TRUE이면 AllowOverlap은 항상 FALSE로 간주한다.""")
    WidthRelTo = PropertyDescriptor("WidthRelTo", r"""개체 너비 기준.0 \= 종이(Paper)1 \= 쪽(Page)2 \= 다단(Column)3 \= 문단(Paragraph)4 \= 고정 값(Absolute)""")
    Width = PropertyDescriptor("Width", r"""개체 너비 값.WidthRelTo에 따라 값의 의미 및 단위가 달라진다.""")
    HeightRelTo = PropertyDescriptor("HeightRelTo", r"""개체 높이 기준.0 \= 종이(Paper)1 \= 쪽(Page)2 \= 고정 값(Absolute)""")
    Height = PropertyDescriptor("Height", r"""개체 높이 값.HeightRelTo에 다라 값의 의미 및 단위가 달라진다.""")
    ProtectSize = PropertyDescriptor("ProtectSize", r"""크기 보호 on / off""")
    TextWrap = PropertyDescriptor("TextWrap", r"""그리기 개체와 본문 사이의 배치 방법.0 \= 어울림(Square)1 \= 자리 차지(Top \& Bottom)2 \= 글 뒤로(Behind Text)3 \= 글 앞으로(In front of Text)4 \= 경계를 명확히 지킴(Tight) \- 현재 사용안함5 \= 경계를 통과함(Through) \- 현재 사용안함(TreatAsChar가 FALSE일 경우에만 사용된다)""")
    TextFlow = PropertyDescriptor("TextFlow", r"""그리기 개체의 좌/우 어느 쪽에 글을 배치할지 지정하는 옵션. TextWrap의 값이 0일 때만 유효하다.0 \= 양쪽 모두(Both)1 \= 왼쪽만(Left Only)2 \= 오른쪽만(Right Only)3 \= 왼쪽과 오른쪽 중 넓은 쪽(Largest Only)""")
    OutsideMarginLeft = PropertyDescriptor("OutsideMarginLeft", r"""개체의 바깥 여백. (왼쪽) HWPUNIT 단위""")
    OutsideMarginRight = PropertyDescriptor("OutsideMarginRight", r"""개체의 바깥 여백. (오른쪽) HWPUNIT 단위""")
    OutsideMarginTop = PropertyDescriptor("OutsideMarginTop", r"""개체의 바깥 여백. (위) HWPUNIT 단위""")
    OutsideMarginBottom = PropertyDescriptor("OutsideMarginBottom", r"""개체의 바깥 여백. (아래) HWPUNIT 단위""")
    NumberingType = PropertyDescriptor("NumberingType", r"""이 개체가 속하는 번호 범주.0 \= 없음, 1 \= 그림, 2 \= 표, 3 \= 수식""")
    LayoutWidth = PropertyDescriptor("LayoutWidth", r"""개체가 페이지에 배열될 때 계산되는 폭의 값""")
    LayoutHeight = PropertyDescriptor("LayoutHeight", r"""개체가 페이지에 배열될 때 계산되는 높이 값""")
    Lock = PropertyDescriptor("Lock", r"""개체 보호하기 on / off""")
    HoldAnchorObj = PropertyDescriptor("HoldAnchorObj", r"""쪽 나눔 방지 on / off (한글2007에 새로 추가)""")
    PageNumber = PropertyDescriptor("PageNumber", r"""개체가 존재 하는 페이지 (한글2007에 새로 추가)""")
    AdjustSelection = PropertyDescriptor("AdjustSelection", r"""개체 Selection 상태 TRUE/FASLE (한글2007에 새로 추가)""")
    AdjustTextBox = PropertyDescriptor("AdjustTextBox", r"""글상자로 TRUE/FASLE (한글2007에 새로 추가)""")
    AdjustPrevObjAttr = PropertyDescriptor("AdjustPrevObjAttr", r"""앞개체 속성 따라가기 TRUE/FASLE (한글2007에 새로 추가)""")
    # 선택적 중첩 ParameterSet 속성 (PIT_SET 타입)
    ShapeDrawLayOut = PropertyDescriptor("ShapeDrawLayOut", r"""그리기 개체의 Layout (PIT_SET: DrawLayout)""")
    ShapeDrawLineAttr = PropertyDescriptor("ShapeDrawLineAttr", r"""그리기 개체의 Line 속성 (PIT_SET: DrawLineAttr)""")
    ShapeDrawFillAttr = PropertyDescriptor("ShapeDrawFillAttr", r"""그리기 개체의 Fill 속성 (PIT_SET: DrawFillAttr)""")
    ShapeDrawImageAttr = PropertyDescriptor("ShapeDrawImageAttr", r"""그림 개체 속성 (PIT_SET: DrawImageAttr)""")
    ShapeDrawRectType = PropertyDescriptor("ShapeDrawRectType", r"""사각형 그리기 개체 유형 (PIT_SET: DrawRectType)""")
    ShapeDrawArcType = PropertyDescriptor("ShapeDrawArcType", r"""호 그리기 개체 유형 (PIT_SET: DrawArcType)""")
    ShapeDrawResize = PropertyDescriptor("ShapeDrawResize", r"""그리기 개체 리사이징 (PIT_SET: DrawResize)""")
    ShapeDrawRotate = PropertyDescriptor("ShapeDrawRotate", r"""그리기 개체 회전 (PIT_SET: DrawRotate)""")
    ShapeDrawEditDetail = PropertyDescriptor("ShapeDrawEditDetail", r"""그리기 개체 EditDetail (PIT_SET: DrawEditDetail)""")
    ShapeDrawImageScissoring = PropertyDescriptor("ShapeDrawImageScissoring", r"""그림 개체 자르기 (PIT_SET: DrawImageScissoring)""")
    ShapeDrawScAction = PropertyDescriptor("ShapeDrawScAction", r"""그리기 개체 ScAction (PIT_SET: DrawScAction)""")
    ShapeDrawCtrlHyperlink = PropertyDescriptor("ShapeDrawCtrlHyperlink", r"""그리기 개체 하이퍼링크 (PIT_SET: DrawCtrlHyperlink)""")
    ShapeDrawCoordInfo = PropertyDescriptor("ShapeDrawCoordInfo", r"""그리기 개체 좌표정보 (PIT_SET: DrawCoordInfo)""")
    ShapeDrawShear = PropertyDescriptor("ShapeDrawShear", r"""그리기 개체 기울이기 (PIT_SET: DrawShear)""")
    ShapeDrawTextart = PropertyDescriptor("ShapeDrawTextart", r"""글맵시 (PIT_SET: DrawTextart)""")
    ShapeDrawShadow = PropertyDescriptor("ShapeDrawShadow", r"""그림자 (PIT_SET: DrawShadow)""")
    ShapeTableCell = PropertyDescriptor("ShapeTableCell", r"""셀 정보 (PIT_SET: Cell)""")
    ShapeListProperties = PropertyDescriptor("ShapeListProperties", r"""서브 list 속성 (PIT_SET: ListProperties)""")
    ShapeCaption = PropertyDescriptor("ShapeCaption", r"""캡션 (PIT_SET: Caption)""")
    ShapeType = PropertyDescriptor("ShapeType", r"""TablePropertyDialog 종류""")
    ShapeCellSize = PropertyDescriptor("ShapeCellSize", r"""셀 크기 적용 여부 (on/off)""")
    ShapeCreationType = PropertyDescriptor("ShapeCreationType", r"""그리기 개체 형태: 0=직선, 1=사각형, 2=원, 3=호""")
    ShapeCreationMode = PropertyDescriptor("ShapeCreationMode", r"""마우스로 그리기 여부 (on/off)""")


class DrawArcType(ParameterSet):
    """
    ### DrawArcType

    27) DrawArcType: 곡선 그리기의 필수적인 속성을 나타내는 클래스

    | Item ID  | Type       | SubType | Description |
    |----------|------------|---------|-------------|
    | Type     | PIT_UI     |         | 곡선 유형: 0 = 선, 1 = 필수곡선, 2 = 화살표 |
    | Interval | PIT_ARRAY  | PIT_I   | 곡선의 시작점과 끝점을 나타내는 배열 |
    """
    _ark_type_map = {"line": 0, "essential": 1, "arrow": 2}
    type = ParameterSet._mapped_prop("Type", _ark_type_map,
                                     doc="곡선 유형: 0 = 선, 1 = 필수곡선, 2 = 화살표")
    interval = ParameterSet._int_list_prop("Interval", "곡선의 시작점과 끝점을 나타내는 배열")



class DrawCoordInfo(ParameterSet):
    """
    ### DrawCoordInfo

    28) DrawCoordInfo : 그리기 좌표의 상세 정보

    CoordInfo(한글2005)에서 DrawCoordInfo로 이름이 변경되었습니다. 정보를 읽고 쓸 수 있도록 지원합니다.

    | Item ID | Type       | SubType | Description                                 |
    |---------|------------|---------|---------------------------------------------|
    | Count   | PIT_UI4    |         | 점의 개수                                  |
    | Point   | PIT_ARRAY  | PIT_I   | 좌표 Array (X1,Y1), (X2,Y2), ..., (Xn,Yn)   |
    | Line    | PIT_ARRAY  | PIT_UI1 | 선 정보 Array(점들에서 연결된 형태)          |
    """
    count = ParameterSet._int_prop("Count", "점의 개수: 정수 값을 입력하세요.")
    point = ParameterSet._tuple_list_prop("Point", "좌표 배열 (X1, Y1), (X2, Y2), ..., (Xn, Yn): 리스트 값을 입력하세요.")
    line  = ParameterSet._int_list_prop("Line", "선 정보 배열: 리스트 값을 입력하세요.")



class DrawCtrlHyperlink(ParameterSet):
    """
    ### DrawCtrlHyperlink

    29) DrawCtrlHyperlink : 그림 개체의 Hyperlink 정보

    CtrlHyperlink(한글2005)에서 DrawCtrlHyperlink로 이름이 변경되었습니다.

    | Item ID | Type      | SubType | Description                          |
    |---------|-----------|---------|--------------------------------------|
    | Command | PIT_BSTR  |         | Command String (명령 문자열)         |
    """
    command = ParameterSet._str_prop("Command", "Command String: 명령 문자열을 입력하세요.")



class DrawEditDetail(ParameterSet):
    """
    ### DrawEditDetail

    30) DrawEditDetail : 그림의 교정과 관련된 상세 설정

    | Item ID | Type   | SubType | Description |
    |---------|--------|---------|-------------|
    | Command | PIT_UI |         | Reserved    |
    | Index   | PIT_UI |         | 교점 정의의 인덱스 |
    | PointX  | PIT_I  |         | 교점의 X 좌표 |
    | PointY  | PIT_I  |         | 교점의 Y 좌표 |
    """
    command = ParameterSet._int_prop("Command", "Command: Reserved.")
    index   = ParameterSet._int_prop("Index", "Index: 교점 정의의 인덱스.")
    point_x = ParameterSet._int_prop("PointX", "PointX: 교점의 X 좌표.")
    point_y = ParameterSet._int_prop("PointY", "PointY: 교점의 Y 좌표.")



class DrawImageAttr(ParameterSet):
    """DrawImageAttr ParameterSet."""
    FileName = PropertyDescriptor("FileName", r"""ShapeObject의 배경을 그림으로 선택했을 경우 또는 그림개체일 경우의 그림파일 경로""")
    Embedded = PropertyDescriptor("Embedded", r"""그림이 문서에 직접 삽입(TRUE) / 파일로 연결(FALSE)""")
    PicEffect = PropertyDescriptor("PicEffect", r"""그림 효과 0 \= 실제 이미지 그대로1 \= 그레이스케일2 \= 흑백 효과""")
    Brightness = PropertyDescriptor("Brightness", r"""명도 (\-100 \~ 100\)""")
    Contrast = PropertyDescriptor("Contrast", r"""밝기 (\-100 \~ 100\)""")
    Reverse = PropertyDescriptor("Reverse", r"""반전 유무""")
    DrawFillImageType = PropertyDescriptor("DrawFillImageType", r"""ShapeObject의 배경일 경우에만 의미 있는 아이템, 배경을 채우는 방식을 결정한다. (그림개체에는 해당사항 없음)0 \= 바둑판식으로1 \= 가로/위만 바둑판식으로 배열2 \= 가로/아래만 바둑판식으로 배열3 \= 세로/왼쪽만 바둑판식으로 배열4 \= 세로/오른쪽만 바둑판식으로 배열5 \= 크기에 맞추어6 \= 가운데로7 \= 가운데 위로8 \= 가운데 아래로9 \= 왼쪽 가운데로10 \= 왼쪽 위로11 \= 왼쪽 아래로12 \= 오른쪽 가운데로13 \= 오른쪽 위로14 \= 오른쪽 아래로""")
    SkipLeft = PropertyDescriptor("SkipLeft", r"""그림 개체일 경우에만 의미 있는 아이템, 왼쪽 자르기""")
    SkipRight = PropertyDescriptor("SkipRight", r"""그림 개체일 경우에만 의미 있는 아이템, 오른쪽 자르기""")
    SkipTop = PropertyDescriptor("SkipTop", r"""그림 개체일 경우에만 의미 있는 아이템, 위 자르기""")
    SkipBottom = PropertyDescriptor("SkipBottom", r"""그림 개체일 경우에만 의미 있는 아이템, 아래 자르기""")
    OriginalSizeX = PropertyDescriptor("OriginalSizeX", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 원본 크기 X size""")
    OriginalSizeY = PropertyDescriptor("OriginalSizeY", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 원본 크기 Y size""")
    InsideMarginLeft = PropertyDescriptor("InsideMarginLeft", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (왼쪽)""")
    InsideMarginRight = PropertyDescriptor("InsideMarginRight", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (오른쪽)""")
    InsideMarginTop = PropertyDescriptor("InsideMarginTop", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (위)""")
    InsideMarginBottom = PropertyDescriptor("InsideMarginBottom", r"""그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (아래)""")
    WindowsBrush = PropertyDescriptor("WindowsBrush", r"""현재 선택된 brush의 type이 면/무늬 브러시인가를 나타냄""")
    GradationBrush = PropertyDescriptor("GradationBrush", r"""현재 선택된 brush의 type이 그러데이션 브러시인가를 나타냄""")
    ImageBrush = PropertyDescriptor("ImageBrush", r"""현재 선택된 brush의 type이 그림 브러시인가를 나타냄""")
    ImageCreateOnDrag = PropertyDescriptor("ImageCreateOnDrag", r"""그림 개체 생성 시 마우스로 끌어 생성할지 여부(한글2007에 새로 추가)""")


class DrawImageScissoring(ParameterSet):
    """
    ### DrawImageScissoring

    33) DrawImageScissoring : 그림 자르기의 좌표 정보

    ImageScissoring(한컴오피스 2005)에서 DrawImageScissoring으로 이름이 변경되었습니다.

    | Item ID       | Type    | SubType | Description          |
    |---------------|---------|---------|----------------------|
    | Xoffset       | PIT_I   |         | 자를 x좌표 오프셋      |
    | Yoffset       | PIT_I   |         | 자를 y좌표 오프셋      |
    | HandleIndex   | PIT_UI  |         | Reserved             |
    """
    x_offset = ParameterSet._int_prop("Xoffset", "자를 x좌표 오프셋: 정수 값을 입력하세요.")
    y_offset = ParameterSet._int_prop("Yoffset", "자를 y좌표 오프셋: 정수 값을 입력하세요.")
    handle_index = ParameterSet._int_prop("HandleIndex", "Reserved: 정수 값을 입력하세요.")



class DrawLayout(ParameterSet):
    """
    ### DrawLayout

    34) DrawLayout : 도형 레이아웃의 일반 속성

    | Item ID          | Type      | SubType | Description |
    |------------------|-----------|---------|-------------|
    | CreateNumPt      | PIT_UI    |         | 생성할 점의 수 |
    | CreatePt         | PIT_ARRAY | PIT_I   | 생성할 점의 좌표정보: POINT(x,y) 형식의 배열로, CreateNumPt * 2 개수만큼 구성 |
    | CurveSegmentInfo | PIT_ARRAY | PIT_UI1 | 곡선 세그먼트 정보 |
    """
    create_num_pt = ParameterSet._int_prop("CreateNumPt", "생성할 점의 수: 정수 값을 입력하세요.")
    create_pt = ParameterSet._tuple_list_prop("CreatePt", "생성할 점의 좌표정보: (x, y) 튜플의 리스트")
    curve_segment_info = ParameterSet._int_list_prop("CurveSegmentInfo", "곡선 세그먼트 정보: 정수 리스트")



class DrawLineAttr(ParameterSet):
    """
    ### DrawLineAttr

    35) DrawLineAttr : 도형 선의 속성

    | Item ID        | Type        | SubType | Description                                                 |
    |----------------|-------------|---------|-------------------------------------------------------------|
    | Color          | PMT_UINT32  |         | 선 색상, RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)    |
    | Style          | PMT_INT     |         | 선의 스타일                                                |
    | Width          | PMT_INT     |         | 선의 두께                                                  |
    | EndCap         | PMT_INT     |         | 선의 끝단                                                  |
    | HeadStyle      | PMT_INT     |         | 선의 시작 화살표 형태                                      |
    | TailStyle      | PMT_INT     |         | 선의 끝 화살표 형태                                        |
    | HeadSize       | PMT_INT     |         | 선의 시작 화살표 크기                                      |
    | TailSize       | PMT_INT     |         | 선의 끝 화살표 크기                                        |
    | HeadFill       | PMT_BOOL    |         | 선의 시작 화살표 채움 여부                                  |
    | TailFill       | PMT_BOOL    |         | 선의 끝 화살표 채움 여부                                    |
    | OutLineStyle   | PMT_UINT    |         | 외곽선 (안쪽/바깥쪽/중앙)                                   |
    | Alpha          | PIT_UI1     |         | 투명도 (한글 2007에 처음 추가됨)                            |
    """
    color         = ParameterSet._int_prop("Color", "선 색상 (RGB 0x00BBGGRR): 정수를 입력하세요.")
    style         = ParameterSet._int_prop("Style", "선의 스타일: 정수를 입력하세요.")
    width         = ParameterSet._int_prop("Width", "선의 두께: 정수를 입력하세요.")
    end_cap       = ParameterSet._int_prop("EndCap", "선의 끝단: 정수를 입력하세요.")
    head_style    = ParameterSet._int_prop("HeadStyle", "선의 시작 화살표 형태: 정수를 입력하세요.")
    tail_style    = ParameterSet._int_prop("TailStyle", "선의 끝 화살표 형태: 정수를 입력하세요.")
    head_size     = ParameterSet._int_prop("HeadSize", "선의 시작 화살표 크기: 정수를 입력하세요.")
    tail_size     = ParameterSet._int_prop("TailSize", "선의 끝 화살표 크기: 정수를 입력하세요.")
    head_fill     = ParameterSet._bool_prop("HeadFill", "선의 시작 화살표 채움 여부: 0 또는 1을 입력하세요.")
    tail_fill     = ParameterSet._bool_prop("TailFill", "선의 끝 화살표 채움 여부: 0 또는 1을 입력하세요.")
    outline_style = ParameterSet._int_prop("OutLineStyle", "외곽선 (안쪽/바깥쪽/중앙): 정수를 입력하세요.")
    alpha         = ParameterSet._int_prop("Alpha", "투명도: 정수를 입력하세요.")



class DrawRectType(ParameterSet):
    """
    ### DrawRectType

    36) DrawRectType : 사각형 도형의 속성

    | Item ID | Type    | SubType | Description                           |
    |---------|---------|---------|---------------------------------------|
    | Type    | PIT_UI  |         | 도형의 종류 지정 (0 ~ 50까지)          |
    """
    type = ParameterSet._int_prop("Type", "도형의 종류: 0 ~ 50 사이의 정수 값을 입력하세요.", 0, 50)



class DrawResize(ParameterSet):
    """
    ### DrawResize

    37) DrawResize : 도형 크기 조정 Resizing 정보

    | Item ID      | Type    | SubType | Description                   |
    |--------------|---------|---------|-------------------------------|
    | Xoffset      | PIT_I   |         | 도형 크기 조정 X좌표 오프셋    |
    | Yoffset      | PIT_I   |         | 도형 크기 조정 Y좌표 오프셋    |
    | HandleIndex  | PIT_UI  |         | Reserved                      |
    | Mode         | PIT_UI  |         | Reserved                      |
    """
    x_offset = ParameterSet._int_prop("Xoffset", "도형 크기 조정 X좌표 오프셋: 정수 값을 입력하세요.")
    y_offset = ParameterSet._int_prop("Yoffset", "도형 크기 조정 Y좌표 오프셋: 정수 값을 입력하세요.")
    handle_index = ParameterSet._int_prop("HandleIndex", "Reserved: 정수 값을 입력하세요.")
    mode = ParameterSet._int_prop("Mode", "Reserved: 정수 값을 입력하세요.")



class DrawRotate(ParameterSet):
    """
    ### DrawRotate

    | Item ID         | Type   | Description                                  |
    |-----------------|--------|----------------------------------------------|
    | Command         | PIT_UI | 회전 설정의 기초 설정 (0: 없음, 1: 설정된 회전, 2: 그림 중심 회전, 3: 회전과 중심) |
    | CenterX         | PIT_I  | 회전 중심의 X 좌표                           |
    | CenterY         | PIT_I  | 회전 중심의 Y 좌표                           |
    | ObjectCenterX   | PIT_I  | 그림 중심의 X 좌표                           |
    | ObjectCenterY   | PIT_I  | 그림 중심의 Y 좌표                           |
    | Angle           | PIT_I  | 회전 각도                                    |
    | RotateImage     | PIT_UI1| 그림 회전 여부 (0: 회전 안 함, 1: 회전함)     |
    """
    command         = ParameterSet._mapped_prop("Command", ROTATION_SETTING_MAP,
                                                doc="회전 설정의 기초 설정 (0: 없음, 1: 설정된 회전, 2: 그림 중심 회전, 3: 회전과 중심)")
    center_x        = ParameterSet._int_prop("CenterX", "회전 중심의 X 좌표.")
    center_y        = ParameterSet._int_prop("CenterY", "회전 중심의 Y 좌표.")
    object_center_x = ParameterSet._int_prop("ObjectCenterX", "그림 중심의 X 좌표.")
    object_center_y = ParameterSet._int_prop("ObjectCenterY", "그림 중심의 Y 좌표.")
    angle           = ParameterSet._int_prop("Angle", "회전 각도.")
    rotate_image    = ParameterSet._bool_prop("RotateImage", "그림 회전 여부 (0: 회전 안 함, 1: 회전함).")



class DrawScAction(ParameterSet):
    """
    ### DrawScAction

    39) DrawScAction : 회전 중심과 90도 회전, 좌우/상하 플립 설정

    ScAction(한글2005)에서 DrawScAction으로 이름이 변경되었습니다.

    | Item ID        | Type    | SubType | Description                 |
    |----------------|---------|---------|-----------------------------|
    | RotateCenterX  | PIT_I4  |         | 회전 중심 X 좌표           |
    | RotateCenterY  | PIT_I4  |         | 회전 중심 Y 좌표           |
    | RotateAngel    | PIT_I   |         | 회전각                     |
    | HorzFlip       | PIT_UI  |         | 수평 flip (좌우 대칭 설정) |
    | VertFlip       | PIT_UI  |         | 수직 flip (상하 대칭 설정) |
    """
    rotate_center_x = ParameterSet._int_prop("RotateCenterX", "회전 중심 X 좌표")
    rotate_center_y = ParameterSet._int_prop("RotateCenterY", "회전 중심 Y 좌표")
    rotate_angle    = ParameterSet._int_prop("RotateAngel", "회전각")
    horz_flip       = ParameterSet._bool_prop("HorzFlip", "수평 flip (좌우 대칭 설정)")
    vert_flip       = ParameterSet._bool_prop("VertFlip", "수직 flip (상하 대칭 설정)")



class DrawShadow(ParameterSet):
    """
    ### DrawShadow

    40) DrawShadow : 그림자 효과 정보

    | Item ID       | Type     | SubType | Description                                     |
    |---------------|----------|---------|-------------------------------------------------|
    | ShadowType    | PIT_I4   |         | 그림자 종류: 0 = none, 1 = drop, 2 = continuous   |
    | ShadowColor   | PIT_UI4  |         | 그림자 색상 (COLORREF)                           |
    | ShadowOffsetX | PIT_I4   |         | 그림자 X축 오프셋 (-48% ~ 48%)                   |
    | ShadowOffsetY | PIT_I4   |         | 그림자 Y축 오프셋 (-48% ~ 48%)                   |
    | ShadowAlpha   | PIT_UI1  |         | 그림자 투명도 (0 ~ 255)                          |
    """
    shadow_type    = ParameterSet._mapped_prop("ShadowType", SHADOW_TYPE_MAP,
                                              doc="그림자 종류: 0 = none, 1 = drop, 2 = continuous")
    shadow_color   = ParameterSet._int_prop("ShadowColor", "그림자 색상 (COLORREF): 정수를 입력하세요.")
    shadow_offset_x = ParameterSet._int_prop("ShadowOffsetX", "그림자 X축 오프셋 (-48% ~ 48%)", -48, 48)
    shadow_offset_y = ParameterSet._int_prop("ShadowOffsetY", "그림자 Y축 오프셋 (-48% ~ 48%)", -48, 48)
    shadow_alpha   = ParameterSet._int_prop("ShadowAlpha", "그림자 투명도 (0 ~ 255)", 0, 255)



class DrawShear(ParameterSet):
    """
    ### DrawShear

    41) DrawShear : Shear transformation parameters

    | Item ID | Type   | SubType | Description        |
    |---------|--------|---------|--------------------|
    | XFactor | PIT_I  |         | X shear factor     |
    | YFactor | PIT_I  |         | Y shear factor     |
    """
    x_factor = ParameterSet._int_prop("XFactor", "X shear factor: 정수 값을 입력하세요.")
    y_factor = ParameterSet._int_prop("YFactor", "Y shear factor: 정수 값을 입력하세요.")



class DrawTextart(ParameterSet):
    """
    ### DrawTextart

    42) DrawTextart : 텍스트아트 속성
    """
    string         = ParameterSet._str_prop("String", "텍스트아트 내용: 문자열 값을 입력하세요.")
    font_name      = ParameterSet._str_prop("FontName", "폰트 이름.")
    font_style     = ParameterSet._str_prop("FontStyle", "폰트 스타일 (0 = Regular).")
    font_type      = ParameterSet._mapped_prop("FontType", FONTTYPE_MAP,
                                             doc="폰트 타입: 0 = don't care, 1 = TTF, 2 = HFT.")
    line_spacing   = ParameterSet._int_prop("LineSpacing", "줄 간격 (50 ~ 500).", 50, 500)
    char_spacing   = ParameterSet._int_prop("CharSpacing", "문자 간격 (50 ~ 500).", 50, 500)
    align_type     = ParameterSet._int_prop("AlignType", "정렬 유형.")
    shape          = ParameterSet._int_prop("Shape", "형태 (0 ~ 54).", 0, 54)
    shadow_type    = ParameterSet._mapped_prop("ShadowType", SHADOW_TYPE_MAP,
                                              doc="그림자 유형: 0 = none, 1 = drop, 2 = continuous.")
    shadow_offset_x = ParameterSet._int_prop("ShadowOffsetX", "그림자 X 오프셋 (-48 ~ 48).", -48, 48)
    shadow_offset_y = ParameterSet._int_prop("ShadowOffsetY", "그림자 Y 오프셋 (-48 ~ 48).", -48, 48)
    shadow_color   = ParameterSet._int_prop("ShadowColor", "그림자 색상 (RGB 32비트).")
    number_of_lines = ParameterSet._int_prop("NumberOfLines", "텍스트아트의 줄 수.")



class DropCap(ParameterSet):
    """DropCap ParameterSet."""
    Style = PropertyDescriptor("Style", r"""글자 장식 모양0\=없음1\=2줄차지2\=3줄차지4\=여백""")
    FaceName = PropertyDescriptor("FaceName", r"""Parameter property""")
    LineStyle = PropertyDescriptor("LineStyle", r"""선 스타일""")
    LineWeight = PropertyDescriptor("LineWeight", r"""선 굵기""")
    LineColor = ColorProperty("LineColor", r"""선 색RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    FaceColor = ColorProperty("FaceColor", r"""글꼴 색RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    Spacing = PropertyDescriptor("Spacing", r"""본문과의 간격""")


class ShapeCopyPaste(ParameterSet):
    """ShapeCopyPaste ParameterSet."""
    Type = PropertyDescriptor("Type", r"""모양 복사 종류0 \= 글자 모양 복사1 \= 문단 모양 복사2 \= 글자와 문단 모양 두개 복사3 \= 글자 스타일 복사4 \= 문단 스타일 복사""")
    CellAttr = PropertyDescriptor("CellAttr", r"""셀 모양 복사""")
    CellBorder = PropertyDescriptor("CellBorder", r"""선 모양 복사""")
    CellFill = PropertyDescriptor("CellFill", r"""셀 배경 복사""")
    TypeBodyAndCellOnly = PropertyDescriptor("TypeBodyAndCellOnly", r"""본문과 셀 모양 둘 다 복사 or 셀 모양만 복사(한글2007에 새로 추가)""")


class ShapeObjectCopyPaste(ParameterSet):
    """ShapeObjectCopyPaste ParameterSet."""
    Type = PropertyDescriptor("Type", r"""그리기 모양 복사/붙여 넣기 종류 (예약.. 현재 사용하지 않음)""")
    ShapeObjectLine = PropertyDescriptor("ShapeObjectLine", r"""그리기 선 모양 복사""")
    ShapeObjectFill = PropertyDescriptor("ShapeObjectFill", r"""그리기 채우기 복사""")
    ShapeObjectSize = PropertyDescriptor("ShapeObjectSize", r"""그리기 개체 크기 복사""")
    ShapeObjectShadow = PropertyDescriptor("ShapeObjectShadow", r"""그리기 개체 그림자 복사""")
    ShapeObjectPicEffect = PropertyDescriptor("ShapeObjectPicEffect", r"""그림 효과 복사""")


class CCLMark(ParameterSet):
    """CCLMark ParameterSet - CCL 마크 삽입."""


class ChartObjShape(ParameterSet):
    """ChartObjShape ParameterSet - 차트 개체."""


class DrawObjTemplateSave(ParameterSet):
    """DrawObjTemplateSave ParameterSet - 그리기 마당에 등록."""


class FindImagePath(ParameterSet):
    """FindImagePath ParameterSet - 그림 경로 찾기."""


class ShapeObjSaveAsPicture(ParameterSet):
    """ShapeObjSaveAsPicture ParameterSet - 그리기 개체를 그림으로 저장."""

