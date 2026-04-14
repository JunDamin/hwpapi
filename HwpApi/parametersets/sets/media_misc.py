"""
Media / macros / miscellaneous ParameterSet classes (27 classes).

Grab-bag of specialized psets:
- OleCreation, MovieProperties, FlashProperties
- HyperLink, HyperlinkJump
- FieldCtrl, InsertFieldTemplate
- Idiom, LinkDocument, Internet
- FtpDownload, FtpUpload, MemoShape
- Sort, Sum, AutoNum, CaptureEnd, EqEdit
- InputDateStyle, Label, MailMergeGenerate, MousePos
- ViewProperties, ViewStatus
- ScriptMacro, KeyMacro, UserQCommandFile
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class AutoNum(ParameterSet):
    """AutoNum ParameterSet."""
    NumType = PropertyDescriptor("NumType", r"""번호 종류 : 0 \= 쪽 번호1 \= 각주 번호2 \= 미주 번호3 \= 그림 번호4 \= 표 번호5 \= 수식 번호""")
    NewNumber = PropertyDescriptor("NewNumber", r"""새 시작 번호 (1 .. n)""")


class CaptureEnd(ParameterSet):
    """CaptureEnd ParameterSet."""
    BeginX = PropertyDescriptor("BeginX", r"""시작점과 X 좌표(페이지 X좌표)""")
    BeginY = PropertyDescriptor("BeginY", r"""시작점과 Y 좌표(페이지 Y좌표)""")
    EndX = PropertyDescriptor("EndX", r"""끝점의 X 좌표(페이지 X좌표)""")
    EndY = PropertyDescriptor("EndY", r"""끝점의 Y 좌표(페이지 Y좌표)""")
    PageNum = PropertyDescriptor("PageNum", r"""페이지 번호""")


class EqEdit(ParameterSet):
    """EqEdit ParameterSet."""
    String = PropertyDescriptor("String", r"""수식 스크립트.""")
    BaseUnit = PropertyDescriptor("BaseUnit", r"""수식이 삽입되는 앞의 글자와 같은 높이 (기본 값은 POINT 10 )""")
    Color = PropertyDescriptor("Color", r"""수식이 삽입되는 글자 색과 같은 색 (기본 값은 WINDOWTEXT 색)RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)""")
    LineMode = PropertyDescriptor("LineMode", r"""줄 단위를 사용할지의 여부 (한글2007에 새로 추가)""")
    Version = PropertyDescriptor("Version", r"""수식 스크립트 버전 정보 (한글2007에 새로 추가)""")
    ApplyTo = PropertyDescriptor("ApplyTo", r"""수식 속성 적용 범위 (한글2007에 새로 추가)0 : 선택된 수식 개체1 : 문서 전체""")


class FieldCtrl(ParameterSet):
    """FieldCtrl ParameterSet."""
    Command = PropertyDescriptor("Command", r"""필드의 명령문""")
    Editable = PropertyDescriptor("Editable", r"""일부분 readonly mode에서 편집 가능한 필드인지 여부""")
    FieldDirty = PropertyDescriptor("FieldDirty", r"""필드가 초기화 상태인지 수정되어 있는 상태인지 여부(한글2007에 새로 추가)""")
    CtrlData = PropertyDescriptor("CtrlData", r"""필드 이름 저장을 위한 영역""")


class FlashProperties(ParameterSet):
    """FlashProperties ParameterSet."""
    Base = PropertyDescriptor("Base", r"""경로의 Base""")
    Qulaity = PropertyDescriptor("Qulaity", r"""재생 품질""")
    Scale = PropertyDescriptor("Scale", r"""스케일""")
    WMode = PropertyDescriptor("WMode", r"""윈도우 모드""")
    AutoPlay = PropertyDescriptor("AutoPlay", r"""자동 재생 여부 : 0 \= off, 1 \= on""")
    LoopPlay = PropertyDescriptor("LoopPlay", r"""반복 재생 여부 : 0 \= off, 1 \= on""")
    ShowMenu = PropertyDescriptor("ShowMenu", r"""메뉴 보이기 : 0 \= Hide, 1 \= Show""")
    BgColor = PropertyDescriptor("BgColor", r"""배경색 (COLORREF)""")


class FtpDownload(ParameterSet):
    """FtpDownload ParameterSet."""
    Server = PropertyDescriptor("Server", r"""Ftp 서버 내임""")
    UserName = PropertyDescriptor("UserName", r"""사용자 이름""")
    Password = PropertyDescriptor("Password", r"""사용자 패스워드""")
    Directory = PropertyDescriptor("Directory", r"""디렉터리""")
    FileName = PropertyDescriptor("FileName", r"""파일 명""")
    SaveType = PropertyDescriptor("SaveType", r"""저장할 포맷. 0 \= HTML 1 \= HWP 2: OOXML""")


class FtpUpload(ParameterSet):
    """FtpUpload ParameterSet."""
    Server = PropertyDescriptor("Server", r"""Ftp 서버 내임""")
    UserName = PropertyDescriptor("UserName", r"""사용자 이름""")
    Password = PropertyDescriptor("Password", r"""사용자 패스워드""")
    Directory = PropertyDescriptor("Directory", r"""디렉터리""")
    FileName = PropertyDescriptor("FileName", r"""파일 명""")
    SiteName = PropertyDescriptor("SiteName", r"""사이트 이름""")
    SaveType = PropertyDescriptor("SaveType", r"""저장할 포맷. 0 \= HTML 1 \= HWP""")


class HyperLink(ParameterSet):
    """HyperLink ParameterSet."""
    Text = PropertyDescriptor("Text", r"""하이퍼링크가 표시되는 문자열""")
    Command = PropertyDescriptor("Command", r"""Command String 참조""")
    NoLInk = PropertyDescriptor("NoLInk", r"""연결 안함 여부""")
    ShapeObject = PropertyDescriptor("ShapeObject", r"""그림 및 그리기객체가 Selection되어 있는지 여부""")
    DirectInsert = PropertyDescriptor("DirectInsert", r"""현재 캐럿 위치에 무조건 하이퍼링크 삽입 여부 (블록지정 상태면 블록해제 후 삽입) (한글2007에 새로 추가)""")


class HyperlinkJump(ParameterSet):
    """HyperlinkJump ParameterSet - 하이퍼링크 이동."""
    Text = PropertyDescriptor("Text", r"""하이퍼링크가 표시되는 문자열""")
    Command = PropertyDescriptor("Command", r"""Command String""")


class Idiom(ParameterSet):
    """Idiom ParameterSet."""
    InputText = PropertyDescriptor("InputText", r"""삽입될 스트링/끼워 넣을 파일""")
    InputType = PropertyDescriptor("InputType", r"""입력기 상용구/한글 상용구""")


class InputDateStyle(ParameterSet):
    """InputDateStyle ParameterSet."""
    DateStyleType = PropertyDescriptor("DateStyleType", r"""문자열로 넣기/코드로 넣기""")
    DateStyleDataForm = PropertyDescriptor("DateStyleDataForm", r"""필드 컨트롤의 안내문/지시문""")


class InsertFieldTemplate(ParameterSet):
    """InsertFieldTemplate ParameterSet."""
    ShowSingle = PropertyDescriptor("ShowSingle", r"""문서마당 정보 대화상자에서 페이지(탭) 보이기 옵션 :\-1 \= 모든 페이지 보이기0 \= 누름틀 페이지만 보이기1 \= 개인 정보 페이지만 보이기2 \= 문서 요약 페이지만 보이기3 \= 만든 날짜 페이지만 보이기4 \= 파일 경로 페이지만 보이기""")
    TemplateDirection = PropertyDescriptor("TemplateDirection", r"""필드 컨트롤의 안내문/지시문""")
    TemplateHelp = PropertyDescriptor("TemplateHelp", r"""필드 컨트롤의 도움말""")
    TemplateName = PropertyDescriptor("TemplateName", r"""필드 이름 (name)""")
    TemplateType = PropertyDescriptor("TemplateType", r"""필드의 종류.TemplateDirection/Help/Name의 값이 실제로 적용되는 위치 :0 \= 누름틀, 1 \= 개인 정보, 2 \= 문서 요약, 3 \= 만든 날짜, 4 \= 파일 경로""")
    Editable = PropertyDescriptor("Editable", r"""필드의 양식모드에서 편집여부 (한글2007에 새로 추가)0 \= 편집 불가능1 \= 편집 가능""")


class Internet(ParameterSet):
    """Internet ParameterSet."""
    OpenUrlWhere = PropertyDescriptor("OpenUrlWhere", r"""웹브라우저로 불러오거나 한글 문서로 불러오기""")
    OpenUrlString = PropertyDescriptor("OpenUrlString", r"""불러올 문서가 존재하는 URL""")


class KeyMacro(ParameterSet):
    """KeyMacro ParameterSet."""
    Index = PropertyDescriptor("Index", r"""정의(or 실행)할 매크로의 인덱스.""")
    RepeatCount = PropertyDescriptor("RepeatCount", r"""실행 반복 횟수""")
    Name = PropertyDescriptor("Name", r"""매크로 이름""")


class LinkDocument(ParameterSet):
    """LinkDocument ParameterSet."""
    Name = PropertyDescriptor("Name", r"""연결 문서 FILE NAME""")
    PageInherit = PropertyDescriptor("PageInherit", r"""TRUE \= 쪽 번호 잇기, FALSE \= 쪽 번호 잇지 않기.""")
    FootnoteInherit = PropertyDescriptor("FootnoteInherit", r"""TRUE \= 각주 번호 잇기, FALSE \= 각주 번호 잇지 않기.""")


class MailMergeGenerate(ParameterSet):
    """MailMergeGenerate ParameterSet."""
    Input = PropertyDescriptor("Input", r"""자료 종류 0 \= WAB1 \= OAB2 \= HWP3 \= DBF""")
    HwpPath = PropertyDescriptor("HwpPath", r"""Hwp 문서 경로.""")
    HwpId = PropertyDescriptor("HwpId", r"""Hwp 문서 ID""")
    DbfPath = PropertyDescriptor("DbfPath", r"""dbf file path""")
    DbfCode = PropertyDescriptor("DbfCode", r"""dbf file codepage0 \= KS1 \= KSSM2 \= GB3 \= BIG54 \= SJIS""")
    Output = PropertyDescriptor("Output", r"""출력 방향0 \= PRINTER1 \= PREVIEW2 \= FILE3 \= MAIL""")
    FileName = PropertyDescriptor("FileName", r"""파일 이름""")
    Continue = PropertyDescriptor("Continue", r"""쪽번호 잇기""")
    PrintSet = PropertyDescriptor("PrintSet", r"""인쇄 선택 사항""")
    Subject = PropertyDescriptor("Subject", r"""메일 제목""")
    Type = PropertyDescriptor("Type", r"""메일 종류0 \= 본문1 \= 첨부파일""")
    Field = PropertyDescriptor("Field", r"""메일 주소 필드""")
    FieldUpdate = PropertyDescriptor("FieldUpdate", r"""필드 단위 업데이트 (한글2007에 새로 추가)""")
    NxlPath = PropertyDescriptor("NxlPath", r"""넥셀 파일 경로 (한글2007에 새로 추가)""")


class MemoShape(ParameterSet):
    """MemoShape ParameterSet."""
    Width = PropertyDescriptor("Width", r"""너비 (HWPUNIT)""")
    LineType = PropertyDescriptor("LineType", r"""선 종류""")
    LineWidth = PropertyDescriptor("LineWidth", r"""선 굵기""")
    LineColor = PropertyDescriptor("LineColor", r"""선 색깔 (COLORREF)""")
    FillColor = PropertyDescriptor("FillColor", r"""채우기 색깔 (COLORREF)""")
    ActiveFillColor = PropertyDescriptor("ActiveFillColor", r"""활성화된 채우기 색깔 (COLORREF)""")
    MemoType = PropertyDescriptor("MemoType", r"""메모 종류 \- 현재 사용안 함1 \= 메모 넣기, 2 \= 메모 지우기, 3 \= 메모 고치기""")


class MousePos(ParameterSet):
    """MousePos ParameterSet."""
    XRelTo = PropertyDescriptor("XRelTo", r"""가로 상대적 기준0 : 종이1 : 쪽""")
    YRelTo = PropertyDescriptor("YRelTo", r"""세로 상대적 기준0 : 종이1 : 쪽""")
    Page = PropertyDescriptor("Page", r"""페이지 번호 ( 0 based)""")
    X = PropertyDescriptor("X", r"""가로 클릭한 위치 (HWPUNIT)""")
    Y = PropertyDescriptor("Y", r"""세로 클릭한 위치 (HWPUNIT)""")


class MovieProperties(ParameterSet):
    """MovieProperties ParameterSet."""
    Base = PropertyDescriptor("Base", r"""동영상 파일의 경로""")
    Caption = PropertyDescriptor("Caption", r"""자막 파일의 경로""")
    AutoPlay = PropertyDescriptor("AutoPlay", r"""자동 재생 여부 : 0 \= off, 1 \= on""")
    AutoRewind = PropertyDescriptor("AutoRewind", r"""되감기 여부 : 0 \= off, 1 \= on""")
    ShowMenu = PropertyDescriptor("ShowMenu", r"""메뉴 보이기 : 0 \= Hide, 1 \= Show""")
    ShowCtrlPanel = PropertyDescriptor("ShowCtrlPanel", r"""컨트롤 패널 보이기 : 0 \= Hide, 1 \= Show""")
    ShowPosCtrl = PropertyDescriptor("ShowPosCtrl", r"""위치 컨트롤 보이기 : 0 \= Hide, 1 \= Show""")
    EnablePos = PropertyDescriptor("EnablePos", r"""위치 컨트롤 조절 여부 : 0 \= Disable, 1 \= Enable""")
    ShowTrackBar = PropertyDescriptor("ShowTrackBar", r"""음량 조절막대(Track Bar) 보이기 : 0 \= Hide, 1 \= Show""")
    EnableTrack = PropertyDescriptor("EnableTrack", r"""음량 조절 여부 : 0 \= Disable, 1 \= Enable""")
    ShowChaption = PropertyDescriptor("ShowChaption", r"""자막 보이기 : 0 \= Hide, 1 \= Show""")
    ShowAudio = PropertyDescriptor("ShowAudio", r"""오디오 설정 보이기 : 0 \= Hide, 1 \= Show""")
    ShowStatus = PropertyDescriptor("ShowStatus", r"""상태바 보기 (진행시간 등을 표시) : 0 \= Hide, 1 \= Show""")


class OleCreation(ParameterSet):
    """OleCreation ParameterSet."""
    Type = PropertyDescriptor("Type", r"""생성 방식0 \= 새로 개체 생성1 \= 파일로부터 개체 생성2 \= 파일로 링크된 개체 생성""")
    Clsid = PropertyDescriptor("Clsid", r"""클래스 ID (새로 개체 생성‘일 때 사용)""")
    Path = PropertyDescriptor("Path", r"""파일 경로 (‘파일로 링크된 개체 생성’, ‘파일로부터 개체 생성’일 때 사용)""")
    Aspect = PropertyDescriptor("Aspect", r"""생성된 OLE 개체를 아이콘으로 표시할지 여부 :0 \= 내용으로 표시, 1 \= 아이콘으로 표시""")
    IconMetafile = PropertyDescriptor("IconMetafile", r"""Aspect가 아이콘일 때 적용할 아이콘 데이터""")
    IconMM = PropertyDescriptor("IconMM", r"""Aspect가 아이콘일 때 아이콘 매핑모드1 \= MM_TEXT2 \= MM_LOMETRIC3 \= MM_HIMETRIC4 \= MM_LOENGLISH5 \= MM_HIENGLISH6 \= MM_TWIPS7 \= MM_ISOTROPIC8 \= MM_ANISOTROPIC※ MFC의 매핑모드와 값/의미가 동일하다.""")
    IconXext = PropertyDescriptor("IconXext", r"""Aspect가 아이콘일 때 X축 매핑단위""")
    IconYext = PropertyDescriptor("IconYext", r"""Aspect가 아이콘일 때 Y축 매핑단위""")
    InnerOCX = PropertyDescriptor("InnerOCX", r"""한글 내부에서 사용되는 OCX인지 여부 (예: 한글의 Chart OLE)""")
    SoProperties = PropertyDescriptor("SoProperties", r"""초기 shape object 속성""")
    FlashProperties = PropertyDescriptor("FlashProperties", r"""플래시 파일 삽입 시 필요한 옵션 셋 (한글2007에 새로 추가)""")
    MovieProperties = PropertyDescriptor("MovieProperties", r"""동영상 파일 삽입 시 필요한 옵션 셋 (한글2007에 새로 추가)""")


class ScriptMacro(ParameterSet):
    """ScriptMacro ParameterSet."""
    Index = PropertyDescriptor("Index", r"""정의(or 실행)할 매크로의 인덱스""")
    RepeatCount = PropertyDescriptor("RepeatCount", r"""실행 반복 횟수""")
    Name = PropertyDescriptor("Name", r"""매크로 이름""")
    Detail = PropertyDescriptor("Detail", r"""매크로 설명""")


class Sort(ParameterSet):
    """Sort ParameterSet."""
    KeyOption = PropertyDescriptor("KeyOption", r"""키 콤보에서 선택된 키를 저장함.""")
    CheckJasoReverse = PropertyDescriptor("CheckJasoReverse", r"""자소 단위 비교 Flag \- 종, 중, 초""")
    DelimiterType = PropertyDescriptor("DelimiterType", r"""필드 구분 기호 형식 : 0 \= 탭(Tab), 1 \= 콤마(,), 2 \= 빈칸(Space), 3 \= 사용자 정의""")
    DelimiterChars = PropertyDescriptor("DelimiterChars", r"""필드 구분 기호들. DelimiterType이 3(사용자 정의)일 경우에만 유효""")
    IgnoreMultiDelimiter = PropertyDescriptor("IgnoreMultiDelimiter", r"""연속되는 구분기호 무시 Flag""")
    CheckFromRear = PropertyDescriptor("CheckFromRear", r"""단어 뒤에서 부터 비교 Flag""")
    CheckExtendYear = PropertyDescriptor("CheckExtendYear", r"""두 자리 년도 확장 check Flag""")
    YearBase = PropertyDescriptor("YearBase", r"""두 자리 년도 시작 년도""")
    LangOrderType = PropertyDescriptor("LangOrderType", r"""사전언어순서 값""")
    CheckJaso = PropertyDescriptor("CheckJaso", r"""자소 단위 비교 Flag \- 초, 중, 종""")


class Sum(ParameterSet):
    """Sum ParameterSet."""
    Sum = PropertyDescriptor("Sum", r"""Parameter property""")
    Average = PropertyDescriptor("Average", r"""Parameter property""")
    LineCount = PropertyDescriptor("LineCount", r"""줄 수""")
    Comma = PropertyDescriptor("Comma", r"""세 자리마다 쉼표로 자리 구분 (on / off)""")
    Option = PropertyDescriptor("Option", r"""형식 옵션""")


class UserQCommandFile(ParameterSet):
    """UserQCommandFile ParameterSet."""
    Save = PropertyDescriptor("Save", r"""저장 (TRUE \= Save, FALSE \= Open)""")
    FileName = PropertyDescriptor("FileName", r"""파일명""")
    LoadType = PropertyDescriptor("LoadType", r"""로드 방법 (TRUE \= Overwrite, FALSE \= Merge)""")


class ViewProperties(ParameterSet):
    """ViewProperties ParameterSet."""
    OptionFlag = PropertyDescriptor("OptionFlag", r"""뷰 옵션 플랙. 여러 개를 OR연산하여 지정할 수 있음.0x00000001 \= off : 쪽윤곽, on : 기본 보기0x00000002 \= 공백과 폭이 없는 컨트롤을 기호로0x00000004 \= 문단 마크 기호로0x00000008 \= 안내선0x00000010 \= 그리기 격자0x00000020 \= 그림 감춤0x00010000 \= 회색조""")
    ZoomType = PropertyDescriptor("ZoomType", r"""화면 확대 종류. 0 \= 사용자 정의 1 \= 쪽 맞춤 2 \= 폭 맞춤 3 \= 여러 쪽""")
    ZoomRatio = PropertyDescriptor("ZoomRatio", r"""화면 확대 종류가 "사용자 정의"인 경우 화면 확대 비율. 10% \~ 500%""")
    ZoomCntX = PropertyDescriptor("ZoomCntX", r"""화면 확대 종류가 "여러 쪽"인 경우 가로 페이지 수. 1 \~ 8""")
    ZoomCntY = PropertyDescriptor("ZoomCntY", r"""화면 확대 종류가 "여러 쪽"인 경우 세로 페이지 수.1 \~ 8""")
    ZoomMirror = PropertyDescriptor("ZoomMirror", r"""맞쪽 보기. 페이지 수가 2의 배수일 때만 동작(한글2007에 새로 추가)""")


class ViewStatus(ParameterSet):
    """ViewStatus ParameterSet."""
    Type = PropertyDescriptor("Type", r"""0 (현재 View의 절대 Pos값만 지원함)""")
    ViewPosX = PropertyDescriptor("ViewPosX", r"""현재 뷰의 X값""")
    ViewPosY = PropertyDescriptor("ViewPosY", r"""현재 뷰의 Y값""")


# ── Runtime-discovered ParameterSet classes ──────────────────────────────
# These SetIDs exist in HWP runtime but are not in official documentation.


class Label(ParameterSet):
    """Label ParameterSet - 라벨 문서."""

