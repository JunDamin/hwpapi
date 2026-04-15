"""
File / print operation ParameterSet classes (14 classes).

Covers file I/O and printing:
- FileOpen, FileSaveAs, FileSaveBlock, FileConvert, FileOpenSave, InsertFile
- FileSendMail, FileSetSecurity
- Print, PrintToImage, PrintWatermark
- Preference, Presentation, EngineProperties
"""
from __future__ import annotations
from hwpapi.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)


class EngineProperties(ParameterSet):
    """EngineProperties ParameterSet."""
    DoAnyCursorEdit = PropertyDescriptor("DoAnyCursorEdit", r"""마우스로 두 번 누르기 한곳에 입력가능""")
    DoOutLineStyle = PropertyDescriptor("DoOutLineStyle", r"""개요 번호 삽입 문단에 개요 스타일 자동 적용""")
    InsertLock = PropertyDescriptor("InsertLock", r"""삽입 잠금""")
    TabMoveCell = PropertyDescriptor("TabMoveCell", r"""표 안에서 \<Tab\>으로 셀 이동""")
    FaxDriver = PropertyDescriptor("FaxDriver", r"""팩스 드라이버""")
    PDFDriver = PropertyDescriptor("PDFDriver", r"""PDF 드라이버""")
    EnableAutoSpell = PropertyDescriptor("EnableAutoSpell", r"""맞춤법 도우미 작동""")


class FileConvert(ParameterSet):
    """FileConvert ParameterSet."""
    Format = PropertyDescriptor("Format", r"""변환 포맷""")
    SrcFileList = PropertyDescriptor("SrcFileList", r"""Source 파일 리스트""")
    DestFileList = PropertyDescriptor("DestFileList", r"""Destination 파일 리스트""")


class FileOpen(ParameterSet):
    """FileOpen ParameterSet."""
    OpenFileName = PropertyDescriptor("OpenFileName", r"""파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다)""")
    OpenFormat = PropertyDescriptor("OpenFormat", r"""파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭)""")
    OpenReadOnly = PropertyDescriptor("OpenReadOnly", r"""읽기 전용""")
    OpenFlag = PropertyDescriptor("OpenFlag", r"""옵션 0x00 \= 새 창으로 열기0x01 \= 현재 창의 새 탭에 열기0x02 \= 현재 창의 현재 탭에 열기0x03 \= 위 세 값의 mask0x04 \=이미 열려진 문서일 때 다시 load할지 뭍을 것인지0x08 \= 초기 모드를 최근 작업 문서 상태로0x10 \= 문서 마당""")
    SaveOverWrite = PropertyDescriptor("SaveOverWrite", r"""덮어 쓰기""")
    ModifiedFlag = PropertyDescriptor("ModifiedFlag", r"""Modify 플래그""")
    Argument = PropertyDescriptor("Argument", r"""Argument""")
    SaveCMFDoc30 = PropertyDescriptor("SaveCMFDoc30", r"""97 호환 저장 (한글2007에 새로 추가)""")
    SaveHwp97 = PropertyDescriptor("SaveHwp97", r"""97 문서로 저장 (한글2007에 새로 추가)""")
    SaveDistribute = PropertyDescriptor("SaveDistribute", r"""배포용 문서로 저장 (한글2007에 새로 추가)""")
    SaveDRMDoc = PropertyDescriptor("SaveDRMDoc", r"""보안 문서로 저장 (한글2007에 새로 추가)""")


class FileSaveAs(ParameterSet):
    """FileSaveAs ParameterSet."""
    OpenFileName = PropertyDescriptor("OpenFileName", r"""파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다)""")
    OpenFormat = PropertyDescriptor("OpenFormat", r"""파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭)""")
    OpenReadOnly = PropertyDescriptor("OpenReadOnly", r"""읽기 전용""")
    OpenFlag = PropertyDescriptor("OpenFlag", r"""옵션 0x00 \= 새 창으로 열기0x01 \= 현재 창의 새 탭에 열기0x02 \= 현재 창의 현재 탭에 열기0x03 \= 위 세 값의 mask0x04 \=이미 열려진 문서일 때 다시 load할지 뭍을 것인지0x08 \= 초기 모드를 최근 작업 문서 상태로0x10 \= 문서 마당""")
    SaveOverWrite = PropertyDescriptor("SaveOverWrite", r"""덮어 쓰기""")
    ModifiedFlag = PropertyDescriptor("ModifiedFlag", r"""Modify 플래그""")
    Argument = PropertyDescriptor("Argument", r"""Argument""")
    SaveCMFDoc30 = PropertyDescriptor("SaveCMFDoc30", r"""97 호환 저장""")
    SaveHwp97 = PropertyDescriptor("SaveHwp97", r"""97 문서로 저장""")
    SaveDistribute = PropertyDescriptor("SaveDistribute", r"""배포용 문서로 저장""")
    SaveDRMDoc = PropertyDescriptor("SaveDRMDoc", r"""보안 문서로 저장""")


class FileSaveBlock(ParameterSet):
    """FileSaveBlock ParameterSet."""
    FileName = PropertyDescriptor("FileName", r"""파일 이름""")
    Format = PropertyDescriptor("Format", r"""파일 포맷""")
    Argument = PropertyDescriptor("Argument", r"""argument""")


class FileSendMail(ParameterSet):
    """FileSendMail ParameterSet."""
    To = PropertyDescriptor("To", r"""받는 사람""")
    Type = PropertyDescriptor("Type", r"""메일 보내기 유형: 0 \= 본문, 1 \= 첨부파일""")
    Subject = PropertyDescriptor("Subject", r"""Parameter property""")
    Filepath = PropertyDescriptor("Filepath", r"""사용자가 직접 입력한 파일 (이 아이템은 Type 아이템이 1(첨부파일)로 설정되어 있을 때만 유효하다) (한글2007에 새로 추가)""")


class FileSetSecurity(ParameterSet):
    """FileSetSecurity ParameterSet."""
    Password = PropertyDescriptor("Password", r"""배포용 문서 암호(7자리 이상 가능)""")
    NoPrint = PropertyDescriptor("NoPrint", r"""프린트 가능한 배포용 문서를 만들지의 여부0 : 프린트 가능1 : 프린트 가능하지 않음""")
    NoCopy = PropertyDescriptor("NoCopy", r"""문서 내용 복사가 가능한 배포용 문서를 만들지의 여부0 : 복사 가능1 : 복사 가능하지 않음""")


class InsertFile(ParameterSet):
    """InsertFile ParameterSet."""
    FileName = PropertyDescriptor("FileName", r"""삽입할 파일의 이름""")
    FileFormat = PropertyDescriptor("FileFormat", r"""삽입할 파일의 확장자""")
    FileArg = PropertyDescriptor("FileArg", r"""삽입할 파일의 Argument""")
    KeepSection = PropertyDescriptor("KeepSection", r"""끼워 넣을 문서를 구역으로 나누어 쪽 모양을 유지할지 여부 on / off (ver:0x0605010E)""")
    KeepCharshape = PropertyDescriptor("KeepCharshape", r"""끼워 넣을 문서의 글자 모양을 유지할지 여부 on / off(한글2007에 새로 추가)""")
    KeepParashape = PropertyDescriptor("KeepParashape", r"""끼워 넣을 문서의 문단 모양을 유지할지 여부 on / off(한글2007에 새로 추가)""")
    KeepStyle = PropertyDescriptor("KeepStyle", r"""끼워 넣을 문서의 스타일을 유지할지 여부 on / off(한글2010에 새로 추가)""")


class Preference(ParameterSet):
    """Preference ParameterSet."""
    ShowSinglePage = PropertyDescriptor("ShowSinglePage", r"""환경 설정 PropertySheet에 표시할 PropertyPage 번호(하나만 선택)""")
    ApplyLinkAttr = PropertyDescriptor("ApplyLinkAttr", r"""하이퍼링크 글자 속성 문서 전체에 적용하기 여부 (on/off)(한글2007에 새로 추가)""")
    ApplyForbidden = PropertyDescriptor("ApplyForbidden", r"""(금칙 처리) 새 문서에 기본 값으로 설정 (on/off)(한글2007에 새로 추가)""")
    StartForbiddenStr = PropertyDescriptor("StartForbiddenStr", r"""(금칙 처리) 새 문서에 적용할 줄 앞 금칙 문자열(한글2007에 새로 추가)""")
    EndForbiddenStr = PropertyDescriptor("EndForbiddenStr", r"""(금칙 처리) 새 문서에 적용할 줄 뒤 금칙 문자열(한글2007에 새로 추가)""")


class Presentation(ParameterSet):
    """Presentation ParameterSet."""
    DialogResult = PropertyDescriptor("DialogResult", r"""프리젠테이션 대화상자의 "실행"버튼이 클릭되었는지 여부.한글2007에서는 이 Set Item이 제거되었다.""")
    Effect = PropertyDescriptor("Effect", r"""화면 전환 효과""")
    Sound = PropertyDescriptor("Sound", r"""효과음""")
    InvertText = PropertyDescriptor("InvertText", r"""검은색 글자를 흰색으로""")
    ShowMode = PropertyDescriptor("ShowMode", r"""자동 전환 모드 (한글2007에 새로 추가)""")
    ShowPage = PropertyDescriptor("ShowPage", r"""현재 쪽""")
    ShowTime = PropertyDescriptor("ShowTime", r"""전환 시간 (초)""")


class Print(ParameterSet):
    """Print ParameterSet."""
    DialogOption = PropertyDescriptor("DialogOption", r"""미리보기 버튼을 보일지 말지를 지정.한글2007에서는 이 Set Item이 제거되었다.""")
    DialogResult = PropertyDescriptor("DialogResult", r"""인쇄 대화상자에서 어떤 버튼이 눌러졌는지 알려준다.한글2007에서는 이 Set Item이 제거되었다.1 \= 인쇄, 2 \= (미리보기), 3 \= 팩스로 인쇄, 그 외 \= 취소""")
    Range = PropertyDescriptor("Range", r"""인쇄 범위0 \= 문서전체 (연결된 문서 포함)1 \= 현재 쪽2 \= 현재부터3 \= 현재까지4 \= 일부분5 \= 선택한 쪽만6 \= 현재 문서 (연결된 문서 미포함)7 \= 현재 구역""")
    RangeCustom = PropertyDescriptor("RangeCustom", r"""사용자가 직접 입력한 인쇄 범위""")
    RangeIncludeLinkedDoc = PropertyDescriptor("RangeIncludeLinkedDoc", r"""연결된 문서 포함""")
    NumCopy = PropertyDescriptor("NumCopy", r"""인쇄 매수""")
    Collate = PropertyDescriptor("Collate", r"""한 부씩 찍기""")
    PrintMethod = PropertyDescriptor("PrintMethod", r"""인쇄 방법0 \= 자동 인쇄1 \= 공급 용지에 맞추어2 \= 나눠 찍기3 \= 자동으로 모아 찍기4 \= 2쪽씩 모아 찍기5 \= 3쪽씩 모아 찍기6 \= 4쪽씩 모아 찍기7 \= 6쪽씩 모아 찍기8 \= 8쪽씩 모아 찍기9 \= 9쪽씩 모아 찍기10 \= 16쪽씩 모아 찍기""")
    PrinterPaperSize = PropertyDescriptor("PrinterPaperSize", r"""공급용지 종류(DEVMODE.dmPaperSize)""")
    PrinterPaperWidth = PropertyDescriptor("PrinterPaperWidth", r"""공급용지 종류(DEVMODE.dmPaperWidth)""")
    PrinterPaperLength = PropertyDescriptor("PrinterPaperLength", r"""공급용지 종류(DEVMODE.dmPaperLength)""")
    PrintAutoHeadNote = PropertyDescriptor("PrintAutoHeadNote", r"""머리말 자동 인쇄""")
    PrintAutoFootNote = PropertyDescriptor("PrintAutoFootNote", r"""꼬리말 자동 인쇄""")
    PrintAutoHeadnoteLtext = PropertyDescriptor("PrintAutoHeadnoteLtext", r"""자동 머리말의 왼쪽 String""")
    PrintAutoHeadnoteCtext = PropertyDescriptor("PrintAutoHeadnoteCtext", r"""자동 머리말의 가운데 String""")
    PrintAutoHeadnoteRtext = PropertyDescriptor("PrintAutoHeadnoteRtext", r"""자동 머리말의 오른쪽 String""")
    PrintAutoFootnoteLtext = PropertyDescriptor("PrintAutoFootnoteLtext", r"""자동 꼬리말의 왼쪽 String""")
    PrintAutoFootnoteCtext = PropertyDescriptor("PrintAutoFootnoteCtext", r"""자동 꼬리말의 가운데 String""")
    PrintAutoFootnoteRtext = PropertyDescriptor("PrintAutoFootnoteRtext", r"""자동 꼬리말의 오른쪽 String""")
    PrinterName = PropertyDescriptor("PrinterName", r"""프린터""")
    PrintToFile = PropertyDescriptor("PrintToFile", r"""인쇄 결과를 파일로 저장""")
    FileName = PropertyDescriptor("FileName", r"""인쇄 결과를 저장할 파일 이름""")
    ReverseOrder = PropertyDescriptor("ReverseOrder", r"""역순 인쇄""")
    Pause = PropertyDescriptor("Pause", r"""끊어 찍기 매수""")
    PrintImage = PropertyDescriptor("PrintImage", r"""그림 개체""")
    PrintDrawObj = PropertyDescriptor("PrintDrawObj", r"""그리기 개체""")
    PrintClickHere = PropertyDescriptor("PrintClickHere", r"""누름틀""")
    PrintCropMark = PropertyDescriptor("PrintCropMark", r"""편집 용지 표시""")
    IdcPrintWallPaper = PropertyDescriptor("IdcPrintWallPaper", r"""배경 그림""")
    LastBlankPage = PropertyDescriptor("LastBlankPage", r"""빈 마지막 쪽""")
    BinderHoleType = PropertyDescriptor("BinderHoleType", r"""바인더 구멍""")
    EvenOddPageType = PropertyDescriptor("EvenOddPageType", r"""홀짝 인쇄""")
    ZoomX = PropertyDescriptor("ZoomX", r"""가로 확대""")
    ZoomY = PropertyDescriptor("ZoomY", r"""세로 확대""")
    StartPageNum = PropertyDescriptor("StartPageNum", r"""시작 번호/쪽 번호""")
    StartFootNoteNum = PropertyDescriptor("StartFootNoteNum", r"""시작 번호/각주 번호""")
    Flags = PropertyDescriptor("Flags", r"""문제 해결을 위한 고급 선택 사항""")
    Device = PropertyDescriptor("Device", r"""인쇄 방향(장치)0 : 프린터1: 팩스2: 그림으로 저장3: PDF 파일로 저장4: 미리보기""")
    PrintFormObj = PropertyDescriptor("PrintFormObj", r"""양식 개체 출력여부 (한글2007에 새로 추가)""")
    PrintMarkPen = PropertyDescriptor("PrintMarkPen", r"""형광펜 출력여부 (한글2007에 새로 추가)""")
    PrintMemo = PropertyDescriptor("PrintMemo", r"""메모 출력여부 (한글2007에 새로 추가)""")
    PrintMemoContents = PropertyDescriptor("PrintMemoContents", r"""메모 내용 출력여부 (한글2007에 새로 추가)""")
    PrintRevision = PropertyDescriptor("PrintRevision", r"""교정부호 출력여부 (한글2007에 새로 추가)""")
    PrintWatermark = PropertyDescriptor("PrintWatermark", r"""인쇄워터마크 (한글2007에 새로 추가)""")


class PrintToImage(ParameterSet):
    """PrintToImage ParameterSet."""
    Format = PropertyDescriptor("Format", r"""그림 형식0 : none1 : BMP2 : GIF3 : PNG4 : JPG5 : WMF""")
    FileName = PropertyDescriptor("FileName", r"""그림 경로""")
    ColorDepth = ColorProperty("ColorDepth", r"""색상수 (bits: 8, 16\...)""")
    Resolution = PropertyDescriptor("Resolution", r"""해상도""")


class PrintWatermark(ParameterSet):
    """PrintWatermark ParameterSet."""
    WatermarkType = PropertyDescriptor("WatermarkType", r"""현재 선택된 워터마크의 유형을 나타냄.0 \= 워터마크 없음1 \= 그림 워터마크2 \= 글자 워터마크""")
    PosPage = PropertyDescriptor("PosPage", r"""워터마크의 위치 기준 : 0 \= 종이 기준, 1 \= 쪽 기준""")
    TextWrap = PropertyDescriptor("TextWrap", r"""워터마크의 배치 : 0 \= 글 뒤로, 1 \= 글 앞으로""")
    AlphaText = PropertyDescriptor("AlphaText", r"""글자 투명도 (0 \~ 255\)""")
    AlphaImage = PropertyDescriptor("AlphaImage", r"""그림 투명도 (0 \~ 255\)""")
    FileName = PropertyDescriptor("FileName", r"""그림 파일의 경로 or 그림파일 삽입일 경우에는 binary data""")
    PicEffect = PropertyDescriptor("PicEffect", r"""그림 효과 :0 \= 실제 이미지 그대로, 1 \= 그레이스케일, 2 \= 흑백효과""")
    Brightness = PropertyDescriptor("Brightness", r"""명도 (\-100 \~ 100\)""")
    Contrast = PropertyDescriptor("Contrast", r"""밝기 (\-100 \~ 100\)""")
    DrawFillImageType = PropertyDescriptor("DrawFillImageType", r"""채우기 유형0 \= 바둑판식으로 \- 모두1 \= 바둑판식으로 \- 가로/위2 \= 바둑판식으로 \- 가로/아래3 \= 바둑판식으로 \- 세로/왼쪽4 \= 바둑판식으로 \- 세로/오른쪽5 \= 크기에 맞추어6 \= 가운데로7 \= 가운데 위로 8 \= 가운데 아래로9 \= 왼쪽 가운데로10 \= 왼쪽 위로11 \= 왼쪽 아래로12 \= 오른쪽 가운데로13 \= 오른쪽 위로14 \= 오른쪽 아래로15 \= 원래크기에 비례하여""")
    String = PropertyDescriptor("String", r"""글맵시에 넣을 문자열 내용 : 내용""")
    FontName = PropertyDescriptor("FontName", r"""Parameter property""")
    FontType = PropertyDescriptor("FontType", r"""글꼴 속성 : 0 \= dont care, 1 \= TTF, 2 \= HFT""")
    FontSize = PropertyDescriptor("FontSize", r"""글꼴 크기 (HWPUNIT : 2500(25pt) \~ 25400(254pt)""")
    ShadowType = PropertyDescriptor("ShadowType", r"""그림자 종류 : 0 \= none, 1 \= drop, 2 \= continuous""")
    ShadowOffsetX = PropertyDescriptor("ShadowOffsetX", r"""X축 그림자 간격 (\-48% \~ 48% )""")
    ShadowOffsetY = PropertyDescriptor("ShadowOffsetY", r"""Y축 그림자 간격 (\-48% \~ 48% )""")
    ShadowColor = ColorProperty("ShadowColor", r"""그림자 색 (COLORREF)""")
    FontColor = ColorProperty("FontColor", r"""글자색 (COLORREF)""")
    RotateAngle = PropertyDescriptor("RotateAngle", r"""회전각도 (\-360 \~ 360\)""")
    WaterMarkEff = PropertyDescriptor("WaterMarkEff", r"""워터마크 효과 : 0 \= off, 1 \= on""")


class FileOpenSave(ParameterSet):
    """FileOpenSave ParameterSet - 파일 열기/저장."""

