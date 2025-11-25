"""
HWP Document Parser Module

이 모듈은 HWP 문서를 파싱하여 구조화된 데이터로 변환합니다.
"""

from typing import List, Dict, Any, Optional
from pathlib import Path
from .core import App, MaskOption, ScanStartPosition, ScanEndPosition
from .dataclasses import CharShape, ParaShape


class HwpParser:
    """
    HWP 문서를 파싱하여 구조화된 데이터로 변환하는 클래스

    Parameters
    ----------
    file_path : str
        파싱할 HWP 파일 경로
    visible : bool, optional
        HWP 프로그램을 화면에 표시할지 여부 (기본값: False)

    Examples
    --------
    >>> parser = HwpParser("document.hwp")
    >>> document = parser.parse()
    >>> print(document['metadata']['title'])
    """

    def __init__(self, file_path: str, visible: bool = False):
        """HWP 파서 초기화"""
        self.file_path = Path(file_path)
        self.app = App(new_app=True, is_visible=visible)
        self.document_data = None

        if not self.file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    def parse(self) -> Dict[str, Any]:
        """
        HWP 문서를 파싱하여 구조화된 딕셔너리로 반환

        Returns
        -------
        dict
            파싱된 문서 데이터
        """
        # 문서 열기
        self.app.open(str(self.file_path))

        # 문서 데이터 초기화
        self.document_data = {
            "metadata": self._extract_metadata(),
            "sections": self._extract_sections(),
            "fonts": self._extract_fonts(),
        }

        return self.document_data

    def _extract_metadata(self) -> Dict[str, Any]:
        """문서 메타데이터 추출"""
        doc = self.app.api.XHwpDocuments.Active_XHwpDocument

        metadata = {
            "file_path": str(self.file_path),
            "file_name": self.file_path.name,
            "full_path": doc.FullName if hasattr(doc, 'FullName') else str(self.file_path),
        }

        return metadata

    def _extract_fonts(self) -> List[str]:
        """문서에서 사용된 폰트 목록 추출"""
        try:
            fonts = self.app.get_font_list()
            return fonts if fonts else []
        except Exception:
            return []

    def _extract_sections(self) -> List[Dict[str, Any]]:
        """문서의 모든 섹션 추출"""
        sections = []

        # 문서 처음으로 이동
        from .core import MoveId
        self.app.move(key=MoveId.TopOfFile)

        # 전체 문서 스캔
        paragraphs = self._extract_paragraphs()

        # 기본적으로 하나의 섹션으로 구성
        sections.append({
            "section_number": 0,
            "paragraphs": paragraphs
        })

        return sections

    def _extract_paragraphs(self) -> List[Dict[str, Any]]:
        """모든 문단 추출"""
        paragraphs = []
        para_index = 0

        # 문서의 각 문단을 순회
        with self.app.scan(
            option=MaskOption.All,
            scan_spos=ScanStartPosition.Document,
            scan_epos=ScanEndPosition.Document
        ) as scan:
            current_para_text = []

            for text in scan:
                if text:  # 텍스트가 있는 경우
                    # 현재 위치의 문자 모양 가져오기
                    try:
                        char_shape = self.app.get_charshape()
                        para_shape = self.app.get_parashape()

                        para_data = {
                            "paragraph_number": para_index,
                            "text": text,
                            "char_shape": char_shape.todict() if char_shape else {},
                            "para_shape": para_shape.todict() if para_shape else {},
                        }

                        paragraphs.append(para_data)
                        para_index += 1
                    except Exception as e:
                        # 오류 발생 시 기본 데이터만 저장
                        paragraphs.append({
                            "paragraph_number": para_index,
                            "text": text,
                            "char_shape": {},
                            "para_shape": {},
                        })
                        para_index += 1

        return paragraphs

    def close(self):
        """파서 종료 및 리소스 정리"""
        try:
            self.app.close()
        except Exception:
            pass

        try:
            self.app.quit()
        except Exception:
            pass

    def __enter__(self):
        """컨텍스트 매니저 진입"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """컨텍스트 매니저 종료"""
        self.close()


def parse_hwp(file_path: str, visible: bool = False) -> Dict[str, Any]:
    """
    HWP 파일을 파싱하는 편의 함수

    Parameters
    ----------
    file_path : str
        파싱할 HWP 파일 경로
    visible : bool, optional
        HWP 프로그램을 화면에 표시할지 여부 (기본값: False)

    Returns
    -------
    dict
        파싱된 문서 데이터

    Examples
    --------
    >>> data = parse_hwp("document.hwp")
    >>> print(data['metadata']['file_name'])
    """
    with HwpParser(file_path, visible=visible) as parser:
        return parser.parse()
