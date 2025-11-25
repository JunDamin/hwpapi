"""
JSON Exporter Module

HWP 파싱 데이터를 JSON 형식으로 내보내는 모듈
"""

import json
from typing import Dict, Any, Optional
from pathlib import Path


class JsonExporter:
    """
    HWP 파싱 데이터를 JSON 파일로 내보내는 클래스

    Parameters
    ----------
    data : dict
        내보낼 HWP 문서 데이터
    indent : int, optional
        JSON 들여쓰기 레벨 (기본값: 2)
    ensure_ascii : bool, optional
        ASCII만 사용할지 여부 (기본값: False, 한글 지원)

    Examples
    --------
    >>> from hwpapi.parser import parse_hwp
    >>> data = parse_hwp("document.hwp")
    >>> exporter = JsonExporter(data)
    >>> exporter.save("output.json")
    """

    def __init__(
        self,
        data: Dict[str, Any],
        indent: int = 2,
        ensure_ascii: bool = False
    ):
        """JSON Exporter 초기화"""
        self.data = data
        self.indent = indent
        self.ensure_ascii = ensure_ascii

    def to_json(self) -> str:
        """
        데이터를 JSON 문자열로 변환

        Returns
        -------
        str
            JSON 형식의 문자열
        """
        return json.dumps(
            self.data,
            indent=self.indent,
            ensure_ascii=self.ensure_ascii,
            default=str
        )

    def save(self, output_path: str) -> str:
        """
        JSON 데이터를 파일로 저장

        Parameters
        ----------
        output_path : str
            저장할 파일 경로

        Returns
        -------
        str
            저장된 파일의 절대 경로
        """
        output_file = Path(output_path)

        # 디렉토리가 없으면 생성
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # JSON 파일로 저장
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(
                self.data,
                f,
                indent=self.indent,
                ensure_ascii=self.ensure_ascii,
                default=str
            )

        return str(output_file.absolute())

    def pretty_print(self):
        """JSON 데이터를 콘솔에 예쁘게 출력"""
        print(self.to_json())


def export_to_json(
    data: Dict[str, Any],
    output_path: str,
    indent: int = 2,
    ensure_ascii: bool = False
) -> str:
    """
    HWP 파싱 데이터를 JSON 파일로 내보내는 편의 함수

    Parameters
    ----------
    data : dict
        내보낼 HWP 문서 데이터
    output_path : str
        저장할 파일 경로
    indent : int, optional
        JSON 들여쓰기 레벨 (기본값: 2)
    ensure_ascii : bool, optional
        ASCII만 사용할지 여부 (기본값: False, 한글 지원)

    Returns
    -------
    str
        저장된 파일의 절대 경로

    Examples
    --------
    >>> from hwpapi.parser import parse_hwp
    >>> data = parse_hwp("document.hwp")
    >>> export_to_json(data, "output.json")
    '/path/to/output.json'
    """
    exporter = JsonExporter(data, indent=indent, ensure_ascii=ensure_ascii)
    return exporter.save(output_path)


def hwp_to_json(
    hwp_path: str,
    json_path: Optional[str] = None,
    visible: bool = False,
    indent: int = 2,
    ensure_ascii: bool = False
) -> str:
    """
    HWP 파일을 직접 JSON 파일로 변환하는 올인원 함수

    Parameters
    ----------
    hwp_path : str
        변환할 HWP 파일 경로
    json_path : str, optional
        저장할 JSON 파일 경로 (기본값: hwp 파일명에 .json 확장자)
    visible : bool, optional
        HWP 프로그램을 화면에 표시할지 여부 (기본값: False)
    indent : int, optional
        JSON 들여쓰기 레벨 (기본값: 2)
    ensure_ascii : bool, optional
        ASCII만 사용할지 여부 (기본값: False, 한글 지원)

    Returns
    -------
    str
        저장된 JSON 파일의 절대 경로

    Examples
    --------
    >>> hwp_to_json("document.hwp")
    '/path/to/document.json'

    >>> hwp_to_json("document.hwp", "output.json")
    '/path/to/output.json'
    """
    from .parser import parse_hwp

    # JSON 파일 경로가 지정되지 않은 경우 hwp 파일명 사용
    if json_path is None:
        hwp_file = Path(hwp_path)
        json_path = hwp_file.with_suffix('.json')

    # HWP 파싱
    data = parse_hwp(hwp_path, visible=visible)

    # JSON으로 내보내기
    return export_to_json(data, str(json_path), indent=indent, ensure_ascii=ensure_ascii)
