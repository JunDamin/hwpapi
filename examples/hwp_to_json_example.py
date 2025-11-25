"""
HWP to JSON 변환 예제

이 스크립트는 HWP 파일을 JSON 형식으로 변환하는 다양한 방법을 보여줍니다.
"""

from pathlib import Path
from HwpApi import hwp_to_json, parse_hwp, export_to_json, HwpParser


def example1_simple_conversion():
    """
    예제 1: 가장 간단한 방법 - 한 줄로 변환
    """
    print("=" * 60)
    print("예제 1: 간단한 HWP to JSON 변환")
    print("=" * 60)

    # HWP 파일을 JSON으로 변환 (자동으로 같은 이름의 .json 파일 생성)
    result_path = hwp_to_json("document.hwp")
    print(f"JSON 파일이 생성되었습니다: {result_path}")
    print()


def example2_custom_output_path():
    """
    예제 2: 출력 경로를 지정하여 변환
    """
    print("=" * 60)
    print("예제 2: 출력 경로 지정")
    print("=" * 60)

    # 원하는 경로에 JSON 파일 저장
    result_path = hwp_to_json(
        hwp_path="document.hwp",
        json_path="output/my_document.json"
    )
    print(f"JSON 파일이 생성되었습니다: {result_path}")
    print()


def example3_parse_and_export_separately():
    """
    예제 3: 파싱과 내보내기를 분리
    """
    print("=" * 60)
    print("예제 3: 파싱과 내보내기 분리")
    print("=" * 60)

    # 1단계: HWP 파일 파싱
    data = parse_hwp("document.hwp")

    # 2단계: 파싱된 데이터 확인
    print(f"문서 이름: {data['metadata']['file_name']}")
    print(f"섹션 수: {len(data['sections'])}")
    print(f"문단 수: {len(data['sections'][0]['paragraphs'])}")

    # 3단계: JSON으로 내보내기
    result_path = export_to_json(data, "output/parsed_document.json")
    print(f"JSON 파일이 생성되었습니다: {result_path}")
    print()


def example4_parser_with_context_manager():
    """
    예제 4: 컨텍스트 매니저를 사용한 파싱
    """
    print("=" * 60)
    print("예제 4: 컨텍스트 매니저 사용")
    print("=" * 60)

    # with 문을 사용하여 자동으로 리소스 정리
    with HwpParser("document.hwp") as parser:
        data = parser.parse()

        # 메타데이터 출력
        print("메타데이터:")
        for key, value in data['metadata'].items():
            print(f"  {key}: {value}")

        # 첫 번째 문단 출력
        if data['sections'] and data['sections'][0]['paragraphs']:
            first_para = data['sections'][0]['paragraphs'][0]
            print(f"\n첫 번째 문단:")
            print(f"  텍스트: {first_para['text'][:50]}...")

        # JSON으로 저장
        export_to_json(data, "output/context_manager_output.json")

    print()


def example5_process_multiple_files():
    """
    예제 5: 여러 HWP 파일을 일괄 변환
    """
    print("=" * 60)
    print("예제 5: 여러 파일 일괄 변환")
    print("=" * 60)

    # HWP 파일 목록
    hwp_files = [
        "document1.hwp",
        "document2.hwp",
        "document3.hwp",
    ]

    output_dir = Path("output/batch")
    output_dir.mkdir(parents=True, exist_ok=True)

    for hwp_file in hwp_files:
        try:
            # 각 파일을 JSON으로 변환
            json_file = output_dir / f"{Path(hwp_file).stem}.json"
            result_path = hwp_to_json(hwp_file, str(json_file))
            print(f"✓ {hwp_file} → {result_path}")
        except FileNotFoundError:
            print(f"✗ {hwp_file} - 파일을 찾을 수 없습니다")
        except Exception as e:
            print(f"✗ {hwp_file} - 오류 발생: {e}")

    print()


def example6_custom_json_formatting():
    """
    예제 6: JSON 포맷 커스터마이징
    """
    print("=" * 60)
    print("예제 6: JSON 포맷 커스터마이징")
    print("=" * 60)

    # 파싱
    data = parse_hwp("document.hwp")

    # 들여쓰기 없이 압축된 JSON 생성
    compact_path = hwp_to_json(
        "document.hwp",
        "output/compact.json",
        indent=None
    )
    print(f"압축된 JSON: {compact_path}")

    # 들여쓰기 4칸으로 생성
    formatted_path = hwp_to_json(
        "document.hwp",
        "output/formatted.json",
        indent=4
    )
    print(f"포맷된 JSON: {formatted_path}")

    print()


def main():
    """메인 함수 - 모든 예제 실행"""
    print("\n" + "=" * 60)
    print("HWP to JSON 변환 예제 모음")
    print("=" * 60 + "\n")

    # 각 예제 함수 실행
    # 실제로 실행하려면 유효한 HWP 파일이 필요합니다
    # example1_simple_conversion()
    # example2_custom_output_path()
    # example3_parse_and_export_separately()
    # example4_parser_with_context_manager()
    # example5_process_multiple_files()
    # example6_custom_json_formatting()

    print("\n사용법:")
    print("  1. HWP 파일을 준비합니다")
    print("  2. 위의 예제 중 하나를 선택하여 주석을 해제합니다")
    print("  3. 스크립트를 실행합니다")
    print("\n간단한 사용:")
    print("  from HwpApi import hwp_to_json")
    print("  hwp_to_json('your_document.hwp')")
    print()


if __name__ == "__main__":
    main()
