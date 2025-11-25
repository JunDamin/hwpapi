"""
간단한 HWP 파서 사용 예제
"""

from HwpApi import hwp_to_json, parse_hwp

# 방법 1: 한 줄로 HWP를 JSON으로 변환
hwp_to_json("document.hwp")  # document.json 생성

# 방법 2: 출력 파일명 지정
hwp_to_json("document.hwp", "output.json")

# 방법 3: 파싱만 수행 (데이터 분석)
data = parse_hwp("document.hwp")

# 문서 정보 출력
print(f"파일명: {data['metadata']['file_name']}")
print(f"섹션 수: {len(data['sections'])}")

# 모든 문단 텍스트 출력
for section in data['sections']:
    for para in section['paragraphs']:
        print(para['text'])
