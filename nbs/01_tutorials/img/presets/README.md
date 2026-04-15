# Presets 이미지 디렉터리

이 폴더의 이미지는 **튜토리얼 10~13의 시각 자료** 입니다.

## 이미지 성격

matplotlib 로 생성한 **개념 일러스트** 입니다 — API 호출이 만들어내는
결과의 "의도" 를 보여주기 위함. 실제 HWP 문서를 PDF로 렌더링한 스크린샷은
아닙니다.

실제 HWP 화면 캡처가 필요하면 `tests/generate_doc_artifacts.py` 를
Windows + HWP 환경에서 실행하세요.

## 재생성

이미지를 갱신하려면 (매크로 텍스트가 바뀌었거나 새 프리셋 추가 시):

```bash
# 원본 스크립트가 삭제되었으므로, 필요 시 git log 에서 복원하여 재사용:
git log --all -- nbs/_build_tutorial_images.py
```

## 이미지 목록

| 파일 | 사용 튜토리얼 |
|:---|:---|
| `app_help.png` | 10 |
| `app_repr.png` | 13 |
| `batch_mode_speedup.png` | 12 |
| `debug_trace.png` | 13 |
| `fields_dict.png` | 10 |
| `highlight_yellow.png` | 11 |
| `lint_report.png` | 13 |
| `striped_rows.png` | 11 |
| `subtitle_summary.png` | 11 |
| `table_header.png` | 11 |
| `title_box.png` | 11 |
| `toc.png` | 11 |
