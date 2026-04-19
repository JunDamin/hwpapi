# 문서 기본값 · 프리셋 · 테마 시스템 계획서

> 작성일: 2026-04-15 (v0.0.13 기준)
> 참고: `ref/승승아빠-한글매크로-250124(공개용).zip` — JScript 기반 29개 문서 꾸미기 매크로

---

## 배경

승승아빠 매크로는 **"한번 누르면 문서가 보기 좋아지는"** 29개의 JScript 함수를
`Alt+숫자` 단축키에 bind 한 번들이다. 주요 기능은:

| 단축키 | 기능 |
|:---|:---|
| Alt+1 | 문서 타이틀 박스 |
| Alt+2 | 문서 소제목 바 |
| Alt+3 | 제목차례 + 책갈피 추가 |
| Alt+4~5 | 표 전체 라인 / 제목행·바닥행 배경 |
| Alt+6 | 현재 행 배경 순환 (회색 → 노랑) |
| Alt+7 | 선택 셀 테두리 강조 (빨강 → 파랑) |
| Alt+8 | 표 제목행 반복 |
| Alt+9 | 요약 박스 2종 |
| Alt+0 | 줄 나눔 방식 순환 (글자 ↔ 어절) |
| Alt+- | 표 좌우 선 없음 + 상하 선 두껍게 |
| Alt+Shift+1 | 바탕쪽 테두리 |
| Alt+Shift+2 | 글자 배경 노랑 (형광펜) |
| Alt+Shift+3 | PDF 저장 |
| Alt+Shift+4~5 | 표 정리 · 컬러 보고서용 |
| Alt+Shift+6 | 현재 행 녹색 |
| Alt+Shift+7 | 현재 셀 배경색 순환 (6색) |
| Alt+Shift+8~9 | 자간 · 장평 동시 조정 |
| Alt+Shift+0 | 숫자 → 한글 숫자 |
| Alt+Shift+- | 엑셀 복사 표의 빈칸 삭제 |
| (버튼) | 목차 만들기 · 쪽번호 · 머리글 · 이미지 회색처리 · 크기조절 등 |

핵심 가치: **"공무원·연구자가 복사-붙여넣기로 만드는 HWP 문서를 깔끔하게 정돈"**.
hwpapi 에도 동일한 유형의 "**사전 설정된 보기 좋은 결과**"를 제공하면
사용성이 크게 향상된다.

---

## 설계 원칙

### 원칙 1 — 3 단계 레이어 (raw → preset → theme)

```
┌──────────────────────────────────┐
│ Theme        app.theme = "report" │   ← 문서 전체 분위기 (한번에 여러 preset 적용)
├──────────────────────────────────┤
│ Preset       app.preset.title_box() │   ← 단일 지점 꾸미기 (승승아빠 매크로 대응)
├──────────────────────────────────┤
│ Raw API      app.insert_table(…)    │   ← 세밀 제어 (기존 API)
└──────────────────────────────────┘
```

사용자는 자신의 필요에 맞는 레이어를 고를 수 있다:
- **Theme**: 10초 만에 전문가 문서 느낌
- **Preset**: 표 하나만 보기 좋게
- **Raw**: 완전 수동

### 원칙 2 — "보기 좋은 기본값" 내장 (Opinionated Defaults)

승승아빠 매크로처럼 **특정 값으로 하드코딩된 결과**를 제공.
(사용자가 원하면 override 가능하지만, 기본적으로는 "한번 눌러서 이쁘게".)

예: `app.preset.title_box()` 호출하면 자동으로:
- 168mm x 22.6mm 1x2 표
- 글자 취급
- `table_clear.tbl` 템플릿
- 제목 굵게, 14pt, 가운데 정렬

### 원칙 3 — 단축키 bind (선택 사항)

`app.keyboard.bind("Alt+1", app.preset.title_box)` 같이 승승아빠 매크로처럼
단축키로도 호출 가능하게. Windows API 훅 없이 HWP 의 키보드 이벤트만 사용.

### 원칙 4 — Python 친화적 API

```python
# JScript (승승아빠 매크로) — 100 줄씩 복잡
function OnScriptMacro_문서타이틀박스() {
    HAction.GetDefault("TableCreate", HParameterSet.HTableCreation.HSet);
    with (HParameterSet.HTableCreation) { /* 60 줄 */ }
    HAction.Execute("TableCreate", HParameterSet.HTableCreation.HSet);
    // … TableTreatAsChar, TableCellBlock, TableTemplate, CharShape, ParaShape …
}

# hwpapi 대응 — 한 줄
app.preset.title_box(text="보고서 제목", subtitle="부제목")
```

### 원칙 5 — Overridable

모든 프리셋은 색상·크기·폰트 등을 kwargs 로 override 가능:

```python
app.preset.title_box(
    text="제목",
    width="168mm",              # 기본값
    height="22.6mm",
    title_font_size=14,
    title_color="#000000",
    bg_color="#EEEEEE",         # override
    template="table_clear.tbl",
)
```

---

## v1.0 청사진에 통합할 API

### `app.preset` — 단일 지점 프리셋

`hwpapi.presets.Presets` 클래스. 29개 매크로를 메소드로.

#### 문서 구조
```python
app.preset.title_box(text, subtitle="", *, width="168mm", height="22.6mm")
app.preset.subtitle_bar(text, *, bg_color="#EEEEEE")
app.preset.summary_box(text, variant="rounded" | "boxed")
app.preset.header_footer(filename=True, page_number=True, page_total=True)
app.preset.page_border(style="dashed")             # 바탕쪽 테두리
```

#### 표
```python
app.preset.table_clean(top_only=True)               # Alt+- : 좌우 선 제거
app.preset.table_header_gray()                       # Alt+4 : 상단 진회색
app.preset.table_header_sky()                        # 상단 하늘색 (컬러 보고서용)
app.preset.table_footer_gray()                       # 하단 진회색
app.preset.table_footer_sky()                        # 하단 하늘색
app.preset.repeat_header()                           # Alt+8 : 제목행 반복
app.preset.row_cycle(colors=["#F0F0F0", "#FFFF99"])  # Alt+6 : 현재 행 순환
app.preset.cell_border_highlight(cycle=["#FF0000", "#0000FF"])
app.preset.clean_excel_paste()                       # Alt+Shift+- : 빈칸 삭제
```

#### 문자 / 문단
```python
app.preset.highlight_yellow()                        # Alt+Shift+2 : 형광펜 토글
app.preset.compress_spacing(step=1)                  # Alt+Shift+8 : 자간·장평 축소
app.preset.expand_spacing(step=1)                    # Alt+Shift+9 : 자간·장평 확대
app.preset.wrap_by_word()                            # Alt+0 : 어절 단위 줄 나눔
app.preset.number_to_korean()                        # Alt+Shift+0 : 숫자 → 한글
```

#### 이미지
```python
app.preset.grayscale_selected()                      # 선택된 이미지만 회색
app.preset.resize_all_images(width="100mm")          # 전체 이미지 일괄 크기
```

#### 내비게이션 / 메타
```python
app.preset.toc(with_bookmarks=True, dot_leader=True)  # Alt+3 : TOC + 책갈피
app.preset.page_numbers(style="바탕쪽")
```

#### 파일
```python
app.preset.save_pdf(two_per_page=False)              # Alt+Shift+3 : PDF 저장
```

### `app.theme` — 문서 전체 테마

한 번의 호출로 **여러 프리셋을 조합**해서 문서 분위기를 완성.

```python
# 공식 문서 (회색·진지함)
app.theme = "official"
# = title_box("회색") + table_header_gray() + repeat_header() + page_number + toc

# 컬러 보고서 (하늘색·강조)
app.theme = "color_report"
# = title_box("하늘색") + table_header_sky() + row_cycle() + 바탕쪽_테두리

# 연구 보고서 (담백)
app.theme = "research"
# = subtitle_bar() + table_clean(top_only=True) + toc(dot_leader)

# 템플릿만 준비 (empty skeleton)
app.theme = "blank_report"
# = 표지 + 목차 자리 + 헤더/풋터
```

내부적으로는 `hwpapi.themes.Theme` 클래스의 레지스트리. 커스텀 테마도
등록 가능:

```python
from hwpapi.themes import register_theme, Theme

register_theme("my_theme", Theme(
    title_box={"bg_color": "#003366", "font_color": "#FFFFFF"},
    table_style="header_sky",
    page_border=True,
    repeat_header=True,
))

app.theme = "my_theme"
```

### `app.scaffold` — 문서 뼈대 한 번에

```python
app.scaffold.report(
    title="2026년 정기 보고서",
    subtitle="자동화 부서",
    sections=["1. 개요", "2. 현황", "3. 제언"],
    theme="official",
)
# 결과:
# - 표지 (title_box)
# - 소제목 (subtitle_bar)
# - 목차 placeholder
# - 각 섹션 heading (빈 내용)
# - 꼬리말 · 페이지 번호
```

### `app.config` — 기본값 설정

App 인스턴스 전역의 "기본 선호":

```python
# 모든 insert_text/styled_text 가 이 폰트 사용
app.config.default_font = "함초롬바탕"
app.config.default_size = 11
app.config.default_line_spacing = 160

# 표의 기본 스타일
app.config.default_table_style = "header_sky"

# 색상 팔레트
app.config.palette = {
    "primary": "#003366",
    "accent": "#FF6600",
    "bg_alt": "#EEEEEE",
}

# hwpapirc 파일로도 저장/로드
app.config.load("~/.hwpapirc")
app.config.save()
```

### `app.keyboard` — 단축키 bind (선택 사항)

```python
# 승승아빠 매크로 재현 (선택적)
app.keyboard.bind("Alt+1", app.preset.title_box)
app.keyboard.bind("Alt+2", app.preset.subtitle_bar)
app.keyboard.bind("Alt+3", app.preset.toc)

# 또는 프로필 전체 활성화
app.keyboard.profile("seungseung")   # 승승아빠 매크로와 동일한 bind 적용
```

HWP 의 `RegisterAction` + COM 이벤트를 활용하거나, polling 방식으로
간단히 구현 가능.

---

## 구현 전략

### Phase 1 — 최소 구현 (v0.0.14 ~ v0.0.16, 약 2주)

우선 **가장 자주 쓰이는 10개 preset** 만:

1. `title_box` (Alt+1 대응)
2. `subtitle_bar` (Alt+2)
3. `table_header_gray` / `table_header_sky` (Alt+4, Alt+Shift+4)
4. `repeat_header` (Alt+8)
5. `row_cycle` (Alt+6)
6. `page_border` (Alt+Shift+1)
7. `highlight_yellow` (Alt+Shift+2)
8. `page_numbers`
9. `toc` (Alt+3)
10. `save_pdf` (Alt+Shift+3)

각 preset 은 별도 모듈 (`hwpapi/presets/text.py`, `presets/table.py`, …) 에
순수 함수로 구현, `Presets` 클래스에서 bound method 로 노출.

### Phase 2 — 테마 시스템 (v0.0.17 ~ v0.0.19, 약 2주)

- `hwpapi/themes/` 패키지
- `Theme` dataclass + registry
- 3 ~ 4 개 built-in theme (`official`, `color_report`, `research`, `blank`)
- `app.theme = "…"` setter

### Phase 3 — 스캐폴딩 + 설정 (v0.0.20 ~ v0.0.21, 약 2주)

- `app.scaffold.report(...)`, `.letter(...)`, `.contract(...)`
- `app.config` + `~/.hwpapirc` 지원
- 테마 · 프리셋 · 스캐폴드가 `app.config` 를 공유

### Phase 4 — 고급 (v0.0.22+, 선택)

- `app.keyboard` 단축키 bind
- Preset Gallery 웹페이지 (각 preset 의 before/after PNG)
- 사용자 정의 preset 플러그인 (`hwpapi.preset_plugins` entry_point)

---

## 파일 구조

```
hwpapi/
├── presets/
│   ├── __init__.py           # class Presets
│   ├── structure.py          # title_box, subtitle_bar, summary_box, …
│   ├── table.py              # table_header_*, row_cycle, …
│   ├── text.py               # highlight_yellow, compress_spacing, …
│   ├── image.py              # grayscale_selected, resize_all_images
│   └── navigation.py         # toc, page_numbers, page_border
├── themes/
│   ├── __init__.py           # Theme + registry
│   ├── official.py
│   ├── color_report.py
│   └── research.py
├── scaffold.py               # app.scaffold (report, letter, ...)
├── config.py                 # Config + hwpapirc
└── keyboard.py               # optional, Phase 4
```

---

## 예상 사용 시나리오

### 시나리오 A — "공공기관 보고서" 30초 완성

```python
from hwpapi.core import App

app = App()
app.theme = "official"
app.scaffold.report(
    title="2026년 상반기 업무 보고",
    subtitle="○○과",
    sections=["1. 개요", "2. 추진 경과", "3. 향후 계획"],
)
app.save("report.hwp")
```

→ 표지, 목차, 섹션 헤더, 쪽번호, 바탕쪽 테두리까지 완비.

### 시나리오 B — 표만 보기 좋게

```python
# 엑셀에서 복사한 거친 표를 깔끔하게
app.preset.clean_excel_paste()
app.preset.table_header_sky()
app.preset.repeat_header()
app.preset.table_clean(top_only=True)
```

### 시나리오 C — 대량 문서 일괄 꾸미기

```python
from pathlib import Path
from hwpapi.core import App

app = App(is_visible=False)
with app.silenced("yes"):
    for path in Path("input/").glob("*.hwp"):
        app.open(path)
        app.theme = "color_report"      # 재테마링
        app.preset.page_numbers()
        app.save(f"output/{path.name}")
```

---

## 위험 요소 · 고민 사항

### 1. HWP 버전 의존성
`table_clear.tbl` 같은 template 이 HWP 버전마다 다를 수 있음.
→ **대안**: 프리셋 내부에서 템플릿을 HWP 내장 대신 우리가 직접 row/col/style
설정으로 재현. HWPX 파일로 bundled template 제공 옵션.

### 2. 폰트 의존성
`함초롬바탕`, `Noto Serif KR` 같은 폰트가 없으면 효과가 다름.
→ **대안**: `config.default_font` 에 fallback 체인 지정
(`"함초롬바탕, 맑은 고딕, sans-serif"`).

### 3. 사용자 취향 차이
"보기 좋음" 은 주관적. 과도하게 opinionated 하면 반감.
→ **대안**: 테마를 3 ~ 4 개만 제공, 나머지는 preset 단위 조합 권장.
`app.config.palette` / 각 preset 의 kwargs 로 micro-override 허용.

### 4. 기존 문서 덮어쓰기
`app.theme = "…"` 이 이미 존재하는 문서의 스타일을 덮어쓸 위험.
→ **대안**: `theme.apply(mode="replace" | "add_only" | "preview")`.

### 5. 국문·영문 병행
승승아빠 매크로는 한국어 공공 문서 지향. 영문 보고서에는 안 어울릴 수 있음.
→ **대안**: `theme="official_en"`, `"research_en"` 같은 i18n variant.

---

## 테스트 전략

1. **단위 테스트** — 각 preset 이 호출 시 올바른 COM 액션을 발생시키는지
   (Mock 기반, 현재 테스트 패턴 재사용).
2. **Golden snapshot** — 각 theme 적용 후 PDF export → PNG 스냅샷 비교
   (image-diff).
3. **승승아빠 매크로와 시각적 동등성** — 동일 입력 문서에 JScript 매크로 vs
   `app.preset` 호출 → PDF 비교.

---

## 다음 단계 (즉시 진행 가능)

1. **`hwpapi/presets/__init__.py` 생성** + `Presets` 클래스 뼈대
2. **`title_box` 프리셋 구현** — 승승아빠 매크로의 `OnScriptMacro_문서타이틀박스`
   를 Python 으로 번역 (약 60 줄)
3. **`table_header_sky` 구현** — 표 현재 셀 기준 상단행 배경색
4. **`app.preset` wiring** — `App.__init__` 에서 `self.preset = Presets(self)`
5. **before/after PDF 렌더링 스크립트** — 시각적 회귀 테스트 기반
6. **튜토리얼 추가** — `nbs/01_tutorials/11_presets.ipynb`

이 6단계만 완료하면 v0.0.14 릴리즈 가능. 그 뒤 phase 2 (theme) → phase 3
(scaffold) 순서로 진행.

---

## 참고

- `ref/extracted/승승아빠-매크로-250124(공개용).hmi` — 원본 JScript 매크로
- `ref/extracted/승승아빠_한글매크로.hwpx` — 사용 설명서
- `docs/V1_CONSISTENCY_PLAN.md` — 네이밍/패턴 원칙 (preset 도 이를 따름)
- `docs/REFACTORING_PLAN.md` — 전체 로드맵

**업데이트**: v1.0 청사진에 `app.preset` / `app.theme` / `app.scaffold` /
`app.config` 가 정식 accessor 로 편입되도록 V1_CONSISTENCY_PLAN.md 에
항목 추가 필요 (별도 커밋).
