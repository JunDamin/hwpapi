# Session Checkpoint — 다음 Claude Code 세션 인계용

**최종 업데이트**: 2026-04-29
**현재 버전**: **v3.0.0 (PyPI live)**
**라이브 사이트**: https://JunDamin.github.io/hwpapi/ (en) · https://JunDamin.github.io/hwpapi/ko/ (ko)

---

## 새 세션을 시작할 때 읽을 순서

1. **이 파일** (SESSION_CHECKPOINT.md) — 전체 상황 요약
2. [`CHANGELOG.md`](../CHANGELOG.md) — `[3.0.0]` 섹션 — 무엇이 release 됐는지
3. [`docs/design/adr/adr-003-multi-document.qmd`](../docs/design/adr/adr-003-multi-document.qmd) — accepted, v3 의 토대
4. [`docs/design/adr/adr-005-api-ergonomics.qmd`](../docs/design/adr/adr-005-api-ergonomics.qmd) — proposed, v3.1~v3.5 roadmap
5. [`docs/design/adr/adr-006-python-docx-patterns.qmd`](../docs/design/adr/adr-006-python-docx-patterns.qmd) — proposed, 도메인 패턴 비교
6. [`hwpapi/document.py`](../hwpapi/document.py) — v3 Document 본체 (~30+ 메소드)
7. [`hwpapi/collections/documents.py`](../hwpapi/collections/documents.py) — DocumentCollection
8. [`hwpapi/selection.py`](../hwpapi/selection.py) — Selection / Range (v3.1)

---

## 이번 세션 (~5월 호) 에서 한 일

### 릴리스
- **v1.0.0** → **v2.0.0** → **v3.0.0** PyPI 모두 live (Trusted Publisher 자동)
- v2 는 현재 frozen (`pip install hwpapi==2.*` 로 동결 가능)

### v3.0.0 주요 변경 (BREAKING)
- `App.open / save / save_as / close / doc` 제거 — **clean cut, no shim**
- `app.docs` 신설 (xlwings 의 `app.books` 와 동형)
  - `open(path)` / `add()` / `active` / `[i]` / `["name"]` / `iter` / `len` / `__contains__`
- `Document` 풀 surface — text I/O, lifecycle, 7 컬렉션, cursor, actions proxy
- `_DocActions` proxy — `doc.actions.X` 가 `__getattr__` 시점에 자동 `doc.activate()`

### v3.1 ergonomics (이번 릴리스 포함)
- `doc.set_charshape(**fmt)` / `set_parashape(**fmt)` — `with` 없이 즉시 적용
- `doc.selection` (Selection) / `doc.range(p1, p2)` (Range)
- `doc.find_all(query)` / `doc.replace_brackets({"{key}": "..."})`

### 문서
- ADR-003 (accepted) — Multi-document redesign
- ADR-004 (proposed) — Element-level API (Run/Paragraph/Cell/Section)
- ADR-005 (proposed) — API ergonomics (xlwings 차용 7개 + 직접 setter + 신규 메소드)
- ADR-006 (proposed) — python-docx 패턴 비교 (권장 차용 7개)
- migration v2→v3 가이드 (en + ko)
- 양언어 사이트 47+47 페이지

### CI / Tooling
- `tests/test_documents_collection.py` (24) + `tests/test_v3_setters_selection.py` (12)
- 전체 1,341 / 1,341 mock-based passes (windows-latest × Python 3.9~3.12)
- `tests/generate_v2_doc_artifacts.py` — v3 API 사용
- `docs/render-bilingual.{ps1,sh}` — en → ko 순서 보장
- `deploy-docs.yaml` — sanity check (en/ko 양쪽 최소 30 페이지)

---

## v3.x Roadmap (ADR-005/006 정의됨)

### v3.2 — 데이터 / 표 / 단위 클래스
- 단위 클래스 (`Mm(210) - Mm(25)*2`) — python-docx 차용 (ADR-006 A.4)
- Element API Phase 1+2 — `Paragraph.text`, `Cell.fill/border` (ADR-004)
- `doc.tables[i].to_dataframe()` / `from_csv()` (ADR-005)

### v3.3 — 메타데이터 / 페이지 / 윈도우
- `doc.metadata.{title,author,...}` — `core_properties` (ADR-006 A.5)
- `doc.statistics.{page_count,character_count,...}`
- `doc.zoom`, `doc.print()`, `doc.print_silent(...)`
- **`doc.sections` first-class** — 페이지 설정 + 머리말/꼬리말 (ADR-006 A.1)

### v3.4 — Style / outline / 고급
- **Style 객체 시스템** — `doc.styles["제목 1"]` 풀 객체 (ADR-006 A.2)
- **Tri-state 서식** — `bold = None` (상속) (ADR-006 A.3)
- `doc.headings`, `doc.outline.toc.update()`
- `doc.compare(other)` (diff)
- `doc.protect()` / `password`
- **`doc.contents()` / `body.iter()`** — 단락+표 문서 순서 (ADR-006 A.6)

### v3.5 — 실험적
- `doc.snapshot()` / `restore()` (in-memory)
- `hwpapi.view(df)` (xlwings 영감)
- `doc.preview()` (PDF + OS viewer)

---

## ADR 의 미해결 질문 (다음 spike 전 결정 필요)

### ADR-004 (Element API)
1. **Run 의 경계 정의** — HWP 가 동일 charshape 연속 글자 구간을 직접 노출하는지 검증
2. **Cell.border 의 dict 형 설계** — `{"top": "solid"}` 단변, `{"all": "solid"}` 4변, `{"outer", "inner"}` alias 지원 여부
3. **Element 동일성 (`==`)** — 같은 셀 두 번 가져오면 `==`?
4. **Section 범위** — 이번 ADR? 분리?
5. **iter 안정성** — generator (lazy) vs snapshot

### ADR-005 (Ergonomics)
1. **`set_X` 시그니처** — `**kwargs` only? `CharShape` 객체도?
2. **Range 시그니처** — 단락 / 글자 / 둘 다? "p2:p5" string 약식?
3. **`doc.batch()` vs `app.batch()`** — 어느 레벨에서 효과?
4. **`hwpapi.view(...)` 동작** — 새 App vs 기존 App 에 doc 추가?
5. **`metadata.author` r/w** — HWP COM 의 표준 setter 검증

### ADR-006 (python-docx 차용)
1. **Section 단위** — `doc.sections.add()` vs `doc.add_section()`? — 권장 `doc.sections.add(...)`
2. **단위 클래스 vs 함수** — `Mm(210)*2` vs `U.mm(210)*2` trade-off
3. **Tri-state 표기** — `bold=None` verbose, `bold=...` sentinel?
4. **Run-level 직접 property** — `bold`, `italic`, `underline` 3개만?
5. **`body.iter()` 명칭** — `doc.contents()` / `iter()` / `elements()` ?

---

## 다음 작업 후보 (우선순위 순)

### 1. v3.2 spike — 단위 클래스 + Element Phase 1
가장 자연스러운 다음 단계. ADR-006 A.4 + ADR-004 Phase 1.
- `hwpapi/units.py` 에 `Mm`, `Pt`, `Inch`, `Cm`, `Hwp` 클래스 추가 (기존 함수 보존)
- `hwpapi/elements/paragraph.py` 신설 — `Paragraph.text`, `parashape`
- `hwpapi/elements/run.py` — `Run.text`, `charshape`
- 테스트 + 데모 generator 갱신

### 2. ADR 의 미해결 질문 결정
spike 전에 ADR-004/005/006 의 미해결 질문 리스트를 한 번 훑고 결정.
사용자와 짧은 Q&A 로 5~10분.

### 3. v3.1 의 `find_all` 실제 검증
현재 `doc.find_all` 은 `RepeatFind` raw action 기반 placeholder.
실제 HWP 에서 동작 검증 + 더 견고한 구현 (GetText 파싱 등).

### 4. v2 nbs/01_tutorials 의 v1 코드 migration 또는 archive
`nbs/01_tutorials/*.ipynb` 13개가 여전히 v1 / v2 API. v3 으로 이주
하거나 `archive/` 로 옮길지 결정.

### 5. ADR-004 Phase 2 — Table + Cell
사용자 요청 빈도 높은 영역 (`cell.fill`, `cell.border`).

---

## 주요 파일 위치 (v3 기준)

```
hwpapi/
├── core/
│   ├── app.py             # App (process lifecycle 만)
│   ├── engine.py          # Engine
│   └── document.py        # ⚠️ 이전 Documents collection — 폐기 대상
├── document.py            # ⭐ v3 Document (모든 메소드)
├── selection.py           # ⭐ v3.1 Selection / Range
├── collections/
│   ├── documents.py       # ⭐ v3 DocumentCollection
│   ├── fields.py / bookmarks.py / hyperlinks.py / images.py /
│   ├── paragraphs.py / styles.py / tables.py
├── context/
│   └── scopes.py          # charshape_scope / parashape_scope / styled_text
├── low/                   # raw escape hatch
│   ├── actions.py / engine.py / parametersets/
└── presets/               # 11 preset (v1 시절 유지)

tests/
├── test_documents_collection.py     # ⭐ v3 (24)
├── test_v3_setters_selection.py     # ⭐ v3.1 (12)
├── test_documents.py
├── test_charshape_v2.py
├── test_collections_*.py            # 7 컬렉션
├── unit/                             # io/errors/elements 단위 테스트
└── generate_v2_doc_artifacts.py     # v3 데모 생성기

docs/
├── _quarto.yml + ko/_quarto.yml     # 양언어 sidebar
├── _site/                            # 정적 산출물 (Pages 업로드 대상)
├── _assets/img/v2/                   # 데모 PNG 4개
├── getting-started/                  # install/quickstart/migration v1→v2/v2→v3
├── guide/ recipes/ reference/ design/
└── ko/ (mirror)

claude_docs/
├── SESSION_CHECKPOINT.md  # ⭐ 이 파일
├── architecture.md / code-patterns.md / project-structure.md / ...
└── (legacy v2 시절 작성)
```

---

## 개발 환경 빠른 체크

```bash
# 패키지 설치
pip install -e ".[test,docs]"

# 단위 테스트 (Mock 기반, HWP 불필요)
python -m pytest tests/ -q --ignore=tests/test_all_actions.py \
  --ignore=tests/test_all_parametersets.py \
  --ignore=tests/smoke_features.py \
  --ignore=tests/smoke_scenarios.py \
  --ignore=tests/smoke_visual.py

# Smoke 테스트 (실제 HWP 필요 — Windows + 한컴오피스)
python tests/smoke_features.py

# 데모 PNG 재생성 (Windows + HWP)
python tests/generate_v2_doc_artifacts.py

# 양언어 사이트 빌드
bash docs/render-bilingual.sh    # 또는 pwsh docs/render-bilingual.ps1

# 패키지 빌드
python -m build
python -m twine check dist/*

# Release (수동)
git tag vX.Y.Z -m "..."
git push origin vX.Y.Z
gh release create vX.Y.Z dist/*.tar.gz dist/*.whl --title "..." --notes "..."
# → Trusted Publisher 가 자동으로 PyPI publish
```

---

## 알려진 이슈 / TODO

- `doc.find_all` 은 placeholder — 실제 HWP 에서 견고한 구현 필요
- `hwpapi/core/document.py` 의 `Document` / `Documents` 는 폐기됐지만 파일 잔존
- v1 시절 작성된 13개 nbs/01_tutorials 가 v3 와 비호환 (archive 또는 migrate)
- 데모 generator 의 색상 (`text_color`) 이 PDF 에 일부 적용 안 됨 (styled_text 동작 확인 필요)
- ADR-004/005/006 한국어판은 사실상 한국어이지만 cross-reference 일부 영문 잔존

---

## 사용자 / Maintainer 컨택

- GitHub: https://github.com/JunDamin/hwpapi
- PyPI: https://pypi.org/project/hwpapi/
- Author email: gpt@koica.go.kr (사용자 기록)
