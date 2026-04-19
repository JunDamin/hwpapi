# hwpapi

[![PyPI](https://img.shields.io/pypi/v/hwpapi.svg)](https://pypi.org/project/hwpapi/)
[![Python](https://img.shields.io/pypi/pyversions/hwpapi.svg)](https://pypi.org/project/hwpapi/)
[![License](https://img.shields.io/pypi/l/hwpapi.svg)](https://github.com/JunDamin/hwpapi/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-JunDamin.github.io-2780E3)](https://JunDamin.github.io/hwpapi)
[![Status](https://img.shields.io/badge/status-v2.0-success)](https://github.com/JunDamin/hwpapi)

**한글(HWP) 자동화를 위한 Pythonic 라이브러리.** `hwpapi.App` 슬림 퍼사드 +
`app.doc.*` 컬렉션 + `hwpapi.low.*` escape hatch 의 이중 레이어 API.

> 📖 **전체 문서는 [JunDamin.github.io/hwpapi](https://JunDamin.github.io/hwpapi)**
> — 설치 · 5분 투어 · Guide · Recipes · API Reference · v1 → v2 마이그레이션.

---

## Requirements

- **Windows** + 한컴 오피스 설치
- **Python 3.9+**
- `pywin32`

## Install

```sh
pip install hwpapi
```

## 5분 투어

```python
from hwpapi import App

app = App()                                   # HWP 연결
app.open("report.hwp")

# 컬렉션 — dict-like + iterable + filterable
app.doc.fields["name"] = "홍길동"
app.doc.bookmarks.add("ch1")
for tbl in app.doc.tables:
    print(tbl.caption, tbl.rows, "×", tbl.cols)

# 문단/런
for para in app.doc.paragraphs:
    print(para.style, "|", para.text[:40])

# Escape hatch — 저수준 actions / parametersets / engine
from hwpapi.low import actions, parametersets
char = parametersets.CharShape(app.actions.CharShape.pset)
char.bold = True
actions.CharShape(app).execute(char)

app.save()
```

더 많은 예제는 [Recipes](https://JunDamin.github.io/hwpapi/recipes/) 를 참고하세요.

---

## v1 → v2 변경점

v2 는 **public API break** 를 허용한 클린컷 릴리즈입니다.

| v1.x | v2.0 |
|---|---|
| `from hwpapi.core import App` | `from hwpapi import App` |
| `app.fields[...]` | `app.doc.fields[...]` |
| `app.bookmarks`, `app.images`, ... | `app.doc.bookmarks`, `app.doc.images`, ... |
| `app.charshape = {...}` | `with hwpapi.context.charshape_scope(app, bold=True): ...` |
| `app.preset.*`, `app.template.*` | *(Phase-out — Recipes 페이지에서 레시피 형태로 재공급)* |
| `hwpapi.parametersets` | `hwpapi.low.parametersets` |

전체 매핑 80+ 항목: [Migration Guide](https://JunDamin.github.io/hwpapi/getting-started/migration-v1-to-v2.html)

v1.x 에 머무르려면 `pip install "hwpapi<2"` 또는 `git checkout v1.x`.

---

## 기여

- 이슈/PR: [github.com/JunDamin/hwpapi](https://github.com/JunDamin/hwpapi)
- 아키텍처/의사결정: [docs/design/adr/](https://JunDamin.github.io/hwpapi/design/)

## License

MIT
