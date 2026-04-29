"""
Generate v3 docs screenshots — REAL HWP → PDF → cropped PNG.

각 데모는 별도 Python subprocess (= 별도 HWP 세션) 에서 실행.
v3 surface (App + app.docs + Document) 만 사용 — `doc.insert_text` 가
이제 실제 작동하므로 헬퍼 (insert/styled/block) 도 단순해짐.

워커는 시작 직후 SetMessageBoxMode(0x00111111) 로 모든 native HWP
대화상자에 자동 Yes 응답. 문서 close 는 ``doc.close(save=False)`` —
Close(False) 직접 호출이라 dialog 자체가 안 뜸.

산출물: docs/_assets/img/v2/<demo_name>.png
qmd 페이지에서 ![](../_assets/img/v2/<name>.png) 으로 참조.

Run: python tests/generate_v2_doc_artifacts.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
IMG_DIR = os.path.join(ROOT, "docs", "_assets", "img", "v2")
os.makedirs(IMG_DIR, exist_ok=True)


# ── v3 API 데모 (각 문자열은 독립 HWP 세션에서 실행) ────────────
# 워커가 노출하는 변수:
#   app  : hwpapi.App
#   doc  : Document — app.docs.add() 로 막 만든 빈 문서
#   ctx  : hwpapi.context.scopes
#   U    : hwpapi.units
#
# 헬퍼:
#   styled(text, **fmt) - ctx.styled_text(app, text, **fmt) 단축
#   block(**fmt)        - ctx.charshape_scope(app, **fmt) 단축
DEMOS = {}


DEMOS["quickstart_overview"] = r"""
with block(bold=True, height=2400):
    doc.insert_text("hwpapi v3 — 5분 둘러보기")
doc.insert_text("\n\n")

doc.insert_text("v3 의 슬림한 facade 는 lifecycle 만 책임지고, 모든 문서 작업은 ")
styled("doc 인스턴스", bold=True, text_color="#2980B9")
doc.insert_text(" 에서 직접 수행합니다.\n\n")

with block(bold=True, height=1400, text_color="#34495E"):
    doc.insert_text("핵심 컬렉션")
doc.insert_text("\n")

for name, desc in [
    ("fields", "누름틀 — dict 처럼 사용"),
    ("bookmarks", "책갈피 — add/remove/goto"),
    ("hyperlinks", "하이퍼링크"),
    ("images", "삽입된 그림"),
    ("tables", "표 — Cell 단위 조작"),
]:
    doc.insert_text("  • ")
    styled("doc." + name, bold=True, text_color="#16A085")
    doc.insert_text(" — " + desc + "\n")
"""


DEMOS["charshape_formatting"] = r"""
with block(bold=True, height=2200):
    doc.insert_text("Character Shape Formatting")
doc.insert_text("\n\n")

doc.insert_text("styled_text() 와 charshape_scope() 로 깨끗한 서식 적용.\n\n")

items = [
    ("Bold",       {"bold": True}),
    ("Italic",     {"italic": True}),
    ("Underline",  {"underline_type": 1}),
    ("Strike",     {"strike_out_type": 1}),
    ("Red",        {"text_color": "#E74C3C"}),
    ("Highlight",  {"shade_color": "#F1C40F"}),
    ("Big",        {"height": 1800}),
    ("Mono",       {"text_color": "#16A085", "bold": True}),
]
for label, fmt in items:
    doc.insert_text("  " + label.ljust(11) + " │ ")
    styled("샘플 텍스트입니다", **fmt)
    doc.insert_text("\n")

doc.insert_text("\n")
with block(bold=True, height=1300, text_color="#34495E"):
    doc.insert_text("blockwise — charshape_scope")
doc.insert_text("\n")
with block(bold=True, text_color="#8E44AD"):
    doc.insert_text("  이 한 단락 전체가 보라색 굵은 글씨로 적용됩니다.\n")
doc.insert_text("  종료 후엔 자동으로 원래 서식으로 복원.\n")
"""


DEMOS["mail_merge_letter"] = r"""
import datetime

with block(bold=True, height=2400, text_color="#2C3E50"):
    doc.insert_text("주식회사 한컴오피스")
doc.insert_text("\n")
with block(height=1000, text_color="#7F8C8D"):
    doc.insert_text("서울시 종로구 신문로 X-XX  │  contact@example.com")
doc.insert_text("\n\n")

today = datetime.datetime.now().strftime("%Y년 %m월 %d일")
doc.insert_text(today + "\n\n")

styled("수신: ", bold=True)
doc.insert_text("홍길동 귀하\n")
styled("참조: ", bold=True)
doc.insert_text("기획팀\n\n")

with block(bold=True, height=1600):
    doc.insert_text("제목: 2026년 1분기 정기회의 안내")
doc.insert_text("\n\n")

doc.insert_text("안녕하세요, ")
styled("홍길동", bold=True, text_color="#2980B9")
doc.insert_text(" 님.\n\n")

doc.insert_text("귀사의 무궁한 발전을 기원합니다. 1분기 정기회의를 다음과 같이 ")
doc.insert_text("개최하오니 부디 참석하여 주시기 바랍니다.\n\n")

with block(bold=True, text_color="#34495E"):
    doc.insert_text("□ 일시")
doc.insert_text(": 2026년 4월 30일 (금) 오후 2시\n")
with block(bold=True, text_color="#34495E"):
    doc.insert_text("□ 장소")
doc.insert_text(": 본사 회의실 3층\n")
with block(bold=True, text_color="#34495E"):
    doc.insert_text("□ 안건")
doc.insert_text(": 1분기 실적 검토 / 2분기 계획 수립\n\n")

doc.insert_text("감사합니다.\n\n")
styled("기획팀 김철수 드림", bold=True)
"""


DEMOS["data_to_table"] = r"""
with block(bold=True, height=2000):
    doc.insert_text("매출 보고서 — Q1 2026")
doc.insert_text("\n\n")

doc.insert_text("리스트 → 표 자동 변환. ")
styled("doc.insert_table(rows, cols)", text_color="#2980B9", bold=True)
doc.insert_text(" 로 빈 표를 만든 뒤 셀별로 채웁니다.\n\n")

rows = [
    ("제품",   "1월",       "2월",       "3월",       "합계"),
    ("HWP",    "12,000원", "15,500원", "13,200원", "40,700원"),
    ("Cell",   "8,200원",  "9,100원",  "10,400원", "27,700원"),
    ("Office", "5,500원",  "6,300원",  "7,800원",  "19,600원"),
    ("총계",   "25,700원", "30,900원", "31,400원", "88,000원"),
]

doc.insert_table(rows=len(rows), cols=len(rows[0]))

for r, row in enumerate(rows):
    for c, cell in enumerate(row):
        if r == 0 or c == 0 or r == len(rows) - 1:
            styled(cell, bold=True)
        else:
            doc.insert_text(cell)
        app.api.Run("TableRightCell")

# 표 빠져나가기
app.api.Run("CloseEx")
doc.insert_text("\n\n")
styled("Tip: ", bold=True, text_color="#27AE60")
doc.insert_text("실무에선 Excel/CSV 데이터를 dict 로 읽어 같은 패턴으로 일괄 삽입.\n")
"""


# ── 워커 (subprocess 가 데모 하나당 1회 실행) ────────────────
WORKER = r'''
import sys, os, tempfile, time
import fitz  # PyMuPDF

demo_code_path = sys.argv[1]
out_png = sys.argv[2]
with open(demo_code_path, "r", encoding="utf-8") as f:
    demo_code = f.read()

from hwpapi import App, units as U
from hwpapi.context import scopes as ctx

app = App(is_visible=False)
time.sleep(0.4)

SILENCE_ALL_YES = 0x00111111
try:
    rc = app.api.SetMessageBoxMode(SILENCE_ALL_YES)
    if rc is False:
        raise RuntimeError("SetMessageBoxMode returned False")
except Exception as e:
    raise RuntimeError(f"SetMessageBoxMode 실패: {e}") from e

# v3 API — 새 빈 문서를 명시적으로 생성, doc 인스턴스 획득
doc = app.docs.add()
time.sleep(0.3)


def styled(text, **fmt):
    ctx.styled_text(app, text, **fmt)


def block(**fmt):
    return ctx.charshape_scope(app, **fmt)


# 데모 실행
exec(
    compile(demo_code, "<demo>", "exec"),
    {
        "app": app, "doc": doc, "ctx": ctx, "U": U,
        "styled": styled, "block": block,
        "__builtins__": __builtins__,
    },
)
time.sleep(0.5)

# PDF 저장 — v3: doc.save(path) 가 SaveAs 처리
with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
    pdf_path = tmp.name
doc.save(pdf_path)
time.sleep(0.7)

# 문서 close — v3: doc.close(save=False) 가 dialog 없이 닫음
doc.close(save=False)
time.sleep(0.2)

# PDF → cropped PNG
size = os.path.getsize(pdf_path) if os.path.isfile(pdf_path) else 0
if size <= 500:
    print(f"FAIL pdf size={size}")
    sys.exit(2)

with fitz.open(pdf_path) as pdf:
    page = pdf[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
    pix.save(out_png)
try:
    os.remove(pdf_path)
except OSError:
    pass

# Auto-crop
try:
    from PIL import Image, ImageChops
    img = Image.open(out_png)
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if bbox:
        L, T, R, B = bbox
        pad = 40
        L = max(0, L - pad); T = max(0, T - pad)
        R = min(img.width, R + pad); B = min(img.height, B + pad)
        img.crop((L, T, R, B)).save(out_png)
except Exception as e:
    print(f"  crop skipped: {e}")

# HWP 명시 종료는 생략 (subprocess 종료 시 정리)
print(f"OK {out_png}")
'''


def main() -> int:
    print(f"Output dir: {IMG_DIR}")
    print(f"Total demos: {len(DEMOS)}\n")

    with tempfile.NamedTemporaryFile(
        suffix=".py", delete=False, mode="w", encoding="utf-8"
    ) as f:
        f.write(WORKER)
        worker_path = f.name

    results = {}
    try:
        for i, (name, code) in enumerate(DEMOS.items(), 1):
            print(f"[{i}/{len(DEMOS)}] {name} — fresh HWP session")
            with tempfile.NamedTemporaryFile(
                suffix=".py", delete=False, mode="w", encoding="utf-8"
            ) as f:
                f.write(code)
                code_path = f.name
            out_png = os.path.join(IMG_DIR, name + ".png")
            try:
                rc = subprocess.call(
                    [sys.executable, worker_path, code_path, out_png],
                    cwd=ROOT,
                    timeout=120,
                )
                results[name] = "OK" if rc == 0 else f"exit={rc}"
            except subprocess.TimeoutExpired:
                results[name] = "TIMEOUT"
            except Exception as e:
                results[name] = f"ERR {type(e).__name__}: {e}"
            finally:
                try:
                    os.remove(code_path)
                except OSError:
                    pass
    finally:
        try:
            os.remove(worker_path)
        except OSError:
            pass

    print("\n=== Results ===")
    bad = 0
    for n, r in results.items():
        marker = "OK " if r == "OK" else "FAIL"
        print(f"  {marker} {n:30s} {r}")
        if r != "OK":
            bad += 1
    print(f"\n{len(results) - bad}/{len(results)} OK")
    print(f"Images: {IMG_DIR}")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
