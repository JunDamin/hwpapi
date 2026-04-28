"""
Generate v2 docs screenshots — REAL HWP → PDF → cropped PNG.

각 데모는 v2 API (App + context.scopes + actions) 만 사용.
모든 데모를 단일 HWP 세션에서 실행 (FileClose + FileNew 사이로 분리).

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


# ── v2 API 데모 (각 문자열은 별도 HWP 문서에서 실행) ────────────
# 사용 가능한 helper:
#   insert(text)             - 일반 텍스트 삽입 ("\n" → BreakPara)
#   styled(text, **fmt)      - ctx.styled_text 단축
#   block(**fmt) as ctx_mgr  - ctx.charshape_scope 단축
#   app, ctx, U              - hwpapi 모듈
DEMOS = {}


DEMOS["quickstart_overview"] = r"""
with block(bold=True, height=2400):
    insert("hwpapi v2 — 5분 둘러보기")
insert("\n\n")

insert("v2 의 슬림한 facade 는 lifecycle 만 책임지고, 모든 문서 작업은 ")
styled("app.doc", bold=True, text_color="#2980B9")
insert(" 아래에 모입니다.\n\n")

with block(bold=True, height=1400, text_color="#34495E"):
    insert("핵심 컬렉션")
insert("\n")

for name, desc in [
    ("fields", "누름틀 — dict 처럼 사용"),
    ("bookmarks", "책갈피 — add/remove/goto"),
    ("hyperlinks", "하이퍼링크"),
    ("images", "삽입된 그림"),
    ("tables", "표 — Cell 단위 조작"),
]:
    insert("  • ")
    styled("app.doc." + name, bold=True, text_color="#16A085")
    insert(" — " + desc + "\n")
"""


DEMOS["charshape_formatting"] = r"""
with block(bold=True, height=2200):
    insert("Character Shape Formatting")
insert("\n\n")

insert("styled_text() 와 charshape_scope() 로 깨끗한 서식 적용.\n\n")

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
    insert("  " + label.ljust(11) + " │ ")
    styled("샘플 텍스트입니다", **fmt)
    insert("\n")

insert("\n")
with block(bold=True, height=1300, text_color="#34495E"):
    insert("blockwise — charshape_scope")
insert("\n")
with block(bold=True, text_color="#8E44AD"):
    insert("  이 한 단락 전체가 보라색 굵은 글씨로 적용됩니다.\n")
insert("  종료 후엔 자동으로 원래 서식으로 복원.\n")
"""


DEMOS["mail_merge_letter"] = r"""
import datetime

with block(bold=True, height=2400, text_color="#2C3E50"):
    insert("주식회사 한컴오피스")
insert("\n")
with block(height=1000, text_color="#7F8C8D"):
    insert("서울시 종로구 신문로 X-XX  │  contact@example.com")
insert("\n\n")

today = datetime.datetime.now().strftime("%Y년 %m월 %d일")
insert(today + "\n\n")

styled("수신: ", bold=True)
insert("홍길동 귀하\n")
styled("참조: ", bold=True)
insert("기획팀\n\n")

with block(bold=True, height=1600):
    insert("제목: 2026년 1분기 정기회의 안내")
insert("\n\n")

insert("안녕하세요, ")
styled("홍길동", bold=True, text_color="#2980B9")
insert(" 님.\n\n")

insert("귀사의 무궁한 발전을 기원합니다. 1분기 정기회의를 다음과 같이 ")
insert("개최하오니 부디 참석하여 주시기 바랍니다.\n\n")

with block(bold=True, text_color="#34495E"):
    insert("□ 일시")
insert(": 2026년 4월 30일 (금) 오후 2시\n")
with block(bold=True, text_color="#34495E"):
    insert("□ 장소")
insert(": 본사 회의실 3층\n")
with block(bold=True, text_color="#34495E"):
    insert("□ 안건")
insert(": 1분기 실적 검토 / 2분기 계획 수립\n\n")

insert("감사합니다.\n\n")
styled("기획팀 김철수 드림", bold=True)
"""


DEMOS["data_to_table"] = r"""
with block(bold=True, height=2000):
    insert("매출 보고서 — Q1 2026")
insert("\n\n")

insert("리스트(딕셔너리) → 표 자동 변환. ")
styled("app.actions.TableCreate", text_color="#2980B9", bold=True)
insert(" 로 헤더와 데이터 행을 한 번에 작성.\n\n")

rows = [
    ("제품",   "1월",       "2월",       "3월",       "합계"),
    ("HWP",    "12,000원", "15,500원", "13,200원", "40,700원"),
    ("Cell",   "8,200원",  "9,100원",  "10,400원", "27,700원"),
    ("Office", "5,500원",  "6,300원",  "7,800원",  "19,600원"),
    ("총계",   "25,700원", "30,900원", "31,400원", "88,000원"),
]

action = app.actions.TableCreate
action.pset.Rows = len(rows)
action.pset.Cols = len(rows[0])
action.run()

for r, row in enumerate(rows):
    for c, cell in enumerate(row):
        if r == 0 or c == 0 or r == len(rows) - 1:
            styled(cell, bold=True)
        else:
            insert(cell)
        app.api.Run("TableRightCell")

# 표 빠져나가기
app.api.Run("CloseEx")
insert("\n\n")
styled("Tip: ", bold=True, text_color="#27AE60")
insert("실무에선 Excel/CSV 데이터를 dict 로 읽어 같은 패턴으로 일괄 삽입.\n")
"""


# ── 워커 (subprocess 가 실행할 코드) ────────────────────────
WORKER = r'''
import sys, json, os, tempfile, time
import fitz  # PyMuPDF
from contextlib import contextmanager
from hwpapi import App, units as U
from hwpapi.context import scopes as ctx

with open(sys.argv[1], "r", encoding="utf-8") as f:
    demos = json.load(f)
out_dir = sys.argv[2]

app = App(is_visible=False)
time.sleep(0.5)

# Silence dialogs (Save? prompts)
try:
    app.engine.api.SetMessageBoxMode(0x00111111)
except Exception:
    pass


def insert(text):
    """일반 텍스트 삽입 — \\n 은 BreakPara 로 변환."""
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if part:
            act = app.actions.InsertText
            act.pset.Text = part
            act.run()
        if i < len(parts) - 1:
            app.api.Run("BreakPara")


def styled(text, **fmt):
    """ctx.styled_text 단축 — \\n 분리는 사용하지 않음."""
    ctx.styled_text(app, text, **fmt)


def block(**fmt):
    """ctx.charshape_scope 단축."""
    return ctx.charshape_scope(app, **fmt)


results = {}
for name, code in demos.items():
    print(f"[{name}] running...")
    try:
        # Fresh document
        try:
            app.api.Run("FileClose")
        except Exception:
            pass
        try:
            app.api.Run("FileNew")
        except Exception:
            pass
        time.sleep(0.4)

        exec(
            compile(code, "<" + name + ">", "exec"),
            {
                "app": app, "ctx": ctx, "U": U,
                "insert": insert, "styled": styled, "block": block,
                "__builtins__": __builtins__,
            },
        )
        time.sleep(0.5)

        # Save as PDF — use app.save(path) which is full-document SaveAs;
        # app.save_as is the v1 save_block (selection only).
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = tmp.name
        app.save(pdf_path)
        time.sleep(0.7)

        if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 500:
            out_png = os.path.join(out_dir, name + ".png")
            with fitz.open(pdf_path) as pdf:
                page = pdf[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                pix.save(out_png)
            try: os.remove(pdf_path)
            except Exception: pass

            # Auto-crop trailing whitespace
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
            results[name] = "OK"
            print(f"  -> {out_png}")
        else:
            results[name] = "PDF empty"
    except Exception as e:
        results[name] = f"ERR {type(e).__name__}: {e}"
        print(f"  ERROR: {e}")

# Cleanup — try multiple paths since v2 surface is slim
for cmd in ("FileClose", "FileQuit"):
    try:
        app.api.Run(cmd)
    except Exception:
        pass

print("\n=== Results ===")
for n, r in results.items():
    print(f"  {n:30s} {r}")
'''


def main() -> int:
    print(f"Output dir: {IMG_DIR}\n")
    with tempfile.NamedTemporaryFile(
        suffix=".json", delete=False, mode="w", encoding="utf-8"
    ) as f:
        json.dump(DEMOS, f, ensure_ascii=False)
        demos_path = f.name
    with tempfile.NamedTemporaryFile(
        suffix=".py", delete=False, mode="w", encoding="utf-8"
    ) as f:
        f.write(WORKER)
        worker_path = f.name

    try:
        rc = subprocess.call(
            [sys.executable, worker_path, demos_path, IMG_DIR],
            cwd=ROOT,
        )
    finally:
        for p in (demos_path, worker_path):
            try:
                os.remove(p)
            except OSError:
                pass

    print(f"\nImages: {IMG_DIR}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
