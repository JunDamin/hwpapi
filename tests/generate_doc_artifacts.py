"""
Generate tutorial screenshots from REAL HWP → PDF → PNG output.

Each demo runs in its own Python subprocess for clean state.  Drives
HWP, saves the document as PDF, then renders page 1 to PNG via PyMuPDF.

IMPORTANT: we never set shade_color="#FFFFFF" as a "reset" because that
creates a white-shade attribute on characters which HWP renders as a
subtle tinted background in the PDF output.  The only way to REMOVE
shade is StyleClearCharStyle.  For demos that need to end a highlighter
block, we clear char style before continuing (via SelectAll is
destructive, so instead we turn off shade by setting it back to the
background color and also explicitly clear subsequent segments).

Run:  python tests/generate_doc_artifacts.py
"""
from __future__ import annotations
import os
import sys
import tempfile
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.abspath(os.path.join(HERE, "..", "nbs", "01_tutorials", "img"))
os.makedirs(IMG_DIR, exist_ok=True)


# ── Demo scripts (each runs in a fresh HWP session) ────────────
# Each string is Python code.  `app` is an already-initialised App() with
# an empty document and clean char/para state (no shade, no shadow).

DEMOS = {}

DEMOS["demo_feature_tour"] = r"""
# Title (center, big, bold)
app.insert_text("hwpapi 기능 둘러보기\n")
app.move.top_of_file(); app.select_text()
app.set_charshape(bold=True, height=2400)
app.set_parashape(align_type=3)
app.move.bottom_of_file()
# Clear styles so following content starts clean
app.api.Run("StyleClearCharStyle")
app.set_parashape(align_type=1)

app.insert_text("\n── 문자 서식 ──\n")

# Each format on its own line so paragraph-level state resets between.
# For each line: clear, set format, type, break paragraph.
def demo_line(label, text, **fmt):
    app.api.Run("StyleClearCharStyle")
    app.insert_text(label + ": ")
    if fmt:
        app.set_charshape(**fmt)
    app.insert_text(text)
    app.api.Run("StyleClearCharStyle")
    app.insert_text("\n")

demo_line("굵게",    "이 줄은 굵은 글씨입니다", bold=True)
demo_line("기울임",  "이 줄은 기울임 꼴입니다", italic=True)
demo_line("밑줄",    "이 줄은 밑줄이 있습니다", underline_type=1)
demo_line("취소선",  "이 줄은 취소선이 있습니다", strike_out_type=1)
demo_line("색",      "이 줄은 빨간 글씨입니다", text_color="#FF0000")
demo_line("큰 글씨", "이 줄은 크게 표시됩니다", height=2000)
# NB: shade_color (형광펜) intentionally omitted from this multi-line
# demo — once applied at cursor, HWP's current paragraph retains the
# paragraph-level fill and there's no public action to clear it without
# affecting the whole document. The highlighter feature IS shown in
# the report-generation demo where it's applied to a single span with
# paragraph-break boundaries.

app.api.Run("StyleClearCharStyle")
app.insert_text("\n── 문단 정렬 ──\n")
app.set_parashape(align_type=1); app.insert_text("왼쪽 정렬 문단입니다.\n")
app.set_parashape(align_type=3); app.insert_text("가운데 정렬 문단입니다.\n")
app.set_parashape(align_type=2); app.insert_text("오른쪽 정렬 문단입니다.\n")
app.set_parashape(align_type=1)

app.insert_text("\n── 표 자동 생성 ──\n")
action = app.actions.TableCreate
action.pset.Rows = 3; action.pset.Cols = 3
action.run()
for row in [["제품","매출","증감"],["A","1,200만원","+15%"],["B","800만원","-3%"]]:
    for cell in row:
        app.insert_text(cell); app.api.Run("TableRightCell")
app.move.bottom_of_file()
"""

DEMOS["demo_meeting_minutes"] = r"""
import datetime

# Title
app.insert_text("2026년 1분기 정기회의\n")
app.move.top_of_file(); app.select_text()
app.set_charshape(bold=True, height=2000)
app.set_parashape(align_type=3)
app.move.bottom_of_file()
app.set_charshape(bold=False, height=1000)
app.set_parashape(align_type=1)

today = datetime.datetime.now().strftime("%Y년 %m월 %d일")
app.insert_text("\n□ 회의 일시: " + today + "\n\n")
app.insert_text("□ 참석자\n")

action = app.actions.TableCreate
action.pset.Rows = 4; action.pset.Cols = 2
action.run()
for name, org in [("이름", "소속"), ("김철수", "기획팀"),
                   ("이영희", "개발팀"), ("박민수", "마케팅팀")]:
    app.insert_text(name); app.api.Run("TableRightCell")
    app.insert_text(org);  app.api.Run("TableRightCell")
app.move.bottom_of_file()

app.insert_text("\n□ 안건\n")
for i, item in enumerate(["1분기 실적 검토",
                           "2분기 신규 프로젝트 계획",
                           "인력 충원 논의"], 1):
    app.insert_text("  " + str(i) + ". " + item + "\n")

app.insert_text("\n□ 결론\n")
app.set_charshape(bold=True); app.insert_text("● ")
app.set_charshape(bold=False)
app.insert_text("2분기 프로젝트 A/B/C 착수, 인력 2명 충원 확정\n")
"""

DEMOS["demo_data_to_table"] = r"""
app.insert_text("2026년 1분기 지역별 매출\n\n")
app.move.top_of_file(); app.select_text()
app.set_charshape(bold=True, height=1600)
app.move.bottom_of_file()
app.set_charshape(bold=False, height=1000)

action = app.actions.TableCreate
action.pset.Rows = 6; action.pset.Cols = 3
action.run()
app.set_charshape(bold=True); app.set_parashape(align_type=3)
for h in ["지역", "매출 (원)", "전년비"]:
    app.insert_text(h); app.api.Run("TableRightCell")
app.set_charshape(bold=False); app.set_parashape(align_type=1)
for r in [("서울", "1,850,000,000", "+18.5%"),
          ("경기", "1,320,000,000", "+14.2%"),
          ("인천", "620,000,000", "+8.9%"),
          ("부산", "440,000,000", "+5.6%"),
          ("기타", "350,000,000", "-2.1%")]:
    for v in r:
        app.insert_text(v); app.api.Run("TableRightCell")
app.move.bottom_of_file()
"""

DEMOS["demo_report_generation"] = r"""
import datetime

app.insert_text("2026년 1분기 영업 보고\n")
app.move.top_of_file(); app.select_text()
app.set_charshape(bold=True, height=2400)
app.set_parashape(align_type=3)
app.move.bottom_of_file()
app.set_charshape(bold=False, height=1400)
app.insert_text("― 주요 지역 매출 현황 분석 ―\n")
app.set_charshape(height=1100, italic=True)
today = datetime.datetime.now().strftime("%Y년 %m월 %d일")
app.insert_text("\n" + today + "\n\n\n")
app.set_charshape(italic=False, height=1000)
app.set_parashape(align_type=1)

app.set_charshape(bold=True, height=1600)
app.insert_text("1. 요약\n")
app.set_charshape(bold=False, height=1000)
app.insert_text("2026년 1분기 전사 매출은 전년 동기 대비 12.5% 성장한 "
                "45.8억원을 기록했습니다.\n\n")

app.set_charshape(bold=True, height=1600)
app.insert_text("2. 지역별 매출\n")
app.set_charshape(bold=False, height=1000)
action = app.actions.TableCreate
action.pset.Rows = 4; action.pset.Cols = 3
action.run()
app.set_charshape(bold=True); app.set_parashape(align_type=3)
for h in ["지역", "매출", "전년비"]:
    app.insert_text(h); app.api.Run("TableRightCell")
app.set_charshape(bold=False); app.set_parashape(align_type=1)
for r in [("서울", "18.5억", "+18.5%"),
          ("경기", "13.2억", "+14.2%"),
          ("인천", "6.2억", "+8.9%")]:
    for v in r:
        app.insert_text(v); app.api.Run("TableRightCell")
app.move.bottom_of_file()

app.insert_text("\n")
app.set_charshape(bold=True, height=1600)
app.insert_text("3. 분석\n")
app.set_charshape(bold=False, height=1000)
app.insert_text("수도권 매출 비중이 전체의 83%에 달합니다. ")
app.set_charshape(shade_color="#FFFF00")
app.insert_text("서울·경기 지역 성장세가 전체 실적을 견인하는 구조")
# Turn off highlighter — clear char style at cursor
app.api.Run("StyleClearCharStyle")
app.insert_text("입니다.\n\n")

app.set_charshape(bold=True, height=1600)
app.insert_text("4. 결론\n")
app.set_charshape(bold=False, height=1000)
app.set_charshape(bold=True); app.insert_text("▸ ")
app.set_charshape(bold=False)
app.insert_text("2분기에는 지방 거점 영업력 강화와 신제품 B의 "
                "전국 확대에 집중\n")
"""

DEMOS["demo_bulk_edit"] = r"""
app.set_charshape(bold=True, height=1600)
app.insert_text("폴더 내 일괄 편집 — 실행 로그 예시\n\n")
app.set_charshape(bold=False, height=900)
app.insert_text("대상 파일 4개 발견\n")
app.insert_text("[1/4] contract_kim.hwp 처리 중...\n")
app.insert_text("  '(주)이전회사명' → '(주)새회사명': 3회\n")
app.insert_text("  'Old Address' → 'New Address': 1회\n")
app.insert_text("[2/4] report_2026_01.hwp 처리 중...\n")
app.insert_text("  '(주)이전회사명' → '(주)새회사명': 8회\n")
app.insert_text("  'Old Address' → 'New Address': 2회\n")
app.insert_text("[3/4] minutes.hwp 처리 중...\n")
app.insert_text("  '(주)이전회사명' → '(주)새회사명': 5회\n")
app.insert_text("  'Old Address' → 'New Address': 1회\n")
app.insert_text("[4/4] invoice.hwp 처리 중...\n")
app.insert_text("  '(주)이전회사명' → '(주)새회사명': 2회\n")
app.insert_text("  'Old Address' → 'New Address': 1회\n\n")
app.insert_text("완료: 4개 파일, 23회 치환\n")
"""


# Worker script — receives demo code path + output PNG path as argv
WORKER = r'''
import sys, time, os, tempfile
demo_code = open(sys.argv[1], encoding="utf-8").read()
out_png = sys.argv[2]

from hwpapi.core import App
import fitz  # PyMuPDF

app = App(is_visible=True)
time.sleep(0.5)

# Get a COMPLETELY fresh document with untouched char/para state.
# Simply SelectAll + Delete + StyleClearCharStyle does NOT reset the
# paragraph-level BorderFill (which carries shade color). FileNew is
# the only reliable way to drop all state.
try:
    app.api.Run("FileNew")
except Exception:
    pass
time.sleep(0.8)

# Monkey-patch insert_text so that "\n" in the text actually creates a
# paragraph break via BreakPara action. HWP's InsertText treats "\n"
# as a literal character, not a line break.
_orig_insert = app.insert_text
def _insert_text(text, *a, **kw):
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if part:
            _orig_insert(part, *a, **kw)
        if i < len(parts) - 1:
            app.api.Run("BreakPara")
app.insert_text = _insert_text

exec(compile(demo_code, "<demo>", "exec"),
     {"app": app, "__builtins__": __builtins__})
time.sleep(0.5)

# Save as PDF (extension-detected)
with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
    pdf_path = tmp.name
app.save(pdf_path)
time.sleep(0.8)

if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 500:
    with fitz.open(pdf_path) as pdf:
        page = pdf[0]
        # Crop to the content bounding box — removes most of the empty
        # page margin so the image isn't mostly blank.
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        pix.save(out_png)
    try: os.remove(pdf_path)
    except Exception: pass

    # Auto-crop: trim bottom whitespace so embedding the image in the
    # web page doesn't leave a huge empty area.
    try:
        from PIL import Image, ImageChops
        img = Image.open(out_png)
        bg = Image.new(img.mode, img.size, (255, 255, 255))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox:
            # Add small margin
            L, T, R, B = bbox
            pad = 40
            L = max(0, L - pad); T = max(0, T - pad)
            R = min(img.width, R + pad); B = min(img.height, B + pad)
            img.crop((L, T, R, B)).save(out_png)
    except Exception as e:
        print(f"crop skipped: {e}")

    print(f"OK {out_png}")
else:
    print(f"FAIL size={os.path.getsize(pdf_path) if os.path.isfile(pdf_path) else 'no file'}")
    sys.exit(1)
'''


def main():
    print(f"📁 Output: {IMG_DIR}")
    worker_path = os.path.join(tempfile.gettempdir(), "hwpapi_demo_worker.py")
    with open(worker_path, "w", encoding="utf-8") as f:
        f.write(WORKER)

    for name, code in DEMOS.items():
        print(f"\n[{name}]")
        code_path = os.path.join(tempfile.gettempdir(), f"hwpapi_{name}.py")
        with open(code_path, "w", encoding="utf-8") as f:
            f.write(code)
        out_png = os.path.join(IMG_DIR, f"{name}.png")

        try:
            result = subprocess.run(
                [sys.executable, worker_path, code_path, out_png],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
            )
            lines = (result.stdout + result.stderr).strip().split("\n")
            print("  " + (lines[-1] if lines else "")[:200])
            if result.returncode != 0:
                print("  stderr:", result.stderr[-400:])
        except Exception as e:
            print(f"  ⚠ {e}")

    print(f"\n✅ Done — images in {IMG_DIR}")


if __name__ == "__main__":
    main()
