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
# Title — uses charshape_scope so formatting is auto-reset after
with app.charshape_scope(bold=True, height=2400):
    app.insert_text("hwpapi 기능 둘러보기")
app.set_parashape(align_type=3)
# Move past title, back to left align
app.insert_text("\n")
app.set_parashape(align_type=1)

app.insert_text("\n── 문자 서식 ──\n")

# Each demo uses styled_text so formatting is scoped to the marked word.
# This is the NEW IDIOMATIC way — no manual bold=True/bold=False pairs.
app.insert_text("굵게:    "); app.styled_text("이 줄은 굵은 글씨입니다", bold=True);           app.insert_text("\n")
app.insert_text("기울임:  "); app.styled_text("이 줄은 기울임 꼴입니다", italic=True);          app.insert_text("\n")
app.insert_text("밑줄:    "); app.styled_text("이 줄은 밑줄이 있습니다", underline_type=1);     app.insert_text("\n")
app.insert_text("취소선:  "); app.styled_text("이 줄은 취소선이 있습니다", strike_out_type=1);  app.insert_text("\n")
app.insert_text("색:      "); app.styled_text("이 줄은 빨간 글씨입니다", text_color="#FF0000"); app.insert_text("\n")
app.insert_text("형광펜:  "); app.styled_text("이 줄은 형광펜 표시입니다", shade_color="#FFFF00"); app.insert_text("\n")
app.insert_text("큰 글씨: "); app.styled_text("이 줄은 크게 표시됩니다", height=2000);          app.insert_text("\n")

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

# ═══════════════════════════════════════════════════════════════════
# v0.0.14~22 신규 기능 데모 (튜토리얼 10~13 용)
# ═══════════════════════════════════════════════════════════════════

# 줄무늬 표 (striped_rows preset)
DEMOS["demo_striped_rows"] = r"""
app.set_charshape(bold=True, height=1600)
app.insert_text("app.preset.striped_rows() 결과\n\n")
app.set_charshape(bold=False, height=1000)

# 데이터 표 생성 — insert_table 이 cursor 를 첫 셀에 남김
data = [["지역", "Q1", "Q2", "Q3", "Q4"],
        ["서울", "820", "910", "1,020", "1,150"],
        ["부산", "450", "480", "510", "560"],
        ["대구", "310", "330", "360", "390"],
        ["인천", "280", "310", "340", "370"],
        ["광주", "220", "240", "265", "290"]]
app.insert_table(data=data, headers=None)

# striped_rows 적용 (cursor 가 이미 첫 셀에 있음)
app.preset.striped_rows(
    colors=["#FFFFFF", "#F0F0F0"],
    header_color="#9FC5E8",
)
app.move.bottom_of_file()
"""

# title_box + subtitle_bar 조합
DEMOS["demo_title_subtitle"] = r"""
# 문서 타이틀 박스
app.preset.title_box(
    text="2026년 상반기 업무 보고",
    subtitle="기획재무팀",
    bg_color="#003366",
    title_color="#FFFFFF",
    font_size=1600,
)
app.insert_paragraph_break()
app.insert_paragraph_break()

# 소제목 + 본문
app.preset.subtitle_bar("1. 개요", bg_color="#EEEEEE")
app.insert_text("본 보고서는 2026년 상반기 주요 업무 추진 현황을\n")
app.insert_text("종합적으로 점검하기 위해 작성되었습니다.\n\n")

# 요약 박스
app.preset.summary_box(
    "핵심 요약: Q3 매출 18% 증가 · 신규 채널 3개 확보",
    bg_color="#FFF9E0",
)
app.insert_paragraph_break()

app.preset.subtitle_bar("2. 추진 경과", bg_color="#EEEEEE")
app.insert_text("\n[본문 내용 생략]\n")
"""

# table_header + table_footer 컬러 보고서
DEMOS["demo_table_presets"] = r"""
app.set_charshape(bold=True, height=1600)
app.insert_text("app.preset.table_header + table_footer\n\n")
app.set_charshape(bold=False, height=1000)

# 표 생성 (헤더 + 데이터 + 합계행)
data = [["제품", "수량", "단가", "합계"],
        ["노트북", "5", "1,200,000", "6,000,000"],
        ["모니터", "10", "350,000", "3,500,000"],
        ["키보드", "20", "80,000", "1,600,000"],
        ["합계", "35", "", "11,100,000"]]
app.insert_table(data=data, headers=None)

# 헤더 행 — 하늘색 (cursor 가 첫 셀에 있음)
app.preset.table_header(color="sky", text_color="#FFFFFF", bold=True)

# 합계 행 (마지막) — 회색
app.preset.table_footer(color="gray", text_color="#FFFFFF", bold=True)

app.move.bottom_of_file()
"""

# Fields dict-style 사용 예
DEMOS["demo_fields_dict"] = r"""
app.set_charshape(bold=True, height=1600)
app.insert_text("v0.0.12+ app.fields — dict-style\n\n")
app.set_charshape(bold=False, height=1000)

# 간단한 계약서 템플릿
app.insert_text("계약서\n\n")
app.insert_text("계약자: {{name}}\n")
app.insert_text("계약일: {{date}}\n")
app.insert_text("계약 금액: {{amount}}원\n")
app.insert_text("연락처: {{phone}}\n\n")

# {{tag}} → 필드 자동 변환
app.move.top_of_file()
app.fields.from_brackets()

# dict-style 일괄 주입
app.fields.update({
    "name": "홍길동",
    "date": "2026-04-15",
    "amount": "1,200,000",
    "phone": "010-1234-5678",
})

# 확인
app.move.bottom_of_file()
app.insert_text("\n\n────\n")
app.set_charshape(bold=True, height=900)
app.insert_text("app.fields.to_dict() 결과:\n")
app.set_charshape(bold=False, height=900)
import json
snap = app.fields.to_dict()
app.insert_text(json.dumps(snap, ensure_ascii=False, indent=2))
"""

# highlight_yellow toggle
DEMOS["demo_highlight_yellow"] = r"""
app.set_charshape(bold=True, height=1600)
app.insert_text("app.preset.highlight_yellow() — 토글\n\n")
app.set_charshape(bold=False, height=1100)

# 문장 내 특정 단어에 형광펜
app.insert_text("본 문서에서 가장 ")
app.styled_text("중요한", shade_color="#FFFF00")
app.insert_text(" 부분은 여기입니다.\n")

app.insert_text("또한, ")
app.styled_text("세 가지 원칙", shade_color="#FFFF00")
app.insert_text("을 지켜야 합니다:\n\n")

app.insert_text("  1. 명확성 — 오해 없이 전달\n")
app.insert_text("  2. ")
app.styled_text("일관성", shade_color="#FFFF00")
app.insert_text(" — 용어·스타일 통일\n")
app.insert_text("  3. 간결성 — 핵심만 전달\n")
"""

# TOC (점끌기탭)
DEMOS["demo_toc_preset"] = r"""
app.set_charshape(bold=True, height=2000)
app.set_parashape(align_type=3)
app.insert_text("목    차\n\n")
app.set_parashape(align_type=1)
app.set_charshape(bold=False, height=1100)

# 점끌기탭 흉내 — TabDef 을 기본 점끌기로 설정한 뒤 오른쪽 정렬된 번호
items = [
    ("1. 개요", "1", 0),
    ("1.1 배경", "2", 1),
    ("1.2 범위", "3", 1),
    ("2. 현황 분석", "5", 0),
    ("2.1 성과 지표", "6", 1),
    ("2.2 개선 과제", "8", 1),
    ("3. 향후 계획", "11", 0),
    ("4. 결론", "14", 0),
]
for text, page, indent in items:
    prefix = "  " * indent * 2
    # 점끌기 흉내: text + 공백으로 패딩 + 페이지 번호
    line_text = prefix + text
    dots = " " + "." * max(1, 55 - len(line_text) - len(page)) + " "
    if indent == 0:
        app.set_charshape(bold=True)
    else:
        app.set_charshape(bold=False)
    app.insert_text(line_text + dots + page + "\n")

app.set_charshape(bold=False)
app.insert_paragraph_break()
app.set_parashape(align_type=3)
app.insert_text("— 목차 —\n")
"""

# page_numbers 조합
DEMOS["demo_page_numbers"] = r"""
app.set_charshape(bold=True, height=1400)
app.insert_text("page_numbers + header_filename\n\n")
app.set_charshape(bold=False, height=1000)

# 본문
app.insert_text("본 문서는 쪽번호 + 머리글 preset 예시.\n\n")
app.insert_text("app.preset.page_numbers(header_filename=True)\n")
app.insert_text("  → 하단 중앙에 'n / N' 형식 쪽번호\n")
app.insert_text("  → 머리글에 파일명 자동 삽입\n\n")

# 가짜 머리글 박스 (시각 예시용)
app.insert_paragraph_break()
app.insert_text("────────────────────────────\n")
app.set_parashape(align_type=3)
app.set_charshape(height=800, text_color="#666666")
app.insert_text("report.hwp\n")
app.set_parashape(align_type=1)
app.set_charshape(text_color="#000000", height=1000)
app.insert_text("────────────────────────────\n\n")

# 가짜 푸터
app.insert_text("\n\n\n")
app.set_parashape(align_type=3)
app.set_charshape(height=800, text_color="#666666")
app.insert_text("— 1 / 3 —\n")
"""

# 컬러 보고서 전체 조합 레시피
DEMOS["demo_color_report_recipe"] = r"""
# 레시피: 컬러 보고서 스타일 전체 조합

# 1. 타이틀 박스
app.preset.title_box(
    text="프로젝트 성과 보고",
    subtitle="개발팀",
    bg_color="#003366",
    title_color="#FFFFFF",
    font_size=1600,
)
app.insert_paragraph_break()

# 2. 소제목 + 본문
app.preset.subtitle_bar("1. 핵심 지표", bg_color="#EEEEEE")
app.insert_text("\n")

# 3. 데이터 표 (striped + sky header)
data = [["지표", "목표", "실적", "달성률"],
        ["매출", "100억", "118억", "118%"],
        ["신규 고객", "50", "62", "124%"],
        ["유지율", "85%", "91%", "107%"]]
app.insert_table(data=data, headers=None)
# cursor 는 첫 셀에 있음
app.preset.table_header(color="sky", text_color="#FFFFFF")
app.preset.striped_rows(
    colors=["#FFFFFF", "#F5F5F5"],
    skip_header=True,
)
app.move.bottom_of_file()
app.insert_paragraph_break()

# 4. 요약 박스
app.preset.summary_box(
    "모든 핵심 지표가 목표치를 초과 달성.",
    bg_color="#E8F4F8",
)
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


# Batch worker — processes ALL demos in a SINGLE HWP session.
# Between demos: FileClose (not quit!) + FileNew. Much faster than
# spawning a fresh HWP per demo.
BATCH_WORKER = r'''
import sys, time, os, tempfile, json

# argv[1] = JSON: {demo_name: code}, argv[2] = output dir
with open(sys.argv[1], "r", encoding="utf-8") as f:
    demos = json.load(f)
out_dir = sys.argv[2]

import fitz
from hwpapi.core import App
# Phase 4 (v2) scrubbed ``hwpapi.classes`` — char/para state reset below is
# a no-op instead of calling ``app.set_charshape(CharShape())``.

app = App(is_visible=False)
time.sleep(0.5)

# Silence all dialogs — so FileClose doesn't prompt "save?" etc.
try:
    app.set_message_box_mode(App.SILENCE_ALL_NO)   # "No" to save prompts
except Exception:
    pass

# Monkey-patch insert_text so "\n" splits into paragraphs
_orig_insert = app.insert_text
def _insert_text(text, *a, **kw):
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if part:
            _orig_insert(part, *a, **kw)
        if i < len(parts) - 1:
            app.api.Run("BreakPara")
    return app
app.insert_text = _insert_text

results = {}

def close_all_docs():
    """
    Close ALL open documents via the COM-level Documents.close_all API.
    This bypasses the action system entirely — no save dialog can pop up.
    """
    try:
        app.documents.close_all(save=False)
    except Exception:
        pass


for idx, (name, code) in enumerate(demos.items()):
    print(f"[{idx+1}/{len(demos)}] {name}", flush=True)

    # Close ALL existing docs (including PDF output from previous iter)
    close_all_docs()
    time.sleep(0.3)

    # Fresh blank doc for this demo
    try:
        app.api.Run("FileNew")
    except Exception:
        pass
    time.sleep(0.4)

    # Reset char/para state — v2 removed set_charshape/set_parashape;
    # this block is intentionally left as a best-effort no-op.

    # Execute demo code
    try:
        exec(compile(code, f"<demo:{name}>", "exec"),
             {"app": app, "__builtins__": __builtins__})
        time.sleep(0.3)
    except Exception as e:
        print(f"  DEMO ERROR: {e}", flush=True)
        results[name] = f"demo error: {e}"
        continue

    # Save as PDF → PNG
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_path = tmp.name
    try:
        app.save(pdf_path)
        time.sleep(0.6)
    except Exception as e:
        results[name] = f"save error: {e}"
        continue

    out_png = os.path.join(out_dir, f"{name}.png")
    if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 500:
        try:
            with fitz.open(pdf_path) as pdf:
                pix = pdf[0].get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                pix.save(out_png)
            os.remove(pdf_path)

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
                    img.crop((max(0, L-pad), max(0, T-pad),
                             min(img.width, R+pad), min(img.height, B+pad))
                            ).save(out_png)
            except Exception:
                pass
            results[name] = "OK"
        except Exception as e:
            results[name] = f"render error: {e}"
    else:
        results[name] = "empty PDF"

# Final cleanup — close everything before quit
try:
    close_all_docs()
    time.sleep(0.3)
except Exception:
    pass
try:
    app.quit()
except Exception:
    pass

# Summary
print("\n=== Results ===", flush=True)
for name, status in results.items():
    print(f"  {name:35s} {status}", flush=True)
'''


def main():
    print(f"📁 Output: {IMG_DIR}")

    # Write batch worker + demos JSON
    worker_path = os.path.join(tempfile.gettempdir(), "hwpapi_batch_worker.py")
    with open(worker_path, "w", encoding="utf-8") as f:
        f.write(BATCH_WORKER)

    demos_json = os.path.join(tempfile.gettempdir(), "hwpapi_demos.json")
    import json
    with open(demos_json, "w", encoding="utf-8") as f:
        json.dump(DEMOS, f, ensure_ascii=False)

    # Run entire batch in ONE HWP session
    try:
        result = subprocess.run(
            [sys.executable, worker_path, demos_json, IMG_DIR],
            encoding="utf-8",
            errors="replace",
            timeout=600,   # allow 10 min for whole batch
        )
        if result.returncode != 0:
            print(f"⚠ batch worker exited with code {result.returncode}")
    except Exception as e:
        print(f"⚠ {e}")

    print(f"\n✅ Done — images in {IMG_DIR}")


if __name__ == "__main__":
    main()
