"""
Tutorial inline image generation helpers.

Runs HWP demos in subprocess, exports to PDF, converts to PNG, and
returns an IPython Image for display inside the tutorial notebook.

Used by ``#| echo: false`` cells in the qmd-rendered ipynb tutorials
so the image-generation code is hidden and only the resulting image
appears in the final HTML.

Requires Windows + HWP installed.
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional


# Module-level config
HERE = Path(__file__).parent
IMG_DIR = HERE / "img"
IMG_DIR.mkdir(exist_ok=True)


WORKER = r'''
"""Subprocess worker that runs one HWP demo and saves the result as PNG."""
import sys, os, tempfile, time

code_path, out_png = sys.argv[1], sys.argv[2]
with open(code_path, "r", encoding="utf-8") as f:
    demo_code = f.read()

import fitz  # PyMuPDF
from hwpapi.core import App
from hwpapi.classes.shapes import CharShape, ParaShape

app = App(is_visible=False)

# Start in an empty fresh document
try:
    app.api.Run("FileNew")
except Exception:
    pass
time.sleep(0.6)

# Clear any stale formatting state
try:
    app.set_charshape(CharShape())
except Exception:
    pass
try:
    app.set_parashape(ParaShape())
except Exception:
    pass
time.sleep(0.3)

# Monkey-patch insert_text so "\n" creates real paragraph breaks
_orig_insert = app.insert_text
def _insert_text(text, *a, **kw):
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if part:
            _orig_insert(part, *a, **kw)
        if i < len(parts) - 1:
            app.api.Run("BreakPara")
app.insert_text = _insert_text

# Run the demo code
try:
    exec(compile(demo_code, "<demo>", "exec"),
         {"app": app, "__builtins__": __builtins__})
    time.sleep(0.4)
except Exception as e:
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Save as PDF
with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
    pdf_path = tmp.name
try:
    app.save(pdf_path)
except Exception as e:
    print(f"save failed: {e}")
    sys.exit(1)
time.sleep(0.8)

# PDF → PNG (first page, 2x resolution, auto-crop)
if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 500:
    with fitz.open(pdf_path) as pdf:
        page = pdf[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        pix.save(out_png)
    try:
        os.remove(pdf_path)
    except Exception:
        pass

    # Auto-crop whitespace
    try:
        from PIL import Image, ImageChops
        img = Image.open(out_png)
        bg = Image.new(img.mode, img.size, (255, 255, 255))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox:
            L, T, R, B = bbox
            pad = 30
            L = max(0, L - pad); T = max(0, T - pad)
            R = min(img.width, R + pad); B = min(img.height, B + pad)
            img.crop((L, T, R, B)).save(out_png)
    except Exception:
        pass

    print(f"OK {out_png}")
    sys.exit(0)
else:
    print(f"FAIL: no PDF saved or empty")
    sys.exit(1)
'''


def run_and_show(name: str, demo_code: str,
                  force: bool = False) -> "IPython.display.Image":
    """
    Run a demo in a fresh HWP subprocess and return an IPython Image
    for inline display. Caches by ``name`` so re-running the notebook
    doesn't regenerate unchanged images.

    Parameters
    ----------
    name : str
        Unique identifier — becomes the PNG filename ``demo_{name}.png``.
    demo_code : str
        Python code to execute in the HWP subprocess. ``app`` is available.
    force : bool
        Set True to regenerate even if the image file already exists.

    Returns
    -------
    IPython.display.Image
        The generated PNG ready to display in a notebook cell.

    Examples
    --------
    In a notebook cell (with ``#| echo: false``):

        from tutorial_helpers import run_and_show

        run_and_show("striped_rows", '''
            data = [["지역", "Q1"], ["서울", "820"], ["부산", "450"]]
            app.insert_table(data=data)
            app.move.doc.top()
            app.preset.striped_rows(colors=["#FFFFFF", "#F0F0F0"])
        ''')

    The subprocess:
    1. Launches HWP invisibly
    2. Creates a fresh empty document
    3. Runs the demo code
    4. Saves as PDF → converts to PNG
    5. Returns the PNG path for display
    """
    from IPython.display import Image

    out_png = IMG_DIR / f"demo_{name}.png"

    # Skip regeneration if cached
    if out_png.exists() and not force:
        return Image(str(out_png))

    # Write worker + demo code to temp files
    worker_path = Path(tempfile.gettempdir()) / f"hwpapi_worker_{os.getpid()}.py"
    worker_path.write_text(WORKER, encoding="utf-8")

    code_path = Path(tempfile.gettempdir()) / f"hwpapi_demo_{name}.py"
    code_path.write_text(demo_code, encoding="utf-8")

    # Run subprocess
    try:
        result = subprocess.run(
            [sys.executable, str(worker_path), str(code_path), str(out_png)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
        if result.returncode != 0:
            # Return a placeholder error image or log
            print(f"[run_and_show {name}] FAILED")
            if result.stderr:
                print(result.stderr[-500:])
            return None
    except subprocess.TimeoutExpired:
        print(f"[run_and_show {name}] TIMEOUT")
        return None
    except Exception as e:
        print(f"[run_and_show {name}] {e}")
        return None
    finally:
        try:
            code_path.unlink()
        except Exception:
            pass

    if out_png.exists() and out_png.stat().st_size > 1000:
        return Image(str(out_png))
    return None
