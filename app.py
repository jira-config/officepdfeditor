"""
OfficePDFEditor Backend — Flask + pypdf + reportlab + pdfplumber + qpdf
Real PDF processing: compress, merge, split, rotate, watermark,
page-numbers, unlock, lock, word-count, extract-pages, delete-pages
"""

import os, io, re, json, uuid, zipfile, subprocess, shutil, tempfile
from pathlib import Path
from datetime import datetime

from flask import (
    Flask, request, send_file, jsonify,
    render_template_string, send_from_directory
)

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError
import pdfplumber
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import Color
from PIL import Image

# Fixed static folder config for Railway
app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024   # 150 MB

UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ─── helpers ────────────────────────────────────────────────

def cors(response):
    response.headers["Access-Control-Allow-Origin"]  = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    return response

@app.after_request
def after(r): return cors(r)

# ─── FRONTEND SERVING (FIXED) ───────────────────────────────

@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_frontend(path):
    static_dir = Path(__file__).parent / "static"
    target = static_dir / path
    if path and target.exists():
        return send_from_directory(str(static_dir), path)
    
    # Hamesha index.html bhejega agar koi route match na ho
    return send_from_directory(str(static_dir), "index.html")

@app.route("/api/status", methods=["GET"])
def api_status():
    return jsonify({"status": "OfficePDFEditor API running", "version": "1.0"}), 200

# ────────────────────────────────────────────────────────────

def save_upload(file_storage) -> Path:
    uid = uuid.uuid4().hex[:12]
    ext = Path(file_storage.filename).suffix.lower() or ".pdf"
    dest = UPLOAD_DIR / f"{uid}{ext}"
    file_storage.save(str(dest))
    return dest

def out_path(name: str) -> Path:
    uid = uuid.uuid4().hex[:8]
    return OUTPUT_DIR / f"{uid}_{name}"

def file_size_kb(p: Path) -> float:
    return round(p.stat().st_size / 1024, 1)

def err(msg, code=400):
    return jsonify({"ok": False, "error": msg}), code

def ok_file(path: Path, download_name: str, mimetype="application/pdf"):
    orig_kb  = file_size_kb(path)          # may be used in header
    response = send_file(
        str(path),
        mimetype=mimetype,
        as_attachment=True,
        download_name=download_name
    )
    response.headers["X-File-Size-KB"] = str(orig_kb)
    return cors(response)

# ─── /api/info ──────────────────────────────────────────────

@app.route("/api/info", methods=["POST", "OPTIONS"])
def api_info():
    if request.method == "OPTIONS": return jsonify({}), 200
    f = request.files.get("file")
    if not f: return err("No file uploaded")
    tmp = save_upload(f)
    try:
        reader = PdfReader(str(tmp))
        encrypted = reader.is_encrypted
        if encrypted:
            return jsonify({"ok": True, "pages": "?", "encrypted": True,
                            "size_kb": file_size_kb(tmp)})
        pages  = len(reader.pages)
        meta   = reader.metadata or {}
        w = reader.pages[0].mediabox.width  if pages else 0
        h = reader.pages[0].mediabox.height if pages else 0
        return jsonify({
            "ok": True,
            "pages": pages,
            "encrypted": encrypted,
            "size_kb": file_size_kb(tmp),
            "title": meta.get("/Title",""),
            "author": meta.get("/Author",""),
            "width_pt": float(w), "height_pt": float(h)
        })
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/compress ──────────────────────────────────────────

@app.route("/api/compress", methods=["POST", "OPTIONS"])
def api_compress():
    if request.method == "OPTIONS": return jsonify({}), 200
    f = request.files.get("file")
    if not f: return err("No file")

    target_kb = float(request.form.get("target_kb", 200))
    quality   = request.form.get("quality", "balanced")   # high | balanced | low

    tmp  = save_upload(f)
    orig_kb = file_size_kb(tmp)
    out  = out_path("compressed.pdf")

    try:
        # ---- Strategy 1: qpdf linearize + compress-streams ----
        subprocess.run([
            "qpdf", "--linearize", "--compress-streams=y",
            "--stream-data=compress", "--recompress-flate",
            "--compression-level=9",
            str(tmp), str(out)
        ], check=True, capture_output=True)

        new_kb = file_size_kb(out)

        # ---- Strategy 2: pypdf remove duplicate objects ----
        if new_kb > target_kb * 1.05:
            reader = PdfReader(str(out if out.exists() else tmp))
            writer = PdfWriter()
            writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)
            with open(str(out), "wb") as fh:
                writer.write(fh)
            new_kb = file_size_kb(out)

        # ---- Strategy 3: reduce image DPI via Ghostscript if available ----
        if new_kb > target_kb * 1.1 and shutil.which("gs"):
            dpi_map = {"high": 100, "balanced": 72, "low": 50}
            dpi = dpi_map.get(quality, 72)
            gs_out = out_path("gs_compressed.pdf")
            result = subprocess.run([
                "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS=/screen",
                f"-dColorImageResolution={dpi}",
                f"-dGrayImageResolution={dpi}",
                f"-dMonoImageResolution={dpi}",
                "-dNOPAUSE", "-dBATCH", "-dQUIET",
                f"-sOutputFile={gs_out}", str(out)
            ], capture_output=True)
            if result.returncode == 0 and gs_out.exists():
                out.unlink(missing_ok=True)
                out = gs_out
                new_kb = file_size_kb(out)

        saved_pct = round((1 - new_kb / orig_kb) * 100, 1) if orig_kb > 0 else 0
        response  = send_file(str(out), mimetype="application/pdf",
                              as_attachment=True, download_name="compressed.pdf")
        response.headers["X-Orig-KB"]   = str(orig_kb)
        response.headers["X-New-KB"]    = str(new_kb)
        response.headers["X-Saved-Pct"] = str(saved_pct)
        return cors(response)

    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/merge ─────────────────────────────────────────────

@app.route("/api/merge", methods=["POST", "OPTIONS"])
def api_merge():
    if request.method == "OPTIONS": return jsonify({}), 200
    files = request.files.getlist("files")
    if len(files) < 2: return err("Upload at least 2 PDF files")

    out    = out_path("merged.pdf")
    writer = PdfWriter()
    tmps   = []
    try:
        for fobj in files:
            tmp = save_upload(fobj)
            tmps.append(tmp)
            reader = PdfReader(str(tmp))
            for page in reader.pages:
                writer.add_page(page)
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "merged.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        for t in tmps: t.unlink(missing_ok=True)

# ─── /api/split ─────────────────────────────────────────────

@app.route("/api/split", methods=["POST", "OPTIONS"])
def api_split():
    if request.method == "OPTIONS": return jsonify({}), 200
    f = request.files.get("file")
    if not f: return err("No file")

    mode     = request.form.get("mode", "all")   # all | range | every_n
    from_pg  = int(request.form.get("from_page", 1))
    to_pg    = int(request.form.get("to_page", 999))
    every_n  = int(request.form.get("every_n", 1))

    tmp = save_upload(f)
    try:
        reader     = PdfReader(str(tmp))
        total_pgs  = len(reader.pages)
        zip_buf    = io.BytesIO()

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if mode == "range":
                fp = max(1, from_pg) - 1
                tp = min(total_pgs, to_pg)
                writer = PdfWriter()
                for i in range(fp, tp):
                    writer.add_page(reader.pages[i])
                buf = io.BytesIO()
                writer.write(buf)
                zf.writestr(f"pages_{from_pg}_to_{to_pg}.pdf", buf.getvalue())

            elif mode == "every_n":
                n = max(1, every_n)
                chunk = 0
                for start in range(0, total_pgs, n):
                    chunk += 1
                    writer = PdfWriter()
                    for i in range(start, min(start + n, total_pgs)):
                        writer.add_page(reader.pages[i])
                    buf = io.BytesIO()
                    writer.write(buf)
                    end = min(start + n, total_pgs)
                    zf.writestr(f"chunk_{chunk}_pages_{start+1}_to_{end}.pdf",
                                buf.getvalue())
            else:   # all
                for i, page in enumerate(reader.pages):
                    writer = PdfWriter()
                    writer.add_page(page)
                    buf = io.BytesIO()
                    writer.write(buf)
                    zf.writestr(f"page_{i+1:03d}.pdf", buf.getvalue())

        zip_buf.seek(0)
        return cors(send_file(zip_buf, mimetype="application/zip",
                              as_attachment=True, download_name="split_pages.zip"))
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/rotate ────────────────────────────────────────────

@app.route("/api/rotate", methods=["POST", "OPTIONS"])
def api_rotate():
    if request.method == "OPTIONS": return jsonify({}), 200
    f = request.files.get("file")
    if not f: return err("No file")

    angle     = int(request.form.get("angle", 90))
    apply_to  = request.form.get("apply_to", "all")   # all | odd | even | specific
    pages_str = request.form.get("pages", "")          # e.g. "1,3,5-7"

    tmp = save_upload(f)
    out = out_path("rotated.pdf")
    try:
        reader = PdfReader(str(tmp))
        writer = PdfWriter()
        total  = len(reader.pages)

        def should_rotate(i: int) -> bool:
            if apply_to == "all":   return True
            if apply_to == "odd":   return (i + 1) % 2 == 1
            if apply_to == "even":  return (i + 1) % 2 == 0
            if apply_to == "specific":
                return _page_in_list(i + 1, pages_str, total)
            return True

        for i, page in enumerate(reader.pages):
            if should_rotate(i):
                page.rotate(angle)
            writer.add_page(page)

        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "rotated.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

def _page_in_list(pg: int, spec: str, total: int) -> bool:
    """Parse page spec like '1,3,5-7' → set of ints"""
    result = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            try: result.update(range(int(a), int(b)+1))
            except: pass
        else:
            try: result.add(int(part))
            except: pass
    return pg in result

# ─── /api/delete-pages ──────────────────────────────────────

@app.route("/api/delete-pages", methods=["POST", "OPTIONS"])
def api_delete_pages():
    if request.method == "OPTIONS": return jsonify({}), 200
    f         = request.files.get("file")
    pages_str = request.form.get("pages", "")
    if not f:         return err("No file")
    if not pages_str: return err("No pages specified")

    tmp = save_upload(f)
    out = out_path("pages_deleted.pdf")
    try:
        reader  = PdfReader(str(tmp))
        total   = len(reader.pages)
        to_del  = _page_in_list_set(pages_str, total)
        writer  = PdfWriter()
        for i, page in enumerate(reader.pages):
            if (i + 1) not in to_del:
                writer.add_page(page)
        if len(writer.pages) == 0:
            return err("Cannot delete all pages")
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "pages_deleted.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

def _page_in_list_set(spec: str, total: int) -> set:
    result = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            try: result.update(range(int(a), int(b)+1))
            except: pass
        else:
            try: result.add(int(part))
            except: pass
    return result

# ─── /api/extract-pages ─────────────────────────────────────

@app.route("/api/extract-pages", methods=["POST", "OPTIONS"])
def api_extract_pages():
    if request.method == "OPTIONS": return jsonify({}), 200
    f         = request.files.get("file")
    pages_str = request.form.get("pages", "")
    if not f:         return err("No file")
    if not pages_str: return err("Specify pages to extract")

    tmp = save_upload(f)
    out = out_path("extracted.pdf")
    try:
        reader = PdfReader(str(tmp))
        total  = len(reader.pages)
        wanted = _page_in_list_set(pages_str, total)
        writer = PdfWriter()
        for i, page in enumerate(reader.pages):
            if (i + 1) in wanted:
                writer.add_page(page)
        if len(writer.pages) == 0:
            return err("No valid pages selected")
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "extracted.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/watermark ─────────────────────────────────────────

@app.route("/api/watermark", methods=["POST", "OPTIONS"])
def api_watermark():
    if request.method == "OPTIONS": return jsonify({}), 200
    f        = request.files.get("file")
    text     = request.form.get("text", "CONFIDENTIAL")
    opacity  = float(request.form.get("opacity", 0.25))
    angle    = float(request.form.get("angle", 45))
    fontsize = int(request.form.get("fontsize", 48))
    color    = request.form.get("color", "gray")   # gray | red | blue
    if not f: return err("No file")

    tmp = save_upload(f)
    out = out_path("watermarked.pdf")

    # Build watermark PDF in memory
    try:
        reader   = PdfReader(str(tmp))
        page0    = reader.pages[0]
        pg_w     = float(page0.mediabox.width)
        pg_h     = float(page0.mediabox.height)

        wm_buf = io.BytesIO()
        c = rl_canvas.Canvas(wm_buf, pagesize=(pg_w, pg_h))

        color_map = {
            "gray":  Color(0.5, 0.5, 0.5, alpha=opacity),
            "red":   Color(0.8, 0.1, 0.1, alpha=opacity),
            "blue":  Color(0.1, 0.1, 0.8, alpha=opacity),
        }
        c.setFillColor(color_map.get(color, color_map["gray"]))
        c.setFont("Helvetica-Bold", fontsize)
        c.saveState()
        c.translate(pg_w / 2, pg_h / 2)
        c.rotate(angle)
        c.drawCentredString(0, 0, text)
        c.restoreState()
        c.save()

        wm_buf.seek(0)
        wm_reader = PdfReader(wm_buf)
        wm_page   = wm_reader.pages[0]

        writer = PdfWriter()
        for page in reader.pages:
            page.merge_page(wm_page)
            writer.add_page(page)

        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "watermarked.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/add-page-numbers ──────────────────────────────────

@app.route("/api/add-page-numbers", methods=["POST", "OPTIONS"])
def api_add_page_numbers():
    if request.method == "OPTIONS": return jsonify({}), 200
    f        = request.files.get("file")
    position = request.form.get("position", "bottom-center")
    start_at = int(request.form.get("start_at", 1))
    fontsize = int(request.form.get("fontsize", 11))
    fmt      = request.form.get("format", "numeric")   # numeric | roman | "Page N of T"
    if not f: return err("No file")

    tmp = save_upload(f)
    out = out_path("numbered.pdf")
    try:
        reader = PdfReader(str(tmp))
        total  = len(reader.pages)
        writer = PdfWriter()

        for idx, page in enumerate(reader.pages):
            pg_w = float(page.mediabox.width)
            pg_h = float(page.mediabox.height)
            num  = idx + start_at

            if   fmt == "roman":       label = _to_roman(num)
            elif fmt == "page_of":     label = f"Page {num} of {total}"
            elif fmt == "dash":        label = f"— {num} —"
            else:                      label = str(num)

            overlay_buf = io.BytesIO()
            c = rl_canvas.Canvas(overlay_buf, pagesize=(pg_w, pg_h))
            c.setFont("Helvetica", fontsize)
            c.setFillColorRGB(0.2, 0.2, 0.2)

            margin = 28
            pos_map = {
                "bottom-center": (pg_w/2,        margin,        "center"),
                "bottom-right":  (pg_w - margin, margin,        "right"),
                "bottom-left":   (margin,        margin,        "left"),
                "top-center":    (pg_w/2,        pg_h - margin, "center"),
                "top-right":     (pg_w - margin, pg_h - margin, "right"),
                "top-left":      (margin,        pg_h - margin, "left"),
            }
            x, y, align = pos_map.get(position, pos_map["bottom-center"])
            if align == "center": c.drawCentredString(x, y, label)
            elif align == "right": c.drawRightString(x, y, label)
            else:                  c.drawString(x, y, label)
            c.save()

            overlay_buf.seek(0)
            overlay_pg = PdfReader(overlay_buf).pages[0]
            page.merge_page(overlay_pg)
            writer.add_page(page)

        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "numbered.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

def _to_roman(n: int) -> str:
    vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    syms = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    r = ""
    for v, s in zip(vals, syms):
        while n >= v: r += s; n -= v
    return r

# ─── /api/unlock ────────────────────────────────────────────

@app.route("/api/unlock", methods=["POST", "OPTIONS"])
def api_unlock():
    if request.method == "OPTIONS": return jsonify({}), 200
    f        = request.files.get("file")
    password = request.form.get("password", "")
    if not f: return err("No file")

    tmp = save_upload(f)
    out = out_path("unlocked.pdf")
    try:
        reader = PdfReader(str(tmp))
        if reader.is_encrypted:
            success = reader.decrypt(password)
            if not success:
                # Try common blank / empty passwords
                for pw in ["", " ", "password", "123456", "admin"]:
                    try:
                        if reader.decrypt(pw): break
                    except: pass
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.add_metadata(reader.metadata or {})
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "unlocked.pdf")
    except PdfReadError as e:
        return err("Could not unlock: wrong password or unsupported encryption")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/lock ──────────────────────────────────────────────

@app.route("/api/lock", methods=["POST", "OPTIONS"])
def api_lock():
    if request.method == "OPTIONS": return jsonify({}), 200
    f        = request.files.get("file")
    password = request.form.get("password", "")
    if not f:       return err("No file")
    if not password: return err("Password is required")

    tmp = save_upload(f)
    out = out_path("locked.pdf")
    try:
        reader = PdfReader(str(tmp))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(user_password=password, owner_password=password + "_owner")
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "locked.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/count-words ───────────────────────────────────────

@app.route("/api/count-words", methods=["POST", "OPTIONS"])
def api_count_words():
    if request.method == "OPTIONS": return jsonify({}), 200
    f = request.files.get("file")
    if not f: return err("No file")

    tmp = save_upload(f)
    try:
        text = ""
        with pdfplumber.open(str(tmp)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: text += t + "\n"

        words = len(text.split())
        chars = len(text)
        chars_no_space = len(text.replace(" ","").replace("\n",""))
        sentences = len(re.findall(r"[.!?]+", text))
        pages_count = 0
        with pdfplumber.open(str(tmp)) as pdf:
            pages_count = len(pdf.pages)

        return jsonify({
            "ok": True,
            "words": words,
            "characters": chars,
            "characters_no_space": chars_no_space,
            "sentences": sentences,
            "pages": pages_count,
            "avg_words_per_page": round(words / pages_count, 1) if pages_count else 0
        })
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/pdf-to-jpg ────────────────────────────────────────

@app.route("/api/pdf-to-jpg", methods=["POST", "OPTIONS"])
def api_pdf_to_jpg():
    if request.method == "OPTIONS": return jsonify({}), 200
    f      = request.files.get("file")
    dpi    = int(request.form.get("dpi", 150))
    fmt    = request.form.get("format", "jpg").lower()
    if not f: return err("No file")
    if not shutil.which("pdftoppm"):
        return err("pdftoppm not available on this server")

    tmp     = save_upload(f)
    tmp_dir = Path(tempfile.mkdtemp())
    zip_buf = io.BytesIO()

    try:
        prefix = str(tmp_dir / "page")
        fmt_flag = "-jpeg" if fmt == "jpg" else "-png"
        subprocess.run(
            ["pdftoppm", fmt_flag, f"-r", str(dpi), str(tmp), prefix],
            check=True, capture_output=True
        )
        ext    = ".jpg" if fmt == "jpg" else ".png"
        images = sorted(tmp_dir.glob(f"*{ext}"))
        if not images:
            images = sorted(tmp_dir.glob("*.ppm"))

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, img_path in enumerate(images):
                zf.write(str(img_path), f"page_{i+1:03d}{ext}")

        zip_buf.seek(0)
        return cors(send_file(zip_buf, mimetype="application/zip",
                              as_attachment=True, download_name="pdf_images.zip"))
    except subprocess.CalledProcessError as e:
        return err("Image conversion failed: " + (e.stderr.decode() or "unknown"))
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)
        shutil.rmtree(str(tmp_dir), ignore_errors=True)

# ─── /api/jpg-to-pdf ────────────────────────────────────────

@app.route("/api/jpg-to-pdf", methods=["POST", "OPTIONS"])
def api_jpg_to_pdf():
    if request.method == "OPTIONS": return jsonify({}), 200
    files = request.files.getlist("files")
    if not files: return err("No images uploaded")

    out = out_path("images.pdf")
    tmps = []
    try:
        images_pil = []
        for fobj in files:
            tmp = save_upload(fobj)
            tmps.append(tmp)
            img = Image.open(str(tmp)).convert("RGB")
            images_pil.append(img)

        if not images_pil: return err("No valid images")
        images_pil[0].save(
            str(out), save_all=True, append_images=images_pil[1:],
            resolution=150
        )
        return ok_file(out, "images.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        for t in tmps: t.unlink(missing_ok=True)

# ─── /api/rearrange ─────────────────────────────────────────

@app.route("/api/rearrange", methods=["POST", "OPTIONS"])
def api_rearrange():
    if request.method == "OPTIONS": return jsonify({}), 200
    f         = request.files.get("file")
    order_str = request.form.get("order", "")   # "3,1,2,4" — 1-indexed
    if not f:         return err("No file")
    if not order_str: return err("Provide page order")

    tmp = save_upload(f)
    out = out_path("rearranged.pdf")
    try:
        reader = PdfReader(str(tmp))
        total  = len(reader.pages)
        order  = []
        for x in order_str.split(","):
            try:
                n = int(x.strip())
                if 1 <= n <= total:
                    order.append(n - 1)
            except: pass
        if not order: return err("Invalid page order")

        writer = PdfWriter()
        for idx in order:
            writer.add_page(reader.pages[idx])
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "rearranged.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/resize-page ───────────────────────────────────────

@app.route("/api/resize-page", methods=["POST", "OPTIONS"])
def api_resize_page():
    """Change page size to A4, Letter, etc."""
    if request.method == "OPTIONS": return jsonify({}), 200
    f    = request.files.get("file")
    size = request.form.get("size", "A4")    # A4 | Letter | A3
    if not f: return err("No file")

    size_map = {
        "A4":     A4,
        "Letter": letter,
        "A3":     (842, 1191),
        "A5":     (420,  595),
    }
    pagesize = size_map.get(size, A4)

    tmp = save_upload(f)
    out = out_path("resized_page.pdf")
    try:
        reader = PdfReader(str(tmp))
        writer = PdfWriter()
        for page in reader.pages:
            new_page = writer.add_blank_page(width=pagesize[0], height=pagesize[1])
            new_page.merge_page(page)
        with open(str(out), "wb") as fh:
            writer.write(fh)
        return ok_file(out, "resized_page.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── /api/increase-size ─────────────────────────────────────

@app.route("/api/increase-size", methods=["POST", "OPTIONS"])
def api_increase_size():
    """Pad PDF with dummy data to reach a minimum size in KB."""
    if request.method == "OPTIONS": return jsonify({}), 200
    f         = request.files.get("file")
    target_kb = float(request.form.get("target_kb", 500))
    if not f: return err("No file")

    tmp = save_upload(f)
    out = out_path("increased.pdf")
    try:
        reader = PdfReader(str(tmp))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        # Write once, check size
        with open(str(out), "wb") as fh:
            writer.write(fh)

        current_kb = file_size_kb(out)
        if current_kb < target_kb:
            # Append hidden padding via metadata
            needed_bytes = int((target_kb - current_kb) * 1024)
            padding = "X" * needed_bytes
            writer.add_metadata({"/Padding": padding})
            with open(str(out), "wb") as fh:
                writer.write(fh)

        return ok_file(out, "increased.pdf")
    except Exception as e:
        return err(str(e))
    finally:
        tmp.unlink(missing_ok=True)

# ─── main ────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 55)
    print("  OfficePDFEditor Backend  |  http://localhost:5000")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5000, debug=True)