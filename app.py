from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import os, uuid, hashlib, secrets, io
from datetime import datetime, timedelta
import sqlite3

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

UPLOAD_FOLDER = "/tmp/uploads"
OUTPUT_FOLDER = "/tmp/outputs"
DB_PATH = "/tmp/officepdf.db"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            plan TEXT DEFAULT 'free',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            expires_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS file_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            tool TEXT NOT NULL,
            filename TEXT NOT NULL,
            original_size INTEGER DEFAULT 0,
            output_size INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit()
    conn.close()

init_db()

def hash_password(p): return hashlib.sha256(p.encode()).hexdigest()
def get_token(req): return req.headers.get("Authorization","").replace("Bearer ","").strip()

def get_user_from_token(token):
    if not token: return None
    conn = get_db()
    row = conn.execute(
        "SELECT u.* FROM users u JOIN sessions s ON u.id=s.user_id WHERE s.token=? AND s.expires_at>?",
        (token, datetime.now().isoformat())
    ).fetchone()
    conn.close()
    return dict(row) if row else None

def save_history(user_id, tool, filename, orig, out):
    try:
        conn = get_db()
        conn.execute("INSERT INTO file_history(user_id,tool,filename,original_size,output_size) VALUES(?,?,?,?,?)",
                     (user_id, tool, filename, orig, out))
        conn.commit(); conn.close()
    except: pass

def save_upload(file):
    uid = str(uuid.uuid4())[:8]
    safe = uid + "_" + "".join(c for c in file.filename if c.isalnum() or c in "._- ")
    path = os.path.join(UPLOAD_FOLDER, safe)
    file.save(path)
    return path, safe

def out_path(suffix):
    return os.path.join(OUTPUT_FOLDER, str(uuid.uuid4())[:8] + suffix)

@app.route("/")
def index(): return send_from_directory("static","index.html")

@app.route("/api/status")
def status(): return jsonify({"status":"OfficePDFEditor API running","version":"3.1"})

@app.route("/api/signup", methods=["POST"])
def signup():
    d = request.json or {}
    name=d.get("name","").strip(); email=d.get("email","").strip().lower(); password=d.get("password","")
    if not name or not email or not password: return jsonify({"error":"Sabhi fields bharein"}),400
    if len(password)<6: return jsonify({"error":"Password kam se kam 6 characters ka hona chahiye"}),400
    conn=get_db()
    if conn.execute("SELECT id FROM users WHERE email=?",(email,)).fetchone():
        conn.close(); return jsonify({"error":"Yeh email already registered hai"}),400
    conn.execute("INSERT INTO users(name,email,password) VALUES(?,?,?)",(name,email,hash_password(password)))
    conn.commit()
    user=dict(conn.execute("SELECT * FROM users WHERE email=?",(email,)).fetchone())
    token=secrets.token_hex(32)
    expires=(datetime.now()+timedelta(days=30)).isoformat()
    conn.execute("INSERT INTO sessions(token,user_id,expires_at) VALUES(?,?,?)",(token,user["id"],expires))
    conn.commit(); conn.close()
    return jsonify({"token":token,"user":{"id":user["id"],"name":name,"email":email,"plan":"free"}})

@app.route("/api/login", methods=["POST"])
def login():
    d=request.json or {}
    email=d.get("email","").strip().lower(); password=d.get("password","")
    conn=get_db()
    row=conn.execute("SELECT * FROM users WHERE email=? AND password=?",(email,hash_password(password))).fetchone()
    if not row: conn.close(); return jsonify({"error":"Email ya password galat hai"}),401
    user=dict(row); token=secrets.token_hex(32)
    expires=(datetime.now()+timedelta(days=30)).isoformat()
    conn.execute("INSERT INTO sessions(token,user_id,expires_at) VALUES(?,?,?)",(token,user["id"],expires))
    conn.commit(); conn.close()
    return jsonify({"token":token,"user":{"id":user["id"],"name":user["name"],"email":user["email"],"plan":user["plan"]}})

@app.route("/api/logout", methods=["POST"])
def logout():
    conn=get_db(); conn.execute("DELETE FROM sessions WHERE token=?",(get_token(request),)); conn.commit(); conn.close()
    return jsonify({"message":"Logged out"})

@app.route("/api/me")
def me():
    user=get_user_from_token(get_token(request))
    if not user: return jsonify({"error":"Login karein"}),401
    return jsonify({"id":user["id"],"name":user["name"],"email":user["email"],"plan":user["plan"]})

@app.route("/api/history")
def history():
    user=get_user_from_token(get_token(request))
    if not user: return jsonify([])
    conn=get_db()
    rows=conn.execute("SELECT * FROM file_history WHERE user_id=? ORDER BY created_at DESC LIMIT 20",(user["id"],)).fetchall()
    conn.close(); return jsonify([dict(r) for r in rows])

# ── COMPRESS ─────────────────────────────────────────────────
@app.route("/api/compress", methods=["POST"])
def compress():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    quality=request.form.get("quality","balanced")
    try:
        import pypdf
        input_path,fname=save_upload(file)
        orig_size=os.path.getsize(input_path)
        
        # Simple, safe compress — just copy pages cleanly
        reader=pypdf.PdfReader(input_path)
        writer=pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        
        # compress_identical_objects is safe
        writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
        
        output=out_path("_compressed.pdf")
        with open(output,"wb") as f:
            writer.write(f)
        
        out_size=os.path.getsize(output)
        # If output bigger than input, just return original
        if out_size >= orig_size:
            import shutil
            shutil.copy(input_path, output)
            out_size=orig_size
        
        user=get_user_from_token(get_token(request))
        if user: save_history(user["id"],"Compress PDF",file.filename,orig_size,out_size)
        
        saved=orig_size-out_size
        pct=round((saved/orig_size)*100,1) if orig_size>0 else 0
        return jsonify({"download_id":os.path.basename(output),"original_size":orig_size,"output_size":out_size,"saved":saved,"percent":pct})
    except Exception as e:
        return jsonify({"error":"Compress error: "+str(e)}),500

# ── MERGE ────────────────────────────────────────────────────
@app.route("/api/merge", methods=["POST"])
def merge():
    files=request.files.getlist("files")
    if len(files)<2: return jsonify({"error":"Kam se kam 2 PDF files upload karein"}),400
    try:
        import pypdf
        writer=pypdf.PdfWriter()
        total_size=0
        for file in files:
            path,_=save_upload(file)
            total_size+=os.path.getsize(path)
            reader=pypdf.PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
        output=out_path("_merged.pdf")
        with open(output,"wb") as f:
            writer.write(f)
        user=get_user_from_token(get_token(request))
        if user: save_history(user["id"],"Merge PDF",f"{len(files)} files",total_size,os.path.getsize(output))
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Merge error: "+str(e)}),500

# ── SPLIT ────────────────────────────────────────────────────
@app.route("/api/split", methods=["POST"])
def split():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    split_type=request.form.get("split_type","all")
    page_range=request.form.get("page_range","").strip()
    try:
        import pypdf, zipfile
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        total=len(reader.pages)
        zip_output=out_path("_split.zip")
        with zipfile.ZipFile(zip_output,"w",zipfile.ZIP_DEFLATED) as zf:
            if split_type=="range" and page_range:
                pages=set()
                for part in page_range.split(","):
                    part=part.strip()
                    if "-" in part:
                        a,b=part.split("-")
                        pages.update(range(int(a)-1,min(int(b),total)))
                    elif part: pages.add(int(part)-1)
                writer=pypdf.PdfWriter()
                for p in sorted(pages):
                    if 0<=p<total: writer.add_page(reader.pages[p])
                buf=io.BytesIO(); writer.write(buf)
                zf.writestr("split_pages.pdf",buf.getvalue())
            else:
                for i,page in enumerate(reader.pages):
                    writer=pypdf.PdfWriter(); writer.add_page(page)
                    buf=io.BytesIO(); writer.write(buf)
                    zf.writestr(f"page_{i+1}.pdf",buf.getvalue())
        user=get_user_from_token(get_token(request))
        if user: save_history(user["id"],"Split PDF",file.filename,os.path.getsize(input_path),0)
        return jsonify({"download_id":os.path.basename(zip_output)})
    except Exception as e:
        return jsonify({"error":"Split error: "+str(e)}),500

# ── ROTATE ───────────────────────────────────────────────────
@app.route("/api/rotate", methods=["POST"])
def rotate():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    angle=int(request.form.get("angle",90))
    try:
        import pypdf
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        writer=pypdf.PdfWriter()
        for page in reader.pages:
            page.rotate(angle)
            writer.add_page(page)
        output=out_path("_rotated.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Rotate error: "+str(e)}),500

# ── DELETE PAGES ─────────────────────────────────────────────
@app.route("/api/delete-pages", methods=["POST"])
def delete_pages():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    pages_str=request.form.get("pages","")
    try:
        import pypdf
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        to_del=set()
        for part in pages_str.split(","):
            part=part.strip()
            if "-" in part:
                a,b=part.split("-"); to_del.update(range(int(a)-1,int(b)))
            elif part: to_del.add(int(part)-1)
        writer=pypdf.PdfWriter()
        for i,page in enumerate(reader.pages):
            if i not in to_del: writer.add_page(page)
        output=out_path("_deleted.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Delete pages error: "+str(e)}),500

# ── JPG TO PDF ───────────────────────────────────────────────
@app.route("/api/jpg-to-pdf", methods=["POST"])
def jpg_to_pdf():
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"Images upload karein"}),400
    try:
        from PIL import Image as PILImage
        images=[]
        for f in files:
            path,_=save_upload(f)
            img=PILImage.open(path).convert("RGB")
            images.append(img)
        output=out_path("_converted.pdf")
        if len(images)==1:
            images[0].save(output,"PDF",resolution=100)
        else:
            images[0].save(output,"PDF",save_all=True,append_images=images[1:],resolution=100)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"JPG to PDF error: "+str(e)}),500

# ── PDF TO JPG ───────────────────────────────────────────────
@app.route("/api/pdf-to-jpg", methods=["POST"])
def pdf_to_jpg():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    try:
        import pypdf, zipfile
        from PIL import Image as PILImage
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        zip_output=out_path("_images.zip")
        count=0
        with zipfile.ZipFile(zip_output,"w",zipfile.ZIP_DEFLATED) as zf:
            for page_num,page in enumerate(reader.pages):
                try:
                    if "/Resources" in page:
                        res=page["/Resources"]
                        if "/XObject" in res:
                            xobj=res["/XObject"]
                            for name in list(xobj.keys()):
                                try:
                                    obj=xobj[name]
                                    if hasattr(obj,'get') and obj.get("/Subtype")=="/Image":
                                        data=obj._data
                                        img=PILImage.open(io.BytesIO(data)).convert("RGB")
                                        buf=io.BytesIO()
                                        img.save(buf,format="JPEG",quality=85)
                                        zf.writestr(f"page{page_num+1}_img{count+1}.jpg",buf.getvalue())
                                        count+=1
                                except: pass
                except: pass
        if count==0:
            return jsonify({"error":"PDF mein extract hone wali images nahi mili. Yeh text-based PDF hai."}),400
        return jsonify({"download_id":os.path.basename(zip_output)})
    except Exception as e:
        return jsonify({"error":"PDF to JPG error: "+str(e)}),500

# ── WATERMARK ────────────────────────────────────────────────
@app.route("/api/watermark", methods=["POST"])
def watermark():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    text=request.form.get("text","CONFIDENTIAL")
    try:
        import pypdf
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        wm_buf=io.BytesIO()
        c=canvas.Canvas(wm_buf,pagesize=A4)
        c.setFont("Helvetica-Bold",52)
        c.setFillColorRGB(0.75,0.75,0.75,alpha=0.25)
        c.saveState(); c.translate(A4[0]/2,A4[1]/2); c.rotate(45)
        c.drawCentredString(0,0,text); c.restoreState(); c.save()
        wm_buf.seek(0)
        wm_page=pypdf.PdfReader(wm_buf).pages[0]
        writer=pypdf.PdfWriter()
        for page in reader.pages:
            page.merge_page(wm_page)
            writer.add_page(page)
        output=out_path("_watermarked.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Watermark error: "+str(e)}),500

# ── ADD PAGE NUMBERS ─────────────────────────────────────────
@app.route("/api/add-page-numbers", methods=["POST"])
def add_page_numbers():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    try:
        import pypdf
        from reportlab.pdfgen import canvas
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        writer=pypdf.PdfWriter()
        for i,page in enumerate(reader.pages):
            w=float(page.mediabox.width); h=float(page.mediabox.height)
            buf=io.BytesIO()
            c=canvas.Canvas(buf,pagesize=(w,h))
            c.setFont("Helvetica",10); c.setFillColorRGB(0.3,0.3,0.3)
            c.drawCentredString(w/2,18,f"- {i+1} -"); c.save(); buf.seek(0)
            overlay=pypdf.PdfReader(buf)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
        output=out_path("_numbered.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Page numbers error: "+str(e)}),500

# ── COUNT WORDS ──────────────────────────────────────────────
@app.route("/api/count-words", methods=["POST"])
def count_words():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    try:
        import pdfplumber
        input_path,fname=save_upload(file)
        text=""
        with pdfplumber.open(input_path) as pdf:
            pages=len(pdf.pages)
            for page in pdf.pages:
                t=page.extract_text()
                if t: text+=t+" "
        words=len(text.split()) if text.strip() else 0
        return jsonify({"words":words,"characters":len(text.replace(" ","")),"pages":pages})
    except Exception as e:
        return jsonify({"error":"Count words error: "+str(e)}),500

# ── PDF INFO ─────────────────────────────────────────────────
@app.route("/api/info", methods=["POST"])
def pdf_info():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    try:
        import pypdf
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        meta=reader.metadata or {}
        return jsonify({"pages":len(reader.pages),"title":str(meta.get("/Title","N/A")),
                        "author":str(meta.get("/Author","N/A")),"encrypted":reader.is_encrypted,
                        "file_size":os.path.getsize(input_path)})
    except Exception as e:
        return jsonify({"error":"PDF info error: "+str(e)}),500

# ── LOCK PDF ─────────────────────────────────────────────────
@app.route("/api/lock", methods=["POST"])
def lock_pdf():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    password=request.form.get("password","")
    if not password: return jsonify({"error":"Password enter karein"}),400
    try:
        import pypdf
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        writer=pypdf.PdfWriter()
        for page in reader.pages: writer.add_page(page)
        writer.encrypt(password)
        output=out_path("_locked.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Lock error: "+str(e)}),500

# ── UNLOCK PDF ───────────────────────────────────────────────
@app.route("/api/unlock", methods=["POST"])
def unlock_pdf():
    if "file" not in request.files: return jsonify({"error":"File upload karein"}),400
    file=request.files["file"]
    password=request.form.get("password","")
    try:
        import pypdf
        input_path,fname=save_upload(file)
        reader=pypdf.PdfReader(input_path)
        if reader.is_encrypted:
            result=reader.decrypt(password)
            if not result: return jsonify({"error":"Password galat hai"}),400
        writer=pypdf.PdfWriter()
        for page in reader.pages: writer.add_page(page)
        output=out_path("_unlocked.pdf")
        with open(output,"wb") as f: writer.write(f)
        return jsonify({"download_id":os.path.basename(output)})
    except Exception as e:
        return jsonify({"error":"Unlock error: "+str(e)}),500

# ── DOWNLOAD ─────────────────────────────────────────────────
@app.route("/api/download/<file_id>")
def download(file_id):
    if ".." in file_id or "/" in file_id or "\\" in file_id:
        return jsonify({"error":"Invalid"}),400
    path=os.path.join(OUTPUT_FOLDER,file_id)
    if not os.path.exists(path):
        return jsonify({"error":"File nahi mili ya expire ho gayi"}),404
    # Detect mime type
    if file_id.endswith(".zip"):
        mime="application/zip"
        dl_name=file_id
    else:
        mime="application/pdf"
        dl_name=file_id if file_id.endswith(".pdf") else file_id+".pdf"
    return send_file(path, mimetype=mime, as_attachment=True, download_name=dl_name)

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port,debug=False)
