# OfficePDFEditor – Real PDF Processing Backend

A complete PDF tools web application with real processing.
**Backend:** Flask + pypdf + reportlab + pdfplumber  
**Frontend:** Single-page HTML, no framework needed

---

## 📁 Project Structure

```
officepdfeditor/
├── app.py               ← Flask backend (all 16 PDF APIs)
├── requirements.txt     ← Python dependencies
├── README.md
├── static/
│   └── index.html       ← Full frontend (served by Flask)
├── uploads/             ← Temp uploads (auto-created)
└── outputs/             ← Processed files (auto-created)
```

---

## ⚙️ Setup & Run

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Install system tools (for PDF→JPG conversion)

**Ubuntu / Debian:**
```bash
sudo apt-get install poppler-utils qpdf ghostscript
```

**macOS (Homebrew):**
```bash
brew install poppler qpdf ghostscript
```

**Windows:**
- Install Ghostscript: https://www.ghostscript.com/download/gsdnld.html
- Install poppler: https://github.com/oschwartz10612/poppler-windows

### 3. Run the server

```bash
python app.py
```

Server starts at: **http://localhost:5000**

---

## 🔌 API Endpoints

| Method | Endpoint              | Description                        |
|--------|-----------------------|------------------------------------|
| POST   | `/api/info`           | Get PDF info (pages, size, meta)   |
| POST   | `/api/compress`       | Compress PDF to target KB size     |
| POST   | `/api/merge`          | Merge multiple PDFs into one       |
| POST   | `/api/split`          | Split PDF (all/range/every N)      |
| POST   | `/api/rotate`         | Rotate pages (all/odd/even/specific)|
| POST   | `/api/delete-pages`   | Delete specified pages             |
| POST   | `/api/extract-pages`  | Extract specific pages             |
| POST   | `/api/rearrange`      | Reorder pages                      |
| POST   | `/api/watermark`      | Add text watermark                 |
| POST   | `/api/add-page-numbers`| Stamp page numbers                |
| POST   | `/api/unlock`         | Remove PDF password                |
| POST   | `/api/lock`           | Add password to PDF                |
| POST   | `/api/count-words`    | Count words/chars/sentences        |
| POST   | `/api/pdf-to-jpg`     | Convert pages to JPG/PNG (ZIP)     |
| POST   | `/api/jpg-to-pdf`     | Convert images to PDF              |
| POST   | `/api/increase-size`  | Pad PDF to minimum KB              |
| POST   | `/api/resize-page`    | Change page dimensions (A4/Letter) |

---

## 📦 API Usage Examples

### Compress PDF
```bash
curl -X POST http://localhost:5000/api/compress \
  -F "file=@document.pdf" \
  -F "target_kb=200" \
  -F "quality=balanced" \
  --output compressed.pdf
```

### Merge PDFs
```bash
curl -X POST http://localhost:5000/api/merge \
  -F "files=@file1.pdf" \
  -F "files=@file2.pdf" \
  -F "files=@file3.pdf" \
  --output merged.pdf
```

### Split PDF
```bash
curl -X POST http://localhost:5000/api/split \
  -F "file=@document.pdf" \
  -F "mode=range" \
  -F "from_page=1" \
  -F "to_page=5" \
  --output split.zip
```

### Add Watermark
```bash
curl -X POST http://localhost:5000/api/watermark \
  -F "file=@document.pdf" \
  -F "text=CONFIDENTIAL" \
  -F "opacity=0.25" \
  -F "angle=45" \
  --output watermarked.pdf
```

### Count Words
```bash
curl -X POST http://localhost:5000/api/count-words \
  -F "file=@document.pdf"
# Returns JSON: {"ok":true,"words":1234,"characters":6789,...}
```

### Lock PDF
```bash
curl -X POST http://localhost:5000/api/lock \
  -F "file=@document.pdf" \
  -F "password=mypassword123" \
  --output locked.pdf
```

---

## 🚀 Deploy to Production

### Option A: Gunicorn (Linux)
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Option B: Docker
```dockerfile
FROM python:3.12-slim
RUN apt-get update && apt-get install -y poppler-utils qpdf ghostscript
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 5000
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

Build and run:
```bash
docker build -t officepdfeditor .
docker run -p 5000:5000 officepdfeditor
```

### Option C: Railway / Render / Fly.io
1. Push to GitHub
2. Connect repository to Railway/Render
3. Set start command: `gunicorn app:app`
4. Done ✅

---

## 🔒 Security Notes

- Files are processed in temp directories and deleted after each request
- Max file size: 150 MB
- CORS enabled for all origins (restrict in production)
- No authentication required for API (add middleware for production)

---

## 🛠️ Tech Stack

| Layer     | Technology                    |
|-----------|-------------------------------|
| Backend   | Flask 3.x                     |
| PDF Read  | pypdf 5.x                     |
| PDF Write | pypdf + reportlab             |
| Tables    | pdfplumber                    |
| Images    | Pillow                        |
| Compress  | qpdf + Ghostscript            |
| PDF→IMG   | poppler (pdftoppm)            |
| Frontend  | Vanilla HTML/CSS/JS           |
