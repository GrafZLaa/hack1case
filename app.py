import os, json, base64, io, re, traceback
import requests
from flask import Flask, render_template, request, jsonify, send_file
import fitz
import barcode
from barcode.writer import ImageWriter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image, ImageEnhance
import pytesseract
import cv2
import numpy as np
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__)

# ── CONFIG ──────────────────────────────────────────────────────────────────
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

OLLAMA_URL   = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434").rstrip("/")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.2-vision")

# ═══════════════════════════════════════════════════════════════════════════
# УНИВЕРСАЛЬНЫЕ ПРОМПТЫ (БЕЗ ПРИМЕРОВ БРЕНДОВ И МОДЕЛЕЙ)
# ═══════════════════════════════════════════════════════════════════════════

# Этап 1: Титульный лист. Убраны все "Мастер-модули" и "КЭАЗы" из текста.
PROMPT_TITLE = """Analyze this equipment passport cover. 
Extract the following information based ONLY on the visual text provided:
- "naimenovanie": The main product name and model found in the largest font. (Remove words like 'Passport', 'Manual', 'RE', 'PS').
- "kod_dokumenta": The decimal document number (pattern: XXXX.XXXXXX.XXXX).
- "proizvoditel": The legal name of the manufacturing company (AO, OOO, ZAO).
- "adres": Manufacturer's address.
- "kontakty": Phone/Email.
- "garantia": Warranty period.
- "srok_sluzhby": Service life.
- "sertifikat": EAC certificate number.
- "kod_zakaza": Order code.
- "normativnye_dok": List of GOST/TU numbers.
- "komplektnost": List of included items.

Return ONLY a JSON object. If a field is not found, use null."""

# Этап 2: Серийные номера и даты
PROMPT_SERIALS = """Scan the text for factory serial numbers and production dates.
- "zavodskie_nomera": An array of serial numbers found (usually near 'Заводской номер' or in tables).
- "data_vypuska": Production date in DD.MM.YYYY format. Look at OTK stamps and hand-written dates.
- "data_priemki": Acceptance date.

Return ONLY a JSON object."""

# ═══════════════════════════════════════════════════════════════════════════
# ОБРАБОТКА ИЗОБРАЖЕНИЙ
# ═══════════════════════════════════════════════════════════════════════════

def improve_image(img_bytes):
    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None: return img_bytes
    # Усиление для мелкого текста и штампов
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_LANCZOS4)
    processed = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 12)
    _, buf = cv2.imencode('.png', processed)
    return buf.tobytes()

# ═══════════════════════════════════════════════════════════════════════════
# ВАЛИДАЦИЯ (ЧИСТКА ГАЛЛЮЦИНАЦИЙ)
# ═══════════════════════════════════════════════════════════════════════════

def final_cleanup(data):
    """Принудительная очистка от мусора на уровне Python"""
    # Список слов, которые точно не являются названием прибора
    junk_words = ["паспорт", "руководство", "эксплуатации", "свидетельство", "приемке", "упаковывании"]
    
    name = str(data.get("naimenovanie") or "").lower()
    for word in junk_words:
        name = name.replace(word, "")
    
    data["naimenovanie"] = name.strip(",. ").capitalize()

    # Если ИИ все же подставил образец из памяти, заменяем на "—"
    hallucinations = ["мастер-модуль", "trei", "трэи", "tcc 8l"]
    if any(h in str(data.get("naimenovanie")).lower() for h in hallucinations):
        data["naimenovanie"] = "—"
        
    return data

# ═══════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/")
def index(): return render_template("index.html")

@app.route("/api/extract", methods=["POST"])
def extract():
    file = request.files.get("file")
    if not file: return jsonify({"error": "No file"}), 400
    try:
        fbytes = file.read()
        # Превращаем PDF в картинки
        if file.filename.lower().endswith(".pdf"):
            doc = fitz.open(stream=fbytes, filetype="pdf")
            pages = [page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5)).tobytes("png") for page in doc]
        else:
            pages = [fbytes]

        # Vision анализ титульника
        img_b64 = [base64.b64encode(pages[0]).decode()]
        title_res = requests.post(f"{OLLAMA_URL}/api/generate", json={
            "model": OLLAMA_MODEL,
            "prompt": PROMPT_TITLE,
            "images": img_b64,
            "stream": False, "format": "json", "options": {"temperature": 0.0}
        }).json()
        title_data = json.loads(title_res.get("response", "{}"))

        # OCR анализ всех страниц для серийников
        full_text = ""
        for p in pages:
            txt = pytesseract.image_to_string(Image.open(io.BytesIO(improve_image(p))), lang='rus+eng')
            full_text += txt + "\n"
        
        serials_res = requests.post(f"{OLLAMA_URL}/api/generate", json={
            "model": OLLAMA_MODEL,
            "prompt": PROMPT_SERIALS + "\n\nTEXT:\n" + full_text,
            "stream": False, "format": "json", "options": {"temperature": 0.0}
        }).json()
        serials_data = json.loads(serials_res.get("response", "{}"))

        # Слияние и очистка
        final_result = final_cleanup({**title_data, **serials_data})
        return jsonify(final_result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/barcode", methods=["POST"])
def api_barcode():
    val = (request.json or {}).get("value", "").strip()
    if not val: return jsonify({"error": "Empty"}), 400
    bc_class = barcode.get_barcode_class("code128")
    buf = io.BytesIO()
    # Транслитерация для сканеров
    clean_val = "".join(re.findall(r'[A-Z0-9\-]', val.upper()))
    bc_class(clean_val[:20] or "0000", writer=ImageWriter()).write(buf)
    return jsonify({"barcode": base64.b64encode(buf.getvalue()).decode(), "encoded_value": clean_val})

@app.route("/api/preview", methods=["POST"])
def preview():
    file = request.files.get("file")
    fbytes = file.read()
    if file.filename.lower().endswith(".pdf"):
        doc = fitz.open(stream=fbytes, filetype="pdf")
        img = doc[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).tobytes("png")
    else: img = fbytes
    return jsonify({"image": base64.b64encode(img).decode()})

@app.route("/api/export/excel", methods=["POST"])
def export_excel():
    records = (request.json or {}).get("records", [])
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Изделие", "Код документа", "S/N", "Завод", "Дата"])
    for r in records:
        for sn in (r.get("zavodskie_nomera") or ["—"]):
            ws.append([r.get("naimenovanie"), r.get("kod_dokumenta"), sn, r.get("proizvoditel"), r.get("data_vypuska")])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, download_name="Registry.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, port=5000)