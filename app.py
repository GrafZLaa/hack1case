import os
import json
import base64
import io
import re
import traceback
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional, Tuple

import requests
from flask import Flask, jsonify, render_template, request, send_file
import fitz
import barcode
from barcode.writer import ImageWriter
import openpyxl
from PIL import Image, ImageEnhance, ImageFilter, ImageOps, ImageChops
import pytesseract
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__)
log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# в”Ђв”Ђ CONFIG в”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђв”Ђ
TESSERACT_PATH = os.getenv("TESSERACT_PATH", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

OLLAMA_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434").rstrip("/")
DEFAULT_MODEL_CANDIDATES = "qwen2.5vl:72b,qwen2.5vl:32b,qwen2.5vl:7b,llama3.2-vision"
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5vl:32b")
MODEL_CANDIDATES = [
    m.strip()
    for m in os.getenv("OLLAMA_MODEL_CANDIDATES", DEFAULT_MODEL_CANDIDATES).split(",")
    if m.strip()
]
MAX_LLM_IMAGES = max(1, int(os.getenv("MAX_LLM_IMAGES", "4")))
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "ollama").strip().lower()
LLM_TEXT_LIMIT = max(2000, int(os.getenv("LLM_TEXT_LIMIT", "22000")))
LLM_RETRY_TEXT_LIMIT = max(1200, int(os.getenv("LLM_RETRY_TEXT_LIMIT", "12000")))
LLM_IMAGE_MAX_SIDE = max(900, int(os.getenv("LLM_IMAGE_MAX_SIDE", "1600")))
LLM_REQUEST_TIMEOUT = max(10, int(os.getenv("LLM_REQUEST_TIMEOUT", "45")))
ENABLE_LLM = os.getenv("ENABLE_LLM", "0").strip().lower() in {"1", "true", "yes"}

OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

OCR_LANG = os.getenv("OCR_LANG", "rus+eng")
OCR_PSMS = ["6", "11", "12"]
OCR_WORKERS = max(1, int(os.getenv("OCR_WORKERS", "4")))
OCR_IF_EMBEDDED_CHARS = max(0, int(os.getenv("OCR_IF_EMBEDDED_CHARS", "120")))
USE_OCR_WITH_OPENAI = os.getenv("USE_OCR_WITH_OPENAI", "0").strip().lower() in {"1", "true", "yes"}
PDF_RENDER_SCALE = max(1.3, float(os.getenv("PDF_RENDER_SCALE", "2.0")))
OCR_MAX_VARIANTS = max(1, int(os.getenv("OCR_MAX_VARIANTS", "3")))
OCR_QUALITY_SHORTCIRCUIT = float(os.getenv("OCR_QUALITY_SHORTCIRCUIT", "320"))
PREVIEW_RENDER_SCALE = max(1.0, float(os.getenv("PREVIEW_RENDER_SCALE", "2.2")))
OCR_KEYWORDS = [
    "паспорт",
    "руководство",
    "наименование",
    "издел",
    "заводск",
    "номер",
    "дата",
    "производител",
    "контроллер",
    "модул",
    "сертификат",
    "гарант",
    "срок",
    "служб",
    "код",
    "заказ",
]

RU_MONTHS = {
    "ЯНВ": "01",
    "ФЕВ": "02",
    "МАР": "03",
    "АПР": "04",
    "МАЙ": "05",
    "ИЮН": "06",
    "ИЮЛ": "07",
    "АВГ": "08",
    "СЕН": "09",
    "ОКТ": "10",
    "НОЯ": "11",
    "ДЕК": "12",
}

CYR_TO_LAT_LOOKALIKE = str.maketrans({
    "А": "A",
    "В": "B",
    "Е": "E",
    "К": "K",
    "М": "M",
    "Н": "H",
    "О": "O",
    "Р": "P",
    "С": "C",
    "Т": "T",
    "У": "Y",
    "Х": "X",
})

SERIAL_CONFUSABLES = {
    "1": "I",
    "I": "1",
    "L": "1",
    "5": "S",
    "S": "5",
    "8": "B",
    "B": "8",
    "6": "G",
    "G": "6",
    "2": "Z",
    "Z": "2",
}

PROMPT_EXTRACT = """You extract equipment-passport data from OCR text and page images.
Rules:
1) Use only text present in provided OCR/page snippets.
2) Never invent manufacturer/model values.
3) Keep Cyrillic exactly as in source where possible.
4) If field is absent, return null (or [] for arrays).

Return strict JSON with fields:
- document_type: one of ["single_passport", "group_passport", "cabinet_list", "unknown"]
- naimenovanie
- kod_dokumenta
- proizvoditel
- adres
- kontakty
- garantia
- srok_sluzhby
- sertifikat
- kod_zakaza
- normativnye_dok (array)
- komplektnost (array)
- zavodskie_nomera (array)
- data_vypuska
- data_priemki
"""


def improve_image_variants(img_bytes: bytes) -> List[bytes]:
    """Build several preprocessing variants to improve OCR hit rate."""
    variants = [img_bytes]
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("L")
        img = ImageOps.autocontrast(img)
        img = img.resize((int(img.width * 1.4), int(img.height * 1.4)), Image.Resampling.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=3))
        sharp = ImageEnhance.Sharpness(img).enhance(2.0)
        high_contrast = ImageEnhance.Contrast(sharp).enhance(1.5)
        low_contrast = ImageEnhance.Contrast(sharp).enhance(1.2)

        for variant in (high_contrast, low_contrast):
            out = io.BytesIO()
            variant.save(out, format="PNG")
            variants.append(out.getvalue())

        for threshold in (165,):
            bw = sharp.point(lambda x: 255 if x > threshold else 0, mode="1").convert("L")
            out = io.BytesIO()
            bw.save(out, format="PNG")
            variants.append(out.getvalue())
    except Exception:
        pass
    return variants


def clean_json(text: str) -> Dict:
    if not text:
        return {}
    text = text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.IGNORECASE | re.DOTALL)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                return {}
    return {}


def normalize_serials(serials: Optional[List[str]]) -> List[str]:
    if not serials:
        return []
    seen = set()
    result = []
    for raw in serials:
        s = re.sub(r"\s+", "", str(raw or "").upper())
        s = s.strip(".,;:-")
        if len(s) < 3:
            continue
        if s not in seen:
            seen.add(s)
            result.append(s)
    return result


def regex_fallbacks(text: str) -> Dict:
    data = {}
    doc_code = _pick_document_code(text)
    if doc_code:
        data["kod_dokumenta"] = doc_code

    data["zavodskie_nomera"] = _extract_serials(text)

    dates = []
    for m in re.finditer(r"\b([0-3]?\d[./][01]?\d[./](?:19|20)\d{2})\b", text):
        s = max(0, m.start() - 30)
        e = min(len(text), m.end() + 30)
        context = text[s:e].lower()
        if "версия" in context:
            continue
        if "дата" not in context:
            continue
        dates.append(m.group(1).replace("/", "."))

    text_date = _try_parse_textual_date(text)
    if text_date and text_date not in dates:
        dates.append(text_date)

    if dates:
        data.setdefault("data_vypuska", dates[0])
        if len(dates) > 1:
            data.setdefault("data_priemki", dates[1])
    return data


def ocr_text_quality(text: str, avg_conf: float = 0.0) -> float:
    if not text:
        return float("-inf")
    words = re.findall(r"[A-Za-zА-Яа-я0-9]{2,}", text)
    if not words:
        return float("-inf")

    lower = text.lower()
    keyword_hits = sum(1 for kw in OCR_KEYWORDS if kw in lower)
    good_chars = len(re.findall(r"[A-Za-zА-Яа-я0-9\s.,:;()\-/+№%\"'«»]", text))
    noisy_chars = max(0, len(text) - good_chars)
    noise_ratio = noisy_chars / max(1, len(text))
    unique_words = len(set(w.lower() for w in words))
    unique_ratio = unique_words / max(1, len(words))
    very_short_words = sum(1 for w in words if len(w) <= 2)
    short_ratio = very_short_words / max(1, len(words))
    avg_word_len = sum(len(w) for w in words) / max(1, len(words))

    score = (
        avg_conf * 2.2
        + keyword_hits * 42
        + unique_ratio * 90
        + avg_word_len * 8
        - short_ratio * 55
        - noise_ratio * 45
    )
    if keyword_hits == 0:
        score -= 85
    return score


def ocr_page_text(page_img: bytes) -> str:
    best = ""
    best_score = float("-inf")
    variants = improve_image_variants(page_img)[:OCR_MAX_VARIANTS]
    for processed in variants:
        image = Image.open(io.BytesIO(processed))
        for psm in OCR_PSMS:
            try:
                txt = pytesseract.image_to_string(
                    image,
                    lang=OCR_LANG,
                    config=f"--oem 1 --psm {psm}",
                )
                score = ocr_text_quality(txt)
                if score > best_score:
                    best = txt
                    best_score = score
            except Exception:
                continue
        if best_score >= OCR_QUALITY_SHORTCIRCUIT:
            break
    return best


def pdf_to_images_and_text(pdf_bytes: bytes) -> Tuple[List[bytes], List[str]]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    raw_text_blocks = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(PDF_RENDER_SCALE, PDF_RENDER_SCALE))
        pages.append(pix.tobytes("png"))
        raw_text_blocks.append(page.get_text("text") or "")
    return pages, raw_text_blocks


def get_available_ollama_models() -> List[str]:
    resp = requests.get(f"{OLLAMA_URL}/api/tags", timeout=20)
    resp.raise_for_status()
    models = resp.json().get("models", [])
    return [m.get("name", "") for m in models if m.get("name")]


def choose_active_ollama_model() -> str:
    try:
        available = get_available_ollama_models()
        if not available:
            return OLLAMA_MODEL
        available_set = {m.lower(): m for m in available}
        ranked = [OLLAMA_MODEL] + MODEL_CANDIDATES
        for candidate in ranked:
            c = candidate.lower()
            if c in available_set:
                return available_set[c]
            if ":" not in c:
                matched = next((available_set[a] for a in available_set if a.startswith(c + ":")), None)
                if matched:
                    return matched
        return available[0]
    except Exception as e:
        log.warning("Failed to resolve Ollama model list: %s", e)
        return OLLAMA_MODEL


def resolve_llm_backend() -> Tuple[str, str]:
    provider = LLM_PROVIDER
    if provider == "openai":
        if OPENAI_API_KEY:
            return "openai", OPENAI_MODEL
        log.warning("LLM_PROVIDER=openai but OPENAI_API_KEY is empty, fallback to Ollama")
    return "ollama", choose_active_ollama_model()


def _response_error_text(resp: requests.Response) -> str:
    try:
        payload = resp.json()
        if isinstance(payload, dict):
            if payload.get("error"):
                return str(payload["error"])
            return json.dumps(payload, ensure_ascii=False)[:600]
    except Exception:
        pass
    return (resp.text or "")[:600]


def prepare_llm_images_b64(pages: List[bytes]) -> List[str]:
    encoded = []
    for page in pages[:MAX_LLM_IMAGES]:
        try:
            img = Image.open(io.BytesIO(page))
            max_side = max(img.size)
            if max_side > LLM_IMAGE_MAX_SIDE:
                scale = LLM_IMAGE_MAX_SIDE / float(max_side)
                new_size = (
                    max(1, int(img.width * scale)),
                    max(1, int(img.height * scale)),
                )
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            out = io.BytesIO()
            img.save(out, format="PNG", optimize=True)
            encoded.append(base64.b64encode(out.getvalue()).decode())
        except Exception:
            encoded.append(base64.b64encode(page).decode())
    return encoded


def run_ollama_extraction(page_images_b64: List[str], ocr_text: str, model_name: str) -> Dict:
    attempts = [
        {
            "images": page_images_b64,
            "ocr_text": ocr_text[:LLM_TEXT_LIMIT],
            "format_json": True,
        },
        {
            "images": page_images_b64[:1],
            "ocr_text": ocr_text[:LLM_RETRY_TEXT_LIMIT],
            "format_json": True,
        },
        {
            "images": page_images_b64[:1],
            "ocr_text": ocr_text[:LLM_RETRY_TEXT_LIMIT],
            "format_json": False,
        },
    ]

    last_error = ""
    for attempt in attempts:
        prompt = f"{PROMPT_EXTRACT}\n\nOCR_TEXT:\n{attempt['ocr_text']}"
        payload = {
            "model": model_name,
            "prompt": prompt,
            "images": attempt["images"],
            "stream": False,
            "options": {"temperature": 0.0},
        }
        if attempt["format_json"]:
            payload["format"] = "json"

        resp = requests.post(f"{OLLAMA_URL}/api/generate", json=payload, timeout=LLM_REQUEST_TIMEOUT)
        if resp.ok:
            raw = resp.json().get("response", "")
            parsed = clean_json(raw)
            if parsed:
                return parsed
            last_error = "empty/invalid JSON from model response"
            continue

        detail = _response_error_text(resp)
        last_error = f"HTTP {resp.status_code}: {detail}"

    raise requests.RequestException(f"Ollama extraction failed: {last_error}")


def run_openai_extraction(page_images_b64: List[str], ocr_text: str, model_name: str) -> Dict:
    if not OPENAI_API_KEY:
        raise requests.RequestException("OPENAI_API_KEY is not set")

    user_content: List[Dict[str, Any]] = [{"type": "text", "text": f"OCR_TEXT:\n{ocr_text[:LLM_TEXT_LIMIT]}"}]
    for image_b64 in page_images_b64:
        user_content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{image_b64}"},
        })

    payload = {
        "model": model_name,
        "temperature": 0,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_EXTRACT + "\nReturn strict JSON only."},
            {"role": "user", "content": user_content},
        ],
    }
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }
    resp = requests.post(f"{OPENAI_BASE_URL}/chat/completions", headers=headers, json=payload, timeout=LLM_REQUEST_TIMEOUT)
    if not resp.ok:
        raise requests.RequestException(f"OpenAI extraction failed: HTTP {resp.status_code}: {_response_error_text(resp)}")
    response_data = resp.json()
    content = (((response_data.get("choices") or [{}])[0].get("message") or {}).get("content", ""))
    if isinstance(content, list):
        content = "".join(part.get("text", "") for part in content if isinstance(part, dict))
    return clean_json(str(content))


def run_llm_extraction(page_images_b64: List[str], ocr_text: str, provider: str, model_name: str) -> Dict:
    if provider == "openai":
        return run_openai_extraction(page_images_b64, ocr_text, model_name)
    return run_ollama_extraction(page_images_b64, ocr_text, model_name)


def collect_ocr_texts(pages: List[bytes], embedded_text_blocks: List[str], enabled: bool = True) -> List[str]:
    if not enabled:
        return ["" for _ in pages]

    results = ["" for _ in pages]
    indexed_tasks = []
    workers = min(OCR_WORKERS, max(1, len(pages)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for idx, page in enumerate(pages):
            embedded = (embedded_text_blocks[idx] if idx < len(embedded_text_blocks) else "").strip()
            if len(embedded) >= OCR_IF_EMBEDDED_CHARS:
                continue
            indexed_tasks.append((idx, pool.submit(ocr_page_text, page)))

        for idx, future in indexed_tasks:
            try:
                results[idx] = future.result()
            except Exception:
                results[idx] = ""

    return results


def _clean_lines(text: str) -> List[str]:
    lines = []
    for raw in (text or "").splitlines():
        line = re.sub(r"\s{2,}", " ", raw.replace("\t", " ")).strip()
        if line:
            lines.append(line)
    return lines


def _extract_value_after_label(lines: List[str], label_patterns: List[str], lookahead: int = 3) -> str:
    for idx, line in enumerate(lines):
        for pat in label_patterns:
            m = re.search(pat, line, flags=re.IGNORECASE)
            if not m:
                continue
            tail = line[m.end():].strip(" :;,-—№")
            if tail and len(tail) >= 4 and len(re.findall(r"[A-Za-zА-Яа-я]", tail)) >= 2:
                return tail
            for j in range(idx + 1, min(len(lines), idx + 1 + lookahead)):
                candidate = lines[j].strip(" :;,-—")
                if not candidate:
                    continue
                if any(re.search(lp, candidate, flags=re.IGNORECASE) for lp in label_patterns):
                    continue
                if len(candidate) < 4:
                    continue
                if len(re.findall(r"[A-Za-zА-Яа-я]", candidate)) < 2:
                    continue
                return candidate
    return ""


def _pick_document_code(text: str) -> str:
    suffix_map = {
        "PS": "ПС",
        "ПС": "ПС",
        "RE": "РЭ",
        "РЭ": "РЭ",
        "TU": "ТУ",
        "ТУ": "ТУ",
    }
    separated_suffix = re.search(
        r"\b([А-ЯA-Z]{2,6}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?)\s*(ПС|PS|РЭ|RE|ТУ|TU)\b",
        text,
        flags=re.IGNORECASE,
    )
    if separated_suffix:
        base = separated_suffix.group(1).strip()
        raw_suffix = separated_suffix.group(2).strip().upper().replace("Ё", "Е")
        suffix = suffix_map.get(raw_suffix, raw_suffix)
        return f"{base} {suffix}"

    patterns = [
        r"\b[А-ЯA-Z]{2,6}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?[А-ЯA-Z]{0,3}\b",
        r"\b[А-ЯA-Z]{2,6}\.\d{6}\.\d{3,4}\b",
        r"\b\d{4}\.\d{6}\.\d{4}\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return m.group(0).strip()
    return ""

def _extract_normative_docs(text: str) -> List[str]:
    found = []
    patterns = [
        r"\bГОСТ\s*[A-ZА-Я0-9.\-–/ ]{3,40}",
        r"\bТР\s*ТС\s*\d{3}/\d{4}\b",
        r"\b[А-ЯA-Z]{2,6}\.\d{3,6}\.\d{2,4}\s*ТУ\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            value = re.sub(r"\s{2,}", " ", m.group(0)).strip(" ,.;")
            if len(value) < 5:
                continue
            if value not in found:
                found.append(value)
    return found[:10]


def _is_probable_serial(token: str) -> bool:
    t = token.strip(" .,;:()[]{}").upper()
    if not t or not re.search(r"\d", t):
        return False
    if re.match(r"^(ТУ|TY|ТY)[A-ZА-Я0-9\-]{4,}$", t):
        return False
    if re.fullmatch(r"[A-ZА-Я]{2,8}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?[A-ZА-Я]{0,3}", t):
        return False
    if len(t) < 4 or len(t) > 20:
        return False
    if re.fullmatch(r"\d{1,5}", t):
        return False
    if re.fullmatch(r"\d{6,7}", t):
        return False
    if re.fullmatch(r"(19|20)\d{2}", t):
        return False
    if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{2,4}", t):
        return False
    if re.fullmatch(r"\d{8,15}", t):
        return True
    if re.fullmatch(r"[A-ZА-Я]{2,4}\d{3,4}", t):
        return False
    if re.search(r"\d+[XХ]\d+[XХ]\d+", t):
        return False
    if not re.fullmatch(r"[A-ZА-Я0-9\-]{5,20}", t):
        return False
    digit_count = len(re.findall(r"\d", t))
    letter_count = len(re.findall(r"[A-ZА-Я]", t))
    if len(t) < 7 or digit_count < 3 or letter_count < 1:
        return False
    return bool(re.search(r"\d{4,}", t))


def _normalize_serial_token(token: str) -> str:
    t = re.sub(r"\s+", "", str(token or "").upper()).strip(" .,;:()[]{}")
    return t.translate(CYR_TO_LAT_LOOKALIKE)


def _serial_plausibility_score(token: str) -> int:
    t = _normalize_serial_token(token)
    score = 0
    if re.match(r"^[A-Z][0-9]", t):
        score += 4
    if re.match(r"^[A-Z][0-9][A-Z][0-9]{3,}$", t):
        score += 5
    if re.search(r"\d{4,}", t):
        score += 3
    if re.search(r"[A-Z]", t):
        score += 2
    if t and t[0].isdigit() and re.search(r"[A-Z]", t):
        score -= 3
    if t.count("-") > 1:
        score -= 1
    return score


def _serial_candidates(token: str) -> List[str]:
    base = _normalize_serial_token(token)
    if not re.search(r"[A-Z]", base):
        return [base]
    candidates = {base}
    for pos, ch in enumerate(base[:4]):
        alt = SERIAL_CONFUSABLES.get(ch)
        if not alt:
            continue
        candidates.add(base[:pos] + alt + base[pos + 1 :])
    return [c for c in candidates if c]


def _canonicalize_serial(token: str) -> str:
    candidates = []
    for cand in _serial_candidates(token):
        if _is_probable_serial(cand):
            candidates.append(cand)
    if not candidates:
        return _normalize_serial_token(token)
    return max(candidates, key=_serial_plausibility_score)


def _extract_serials(text: str) -> List[str]:
    serials: List[str] = []
    uppercase_text = re.sub(r"\s+", " ", text.upper())

    labeled_pattern = r"ЗАВОД\w*\s*НОМЕР\s*[:;№\-–—]*\s*([A-ZА-Я0-9\-]{4,20})"
    labeled_matches = list(re.finditer(labeled_pattern, uppercase_text, flags=re.IGNORECASE))
    for m in labeled_matches:
        token = _canonicalize_serial(m.group(1).strip())
        if _is_probable_serial(token):
            serials.append(token)

    table_mode = (
        len(labeled_matches) > 1
        or bool(re.search(r"№\s*П/?П|ПЕРЕЧЕНЬ\s+ДОКУМЕНТАЦИИ|СТРАНИЦ\s*/\s*ЛИСТОВ", uppercase_text, flags=re.IGNORECASE))
    )
    table_heading = re.search(r"ЗАВОД\w*\s*НОМЕР", uppercase_text, flags=re.IGNORECASE) if table_mode else None
    if table_heading is not None:
        tail = uppercase_text[table_heading.end() : table_heading.end() + 5000]
        for token in re.findall(r"\b[A-ZА-Я0-9\-]{4,20}\b", tail):
            corrected = _canonicalize_serial(token)
            if _is_probable_serial(corrected):
                serials.append(corrected)

    # Last-chance fallback: broad scan through whole text.
    for token in re.findall(r"\b[A-ZА-Я0-9\-]{5,20}\b", uppercase_text):
        corrected = _canonicalize_serial(token)
        if _is_probable_serial(corrected):
            serials.append(corrected)

    return normalize_serials(serials)


def _is_date_like(value: str) -> bool:
    if not value:
        return False
    if re.search(r"\b[0-3]?\d[./][01]?\d[./](?:19|20)\d{2}\b", value):
        return True
    if re.search(r"\b[0-3]?\d[./][01]?\d[./]\d{2}\b", value):
        return True
    if re.search(r"\b[0-3]?\d\s+[А-Яа-я]{3,}\s+(?:19|20)\d{2}\b", value, flags=re.IGNORECASE):
        return True
    return False


def _is_duration_like(value: str) -> bool:
    if not value:
        return False
    s = re.sub(r"\s+", " ", str(value).lower().replace("ё", "е")).strip(" .,:;")
    if len(s) < 2 or len(s) > 60:
        return False
    if _is_date_like(s):
        return True
    units = r"(год(?:а|ов)?|лет|месяц(?:а|ев)?|сут(?:ки|ок)?|дн(?:ей|я)?|день|недел(?:я|и|ь)|час(?:ов|а)?)"
    num_words = r"(один|одна|два|две|три|четыре|пять|шесть|семь|восемь|девять|десять|двенадцать|пятнадцать|двадцать|тридцать|сорок|пятьдесят)"
    has_amount = bool(re.search(r"\d", s) or re.search(num_words, s))
    return bool(has_amount and re.search(units, s))


def _normalize_numeric_date(day: str, month: str, year: str) -> str:
    d = max(1, min(31, int(day)))
    m = max(1, min(12, int(month)))
    y = int(year)
    if y < 100:
        y = 2000 + y if y <= 39 else 1900 + y
    return f"{d:02d}.{m:02d}.{y:04d}"


def _try_parse_textual_date(value: str) -> str:
    s = (value or "").upper().replace("Ё", "Е")
    s = s.translate({
        ord("A"): "А",
        ord("B"): "В",
        ord("C"): "С",
        ord("E"): "Е",
        ord("H"): "Н",
        ord("K"): "К",
        ord("M"): "М",
        ord("O"): "О",
        ord("P"): "Р",
        ord("T"): "Т",
        ord("X"): "Х",
        ord("Y"): "У",
    })
    m = re.search(r"\b([0-3]?\d)\s*[-./ ]?\s*([А-Я]{3,8})\.?\s*((?:19|20)?\d{2,4})\b", s)
    if not m:
        return ""
    day, month_word, year = m.group(1), m.group(2), m.group(3)
    month_key = next((k for k in RU_MONTHS if month_word.startswith(k)), "")
    if not month_key:
        return ""
    if len(year) == 4 and not year.startswith(("19", "20")):
        return ""
    if len(year) not in (2, 4):
        return ""
    return _normalize_numeric_date(day, RU_MONTHS[month_key], year)


def _extract_dates_from_chunk(chunk: str) -> List[str]:
    found = []
    for m in re.finditer(r"\b([0-3]?\d)[./]([01]?\d)[./]((?:19|20)?\d{2,4})\b", chunk):
        day, month, year = m.group(1), m.group(2), m.group(3)
        if len(year) == 4 and not year.startswith(("19", "20")):
            continue
        if len(year) not in (2, 4):
            continue
        try:
            found.append(_normalize_numeric_date(day, month, year))
        except Exception:
            continue

    textual = _try_parse_textual_date(chunk)
    if textual:
        found.append(textual)
    return list(dict.fromkeys(found))


def _extract_version_dates(text: str) -> set:
    result = set()
    for m in re.finditer(r"версия[^\n]{0,50}", text, flags=re.IGNORECASE):
        for d in _extract_dates_from_chunk(m.group(0)):
            result.add(d)
    return result


def _find_date_near_label(lines: List[str], label_patterns: List[str], window: int = 6) -> str:
    for idx, line in enumerate(lines):
        if not any(re.search(pat, line, flags=re.IGNORECASE) for pat in label_patterns):
            continue
        start = max(0, idx - 1)
        end = min(len(lines), idx + 1 + window)
        chunk = " ".join(lines[start:end])
        dates = _extract_dates_from_chunk(chunk)
        if dates:
            return dates[0]
    return ""


def _extract_dates(text: str, lines: List[str]) -> Tuple[str, str]:
    date_release = _find_date_near_label(lines, [r"дата\s+выпуск"])
    date_accept = _find_date_near_label(lines, [r"дата\s+при[её]м", r"свидетельство\s+о\s+при[её]мк"])

    if not date_release:
        for m in re.finditer(r"дата\s+выпуск[^\n]{0,90}", text, flags=re.IGNORECASE):
            dates = _extract_dates_from_chunk(m.group(0))
            if dates:
                date_release = dates[0]
                break

    if not date_accept:
        for m in re.finditer(r"дата\s+при[её]м[^\n]{0,90}", text, flags=re.IGNORECASE):
            dates = _extract_dates_from_chunk(m.group(0))
            if dates:
                date_accept = dates[0]
                break

    version_dates = _extract_version_dates(text)
    if date_release in version_dates:
        date_release = ""
    if date_accept in version_dates:
        date_accept = ""

    return date_release, date_accept


def _extract_certificate_value(text: str, lines: List[str]) -> str:
    patterns = [
        r"\bЕАЭС\s*[A-ZА-Я0-9.\-\/]{6,}",
        r"\bСДС\.[A-ZА-Я0-9.\-\/]{6,}",
        r"\b(?:RU|РФ)\s*[A-ZА-Я0-9.\-\/]{5,}",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return re.sub(r"\s{2,}", " ", m.group(0)).strip(" .,:;")

    for idx, line in enumerate(lines):
        if not re.search(r"\bсертификат\w*\b", line, flags=re.IGNORECASE):
            continue
        chunk = " ".join(lines[idx: min(len(lines), idx + 4)])
        num = re.search(r"(?:№|N)\s*([A-ZА-Я0-9.\-\/]{5,})", chunk, flags=re.IGNORECASE)
        if num:
            return num.group(1).strip(" .,:;")

    by_label = _extract_value_after_label(lines, [r"\bсертификат\w*\b"])
    if not by_label:
        return ""
    s = by_label.strip(" .,:;")
    if len(s) > 80 or len(s.split()) > 5:
        return ""
    if not re.search(r"\d", s):
        return ""
    s_norm = s.lower().replace("ё", "е")
    if re.search(r"(соответств|лиценз|налич|услов|контролл|модул|шкаф|блок|перечн|документац)", s_norm):
        return ""
    if re.fullmatch(r"(соответстви[ея]|наличи[ея]|сертификат[а-я ]*)", s_norm, flags=re.IGNORECASE):
        return ""
    if not (
        re.search(r"[A-ZА-Я0-9]{2,}[.\-/][A-ZА-Я0-9.\-/]{2,}", s, flags=re.IGNORECASE)
        or re.search(r"(?:№|N)\s*[A-ZА-Я0-9.\-/]{5,}", s, flags=re.IGNORECASE)
    ):
        return ""
    return s


def _build_processing_hint(raw_text: str, doc_type: str, file_name: str = "") -> str:
    text = (raw_text or "").lower().replace("ё", "е")
    fname = (file_name or "").lower().replace("ё", "е")
    merged = f"{text}\n{fname}"
    if doc_type in {"single_passport", "group_passport"}:
        return "Документ похож на паспорт/РЭ. Проверьте поля справа и сохраните запись."
    if doc_type == "cabinet_list":
        return "Документ похож на перечень шкафа. Лучше загружать его в режим «Шкаф» для сверки позиций."
    if "служебная записка" in merged:
        return "Это служебная записка, а не паспорт изделия. Поля паспорта автоматически не заполняются."
    if "экранная форма" in merged or "исходной таблицы" in merged:
        return "Это шаблон/экранная форма задания. Для извлечения паспортных данных загрузите именно паспорт или РЭ."
    if "разбор паспорта" in merged:
        return "Похоже на разбор/перечень для сверки. Откройте режим «Шкаф» и загрузите файл туда."
    return "Документ не распознан как паспорт. Попробуйте загрузить паспорт/РЭ или перечень документации шкафа."


def _normalize_ocr_ru_token(value: str) -> str:
    s = (value or "").upper().replace("Ё", "Е")
    s = s.translate({
        ord("A"): "А",
        ord("B"): "В",
        ord("C"): "С",
        ord("E"): "Е",
        ord("H"): "Н",
        ord("K"): "К",
        ord("M"): "М",
        ord("O"): "О",
        ord("P"): "Р",
        ord("T"): "Т",
        ord("X"): "Х",
        ord("Y"): "У",
    })
    return re.sub(r"[^А-Я0-9]", "", s)


def _extract_blue_stamp_layer(img_rgb: Image.Image) -> Image.Image:
    r, g, b = img_rgb.split()
    rg_avg = ImageChops.blend(r, g, 0.5)
    blue = ImageChops.subtract(b, rg_avg)
    return ImageOps.autocontrast(blue)


def _date_from_roi(img: Image.Image, banned_dates: set) -> str:
    prepared = [
        img,
        ImageEnhance.Contrast(img).enhance(1.8),
        ImageEnhance.Sharpness(ImageEnhance.Contrast(img).enhance(1.9)).enhance(1.6),
        img.point(lambda px: 255 if px > 145 else 0, mode="1").convert("L"),
        img.point(lambda px: 255 if px > 165 else 0, mode="1").convert("L"),
    ]
    for variant in prepared:
        for psm in ("6", "11", "7"):
            try:
                txt = pytesseract.image_to_string(
                    variant,
                    lang=OCR_LANG,
                    config=f"--oem 1 --psm {psm}",
                )
            except Exception:
                continue
            for d in _extract_dates_from_chunk(txt):
                if d not in banned_dates:
                    return d
    return ""


def _extract_release_date_from_page_image(page_img: bytes, banned_dates: Optional[set] = None) -> str:
    banned = set(banned_dates or set())
    try:
        img_rgb = Image.open(io.BytesIO(page_img)).convert("RGB")
    except Exception:
        return ""

    img = ImageOps.autocontrast(img_rgb.convert("L"))
    blue_layer = _extract_blue_stamp_layer(img_rgb)
    if img.width < 2200:
        scale = 2200 / float(max(1, img.width))
        img = img.resize((int(img.width * scale), int(img.height * scale)), Image.Resampling.LANCZOS)
        blue_layer = blue_layer.resize((int(blue_layer.width * scale), int(blue_layer.height * scale)), Image.Resampling.LANCZOS)
    img = ImageEnhance.Sharpness(img).enhance(1.7)
    blue_layer = ImageEnhance.Sharpness(blue_layer).enhance(1.5)

    label_rois: List[Tuple[int, int, int, int]] = []
    try:
        data = pytesseract.image_to_data(
            img,
            lang=OCR_LANG,
            config="--oem 1 --psm 6",
            output_type=pytesseract.Output.DICT,
        )
        words = data.get("text", [])
        left = data.get("left", [])
        top = data.get("top", [])
        heights = data.get("height", [])
        widths = data.get("width", [])
        normalized = [_normalize_ocr_ru_token(w) for w in words]

        for i, token in enumerate(normalized):
            if not token or not token.startswith("ДАТА"):
                continue
            nearby = normalized[i:i + 6]
            if not any(t.startswith("ВЫПУСК") for t in nearby):
                continue
            x = left[i]
            y = top[i]
            h = heights[i] if i < len(heights) else 20
            w = widths[i] if i < len(widths) else 60
            x1 = max(0, x - int(w * 0.8))
            y1 = max(0, y - int(h * 1.5))
            x2 = min(img.width, x + int(img.width * 0.55))
            y2 = min(img.height, y + int(img.height * 0.30))
            if x2 - x1 > 120 and y2 - y1 > 50:
                label_rois.append((x1, y1, x2, y2))
    except Exception:
        label_rois = []

    for box in label_rois:
        for layer in (img, blue_layer):
            found = _date_from_roi(layer.crop(box), banned)
            if found:
                return found

    fallback_boxes = [
        (0, int(img.height * 0.30), int(img.width * 0.72), int(img.height * 0.78)),
        (0, int(img.height * 0.40), int(img.width * 0.80), int(img.height * 0.86)),
        (0, int(img.height * 0.48), int(img.width * 0.70), int(img.height * 0.96)),
        (int(img.width * 0.02), int(img.height * 0.52), int(img.width * 0.62), int(img.height * 0.90)),
    ]
    for box in fallback_boxes:
        for layer in (img, blue_layer):
            found = _date_from_roi(layer.crop(box), banned)
            if found:
                return found
    return ""


def _extract_release_date_from_pages(pages: List[bytes], banned_dates: Optional[set] = None) -> str:
    banned = set(banned_dates or set())
    for page_img in pages[:2]:
        found = _extract_release_date_from_page_image(page_img, banned_dates=banned)
        if found:
            return found
    return ""


def _extract_namenovanie(lines: List[str], full_text: str) -> str:
    for line in lines[:80]:
        if re.search(r"\bрозетк\w*\b", line, flags=re.IGNORECASE):
            if re.search(r"[A-Za-zА-Яа-я0-9]{4,}", line) and len(line) <= 140:
                return line.strip(" .,:;")

    for line in lines[:15]:
        if re.search(r"\bшкаф\b", line, flags=re.IGNORECASE):
            return line.strip()

    for idx, line in enumerate(lines[:120]):
        if re.search(r"\bконтроллер\b", line, flags=re.IGNORECASE):
            if re.search(r"мфк\s*[- ]?\d+", line, flags=re.IGNORECASE):
                return line.strip(" .,:;")
            for j in range(idx + 1, min(len(lines), idx + 6)):
                if re.search(r"\b(мфк\s*[- ]?\d+|tcc\s*[a-z0-9-]*|tecon|текон)\b", lines[j], flags=re.IGNORECASE):
                    return f"{line}. {lines[j]}".strip(". ")

    for line in lines[:120]:
        if re.search(r"\b(мастер[- ]?модул|контроллер\s+многофункциональ)\b", line, flags=re.IGNORECASE):
            return line.strip(" .,:;")

    by_label = _extract_value_after_label(lines, [r"наименован\w*\s+издел", r"изделие"])
    if by_label and not re.search(r"\b(в соответствии|услови|предприят)\b", by_label, flags=re.IGNORECASE):
        if not re.fullmatch(r"[A-ZА-Я]{2,8}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?[A-ZА-Я]{0,3}", by_label):
            return by_label

    # Prefer explicit "продукт + модель" style lines in the first page chunk.
    for line in lines[:120]:
        if len(line) > 140:
            continue
        if re.search(r"\b(розетк|модул|контроллер|шкаф|блок)\b", line, flags=re.IGNORECASE) and re.search(r"[A-Za-zА-Яа-я]+\d", line):
            return line.strip(" .,:;")

    for line in lines[:100]:
        if len(line) > 150:
            continue
        if re.search(r"\b(соответств|услови|предприят|контроль|эксплуатац)\b", line, flags=re.IGNORECASE):
            continue
        if re.search(r"\b(модул|контроллер|реле|панел|светильник|блок|шкаф)\b", line, flags=re.IGNORECASE):
            if re.search(r"\b(таблиц|гост|параметр|версия)\b", line, flags=re.IGNORECASE):
                continue
            return line.strip(" .,:;")

    m = re.search(r"(?:паспорт|руководство по эксплуатации)\s+([^\n]{4,140})", full_text, flags=re.IGNORECASE)
    if m:
        candidate = m.group(1).strip(" .,:;")
        if not re.fullmatch(r"[A-ZА-Я]{2,8}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?[A-ZА-Я]{0,3}", candidate):
            return candidate
    return ""


def _extract_structured_from_text(raw_text: str) -> Dict:
    lines = _clean_lines(raw_text)
    text = "\n".join(lines)
    text_lower = text.lower().replace("ё", "е")

    serials = _extract_serials(text)
    non_passport_markers = [
        "служебная записка",
        "фактическое поступление",
        "рабочее место оператора",
        "критерии оценивания",
        "контекст задачи",
        "цель задачи",
        "хакатон",
        "экранная форма",
        "исходной таблицы",
    ]
    is_non_passport = any(marker in text_lower for marker in non_passport_markers)

    doc_type = "unknown"
    if "перечень документации" in text_lower or ("шкаф" in text_lower and "зав." in text_lower):
        doc_type = "cabinet_list"
    elif is_non_passport:
        doc_type = "unknown"
    elif "паспорт" in text_lower or "руководство по эксплуатации" in text_lower:
        if "tcc8l" in text_lower or "мфк1500" in text_lower or len(serials) > 1:
            doc_type = "group_passport"
        else:
            doc_type = "single_passport"

    name = _extract_namenovanie(lines, text)
    doc_code = _pick_document_code(text)
    kod_zakaza = _extract_value_after_label(lines, [r"код\s+заказ"])
    if kod_zakaza:
        kz_norm = _normalize_serial_token(kod_zakaza)
        serials = [s for s in serials if _normalize_serial_token(s) != kz_norm]
        if doc_type == "group_passport" and len(serials) <= 1 and "tcc8l" not in text_lower and "мфк1500" not in text_lower:
            doc_type = "single_passport"
    garantia = _extract_value_after_label(lines, [r"\bгарант\w*\b"])
    srok = _extract_value_after_label(lines, [r"\bсрок\s+служб\w*\b"])
    sert = _extract_certificate_value(text, lines)
    date_release, date_accept = _extract_dates(text, lines)

    manufacturer = ""
    manufacturer_match = re.search(r"\b(?:АО|ПАО|ООО|ЗАО)\s*[\"«][^\"»\n]{2,80}[\"»]", text, flags=re.IGNORECASE)
    if manufacturer_match:
        manufacturer = manufacturer_match.group(0).strip()
    else:
        manufacturer_match = re.search(r"\b(?:АО|ПАО|ООО|ЗАО)\s+[А-ЯA-Z][^\n,]{2,80}", text, flags=re.IGNORECASE)
        if manufacturer_match:
            manufacturer = manufacturer_match.group(0).strip()

    address = ""
    for line in lines:
        if len(line) < 10 or len(line) > 180:
            continue
        if "," not in line:
            continue
        if re.search(r"(россия|рф|г\.\s*[а-я]|ул\.|улиц|просп|пр-т|д\.|дом|корп|стр\.)", line, flags=re.IGNORECASE) and re.search(r"\d", line):
            if len(re.findall(r"\w+", line)) > 24:
                continue
            address = line
            break

    contacts = []
    contacts.extend(re.findall(r"(?:\+7|8)\s*\(?\d{3,4}\)?[\s\-]?\d{2,3}[\s\-]?\d{2}[\s\-]?\d{2}", text))
    contacts.extend(re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, flags=re.IGNORECASE))

    extracted = {
        "document_type": doc_type,
        "naimenovanie": name,
        "kod_dokumenta": doc_code,
        "proizvoditel": manufacturer,
        "adres": address,
        "kontakty": ", ".join(dict.fromkeys(contacts)) if contacts else "",
        "garantia": garantia,
        "srok_sluzhby": srok,
        "sertifikat": sert,
        "kod_zakaza": kod_zakaza,
        "normativnye_dok": _extract_normative_docs(text),
        "zavodskie_nomera": serials,
        "data_vypuska": date_release,
        "data_priemki": date_accept,
        "_non_passport": is_non_passport,
    }
    if extracted["garantia"] and not _is_duration_like(extracted["garantia"]):
        extracted["garantia"] = ""
    if extracted["srok_sluzhby"] and not _is_duration_like(extracted["srok_sluzhby"]):
        extracted["srok_sluzhby"] = ""
    if extracted["sertifikat"] and (len(extracted["sertifikat"]) > 80 or not re.search(r"[A-ZА-Я0-9]", extracted["sertifikat"], flags=re.IGNORECASE)):
        extracted["sertifikat"] = ""
    if extracted["kod_zakaza"] and len(extracted["kod_zakaza"]) > 40:
        extracted["kod_zakaza"] = ""

    return extracted


def parse_cabinet_document(raw_text: str) -> Dict:
    lines = _clean_lines(raw_text)
    text = "\n".join(lines)
    result = {
        "shkaf_naim": "",
        "shkaf_kod": "",
        "shkaf_zav_nomer": "",
        "pozicii": [],
    }

    for line in lines[:20]:
        if re.search(r"\bшкаф\b", line, flags=re.IGNORECASE):
            result["shkaf_naim"] = line
            code_match = re.search(r"\(([^)]+)\)", line)
            if code_match:
                result["shkaf_kod"] = code_match.group(1).strip()
            break

    serial_header_idx = next((i for i, l in enumerate(lines) if re.search(r"заводск\w*\s+номер", l, flags=re.IGNORECASE)), -1)
    serial_source = "\n".join(lines[serial_header_idx:]) if serial_header_idx >= 0 else text

    zav_match = re.search(r"зав\.\s*№\s*([A-ZА-Я0-9\-]+)", serial_source, flags=re.IGNORECASE)
    if zav_match:
        result["shkaf_zav_nomer"] = zav_match.group(1).strip()

    doc_code_pattern = re.compile(r"\b[А-ЯA-Z]{2,6}\.\d{3,6}\.\d{2,4}(?:-\d{2,3})?[А-ЯA-Z]{0,3}\b")
    serial_pattern = re.compile(r"\b(?:\d{9,14}|[A-ZА-Я0-9]{5,20}|б/н|Б/Н)\b")

    items = []
    for i, line in enumerate(lines):
        for m in serial_pattern.finditer(line):
            serial = m.group(0).strip()
            if serial.lower() != "б/н" and not _is_probable_serial(serial):
                continue

            name_parts = []
            for j in range(i - 1, max(-1, i - 4), -1):
                candidate = lines[j]
                if re.search(r"(перечень|заводской|страниц|листов|сертификат|паспорт\b)", candidate, flags=re.IGNORECASE):
                    continue
                if re.fullmatch(r"\d+[.)]?", candidate):
                    continue
                if len(candidate) < 4:
                    continue
                name_parts.insert(0, candidate)
                if len(" ".join(name_parts)) >= 90:
                    break

            nearby = " ".join(lines[i:i + 4])
            doc_match = doc_code_pattern.search(nearby)
            doc_code = doc_match.group(0) if doc_match else ""

            items.append({
                "naimenovanie": re.sub(r"\s{2,}", " ", " ".join(name_parts)).strip() or f"Позиция {len(items) + 1}",
                "zavodskoy_nomer": serial,
                "oboznachenie_dok": doc_code,
            })

    seen = set()
    for item in items:
        key = (item["zavodskoy_nomer"], item["oboznachenie_dok"], item["naimenovanie"])
        if key in seen:
            continue
        seen.add(key)
        result["pozicii"].append({
            "nomer": len(result["pozicii"]) + 1,
            **item,
        })

    return result


def final_cleanup(data: Dict, raw_text: str = "") -> Dict:
    data = data or {}
    heur = _extract_structured_from_text(raw_text)

    if heur.get("_non_passport"):
        data = {}

    merge_keys = [
        "document_type",
        "naimenovanie",
        "kod_dokumenta",
        "proizvoditel",
        "adres",
        "kontakty",
        "garantia",
        "srok_sluzhby",
        "sertifikat",
        "kod_zakaza",
        "normativnye_dok",
        "zavodskie_nomera",
        "data_vypuska",
        "data_priemki",
    ]
    for key in merge_keys:
        value = heur.get(key)
        if value:
            data[key] = value

    if data.get("naimenovanie"):
        name = re.sub(
            r"\b(паспорт|руководство|эксплуатации|свидетельство|приемке|упаковывании)\b",
            "",
            str(data["naimenovanie"]),
            flags=re.IGNORECASE,
        )
        data["naimenovanie"] = re.sub(r"\s{2,}", " ", name).strip(" ,.-")

    data["zavodskie_nomera"] = normalize_serials(data.get("zavodskie_nomera"))
    data["normativnye_dok"] = [x for x in (data.get("normativnye_dok") or []) if str(x).strip()]
    data["komplektnost"] = [x for x in (data.get("komplektnost") or []) if str(x).strip()]

    fb = regex_fallbacks(raw_text)
    for k, v in fb.items():
        if k == "zavodskie_nomera":
            merged = normalize_serials((data.get(k) or []) + v)
            data[k] = merged
        elif not data.get(k):
            data[k] = v

    data["zavodskie_nomera"] = [s for s in data.get("zavodskie_nomera", []) if not re.fullmatch(r"\d{1,5}", s)]
    if data.get("kod_zakaza"):
        kz_norm = _normalize_serial_token(str(data.get("kod_zakaza", "")))
        data["zavodskie_nomera"] = [
            s for s in data.get("zavodskie_nomera", [])
            if _normalize_serial_token(s) != kz_norm
        ]
    doc_type = (data.get("document_type") or "").strip().lower()
    if doc_type == "unknown":
        for key in (
            "naimenovanie",
            "kod_dokumenta",
            "kod_zakaza",
            "proizvoditel",
            "adres",
            "kontakty",
            "garantia",
            "srok_sluzhby",
            "sertifikat",
            "data_vypuska",
            "data_priemki",
        ):
            data[key] = None
        data["zavodskie_nomera"] = []
        data["normativnye_dok"] = []
        data["komplektnost"] = []
    elif doc_type == "cabinet_list":
        for key in ("garantia", "srok_sluzhby", "sertifikat", "data_vypuska", "data_priemki"):
            data[key] = None

    if data.get("naimenovanie"):
        name = str(data.get("naimenovanie", "")).strip()
        if re.fullmatch(r"контроллер\s+многофункциональн(ый|ого)?", name, flags=re.IGNORECASE):
            data["naimenovanie"] = None

    if data.get("adres"):
        addr = str(data["adres"]).strip()
        if len(addr) > 140 or len(re.findall(r"\w+", addr)) > 24:
            data["adres"] = None

    if doc_type in {"group_passport", "cabinet_list"}:
        data["tip_pasporta"] = "group"
    elif data.get("zavodskie_nomera"):
        data["tip_pasporta"] = "individual"
    else:
        data["tip_pasporta"] = "no_serial"

    return data


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/extract", methods=["POST"])
def extract():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file"}), 400

    try:
        fbytes = file.read()

        if file.filename.lower().endswith(".pdf"):
            pages, embedded_text_blocks = pdf_to_images_and_text(fbytes)
        else:
            pages, embedded_text_blocks = [fbytes], [""]

        provider, active_model = resolve_llm_backend()
        run_ocr = provider != "openai" or USE_OCR_WITH_OPENAI
        ocr_texts = collect_ocr_texts(pages, embedded_text_blocks, enabled=run_ocr)
        embedded_text = "\n".join(embedded_text_blocks)
        full_text = "\n".join(part for part in (embedded_text, "\n".join(ocr_texts)) if part).strip()

        llm_images = prepare_llm_images_b64(pages)
        llm_data = {}
        llm_error = ""
        if ENABLE_LLM:
            try:
                llm_data = run_llm_extraction(llm_images, full_text, provider=provider, model_name=active_model)
            except requests.RequestException as e:
                llm_error = str(e)
                log.warning("LLM extraction failed, fallback to OCR+rules only: %s", llm_error)
        else:
            llm_error = "LLM disabled (ENABLE_LLM=0)"

        final_result = final_cleanup(llm_data, raw_text=full_text)
        if final_result.get("document_type") != "unknown" and not final_result.get("data_vypuska"):
            version_dates = _extract_version_dates(full_text)
            release_date = _extract_release_date_from_pages(pages, banned_dates=version_dates)
            if release_date:
                final_result["data_vypuska"] = release_date
        final_result["_hint"] = _build_processing_hint(
            full_text,
            str(final_result.get("document_type") or "unknown"),
            file.filename or "",
        )
        final_result["_meta"] = {
            "provider": provider,
            "model": active_model,
            "pages_processed": len(pages),
            "images_sent_to_llm": len(llm_images),
            "ocr_enabled": run_ocr,
            "llm_error": llm_error,
        }
        return jsonify(final_result)

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/parse_cabinet", methods=["POST"])
def parse_cabinet():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file"}), 400

    try:
        fbytes = file.read()
        if file.filename.lower().endswith(".pdf"):
            pages, embedded_text_blocks = pdf_to_images_and_text(fbytes)
        else:
            pages, embedded_text_blocks = [fbytes], [""]

        ocr_texts = collect_ocr_texts(pages, embedded_text_blocks, enabled=True)
        full_text = "\n".join(part for part in ("\n".join(embedded_text_blocks), "\n".join(ocr_texts)) if part).strip()
        result = parse_cabinet_document(full_text)
        return jsonify(result)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/barcode", methods=["POST"])
def api_barcode():
    val = (request.json or {}).get("value", "").strip()
    if not val:
        return jsonify({"error": "Empty"}), 400
    bc_class = barcode.get_barcode_class("code128")
    buf = io.BytesIO()
    clean_val = "".join(re.findall(r"[A-Z0-9\-]", val.upper()))
    bc_class(clean_val[:20] or "0000", writer=ImageWriter()).write(buf)
    return jsonify({"barcode": base64.b64encode(buf.getvalue()).decode(), "encoded_value": clean_val})


@app.route("/api/preview", methods=["POST"])
def preview():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file"}), 400
    fbytes = file.read()
    images = []
    if file.filename.lower().endswith(".pdf"):
        doc = fitz.open(stream=fbytes, filetype="pdf")
        for page in doc:
            img = page.get_pixmap(matrix=fitz.Matrix(PREVIEW_RENDER_SCALE, PREVIEW_RENDER_SCALE)).tobytes("png")
            images.append(base64.b64encode(img).decode())
    else:
        images.append(base64.b64encode(fbytes).decode())
    return jsonify({
        "images": images,
        "image": images[0] if images else "",
        "page_count": len(images),
    })


@app.route("/api/meta", methods=["GET"])
def meta():
    provider, active_model = resolve_llm_backend()
    return jsonify({
        "active_provider": provider,
        "active_model": active_model,
        "model_candidates": MODEL_CANDIDATES if provider == "ollama" else [OPENAI_MODEL],
        "max_llm_images": MAX_LLM_IMAGES,
        "ocr_workers": OCR_WORKERS,
        "enable_llm": ENABLE_LLM,
        "llm_timeout_sec": LLM_REQUEST_TIMEOUT,
    })


@app.route("/api/export/excel", methods=["POST"])
def export_excel():
    payload = request.json or {}
    records = payload.get("records", [])
    cabinet = payload.get("cabinet") or {}

    def _safe_join(values: Any) -> str:
        if not values:
            return ""
        if isinstance(values, list):
            return ", ".join(str(v).strip() for v in values if str(v).strip())
        return str(values).strip()

    def _autosize(ws, max_width: int = 68):
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = "" if cell.value is None else str(cell.value)
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), max_width)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Паспорта"
    ws.append([
        "Файл",
        "Тип документа",
        "Тип паспорта",
        "Наименование",
        "Код документа",
        "Код заказа",
        "Дата выпуска",
        "Дата приемки ОТК",
        "Серийные номера",
        "Количество серийных",
        "Производитель",
        "Адрес",
        "Контакты",
        "Гарантия",
        "Срок службы",
        "Сертификат",
        "Нормативные документы",
        "Сохранено",
    ])
    for r in records:
        serials = normalize_serials(r.get("zavodskie_nomera"))
        ws.append([
            r.get("_fileName"),
            r.get("document_type"),
            r.get("tip_pasporta"),
            r.get("naimenovanie"),
            r.get("kod_dokumenta"),
            r.get("kod_zakaza"),
            r.get("data_vypuska"),
            r.get("data_priemki"),
            _safe_join(serials),
            len(serials),
            r.get("proizvoditel"),
            r.get("adres"),
            r.get("kontakty"),
            r.get("garantia"),
            r.get("srok_sluzhby"),
            r.get("sertifikat"),
            _safe_join(r.get("normativnye_dok")),
            "Да" if r.get("_saved") else "Нет",
        ])
    ws.freeze_panes = "A2"
    _autosize(ws)

    ws_sn = wb.create_sheet("Серийные номера")
    ws_sn.append(["Файл", "Наименование", "Код документа", "Серийный номер"])
    for r in records:
        serials = normalize_serials(r.get("zavodskie_nomera"))
        for sn in serials:
            ws_sn.append([r.get("_fileName"), r.get("naimenovanie"), r.get("kod_dokumenta"), sn])
    ws_sn.freeze_panes = "A2"
    _autosize(ws_sn)

    ws_norm = wb.create_sheet("Нормативы")
    ws_norm.append(["Файл", "Наименование", "Код документа", "Нормативный документ"])
    for r in records:
        docs = r.get("normativnye_dok") or []
        for doc_name in docs:
            ws_norm.append([r.get("_fileName"), r.get("naimenovanie"), r.get("kod_dokumenta"), doc_name])
    ws_norm.freeze_panes = "A2"
    _autosize(ws_norm)

    if cabinet:
        ws_cab = wb.create_sheet("Шкаф")
        ws_cab.append([
            "Шкаф",
            "Код шкафа",
            "Заводской номер шкафа",
            "№",
            "Позиция",
            "Заводской номер",
            "Обозначение документа",
        ])
        for item in cabinet.get("pozicii") or []:
            ws_cab.append([
                cabinet.get("shkaf_naim"),
                cabinet.get("shkaf_kod"),
                cabinet.get("shkaf_zav_nomer"),
                item.get("nomer"),
                item.get("naimenovanie"),
                item.get("zavodskoy_nomer"),
                item.get("oboznachenie_dok"),
            ])
        ws_cab.freeze_panes = "A2"
        _autosize(ws_cab)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, download_name="Registry.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)


