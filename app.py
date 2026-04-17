import base64
import io
import re
import traceback

import barcode
import fitz
import openpyxl
import requests
from barcode.writer import ImageWriter
from flask import Flask, jsonify, render_template, request, send_file

from processing import *
from processing import (
    _extract_release_date_from_pages,
    _extract_version_dates,
    _has_strong_heuristic_result,
    _ollama_health_cache,
    _prioritize_ocr_for_llm,
    _should_skip_llm_for_file,
    _should_try_release_date_scan,
)

app = Flask(__name__)

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
        run_ocr = True
        ocr_texts = collect_ocr_texts(pages, embedded_text_blocks, enabled=run_ocr)
        embedded_text = "\n".join(embedded_text_blocks)
        full_text = "\n".join(part for part in (embedded_text, "\n".join(ocr_texts)) if part).strip()
        heur_preview = final_cleanup({}, raw_text=full_text, source_name=file.filename)

        llm_images: List[str] = []
        llm_data = {}
        llm_error = ""
        skip_by_content = _should_skip_llm_for_file(file.filename, full_text)
        skip_by_quality = _has_strong_heuristic_result(heur_preview)
        llm_available = ENABLE_LLM and provider != "disabled"

        if llm_available and not skip_by_content and not skip_by_quality:
            llm_images = prepare_llm_images_b64(pages)
            llm_input_text = full_text
            if int(heur_preview.get("quality_score") or 0) < 25:
                llm_input_text = _prioritize_ocr_for_llm(full_text, max(1200, LLM_FALLBACK_TEXT_LIMIT))
            try:
                llm_data = run_llm_extraction(llm_images, llm_input_text, provider=provider, model_name=active_model)
            except requests.RequestException as e:
                llm_error = str(e)
                log.info("LLM extraction failed, fallback to OCR+rules only: %s", llm_error)
        elif llm_available:
            if skip_by_content:
                llm_error = "LLM skipped: likely non-passport/service input"
            elif skip_by_quality:
                llm_error = "LLM skipped: OCR+rules confidence is already high"
            else:
                llm_error = "LLM skipped"
        else:
            if ENABLE_LLM:
                llm_error = "LLM unavailable: Ollama is not reachable"
            else:
                llm_error = "LLM disabled (ENABLE_LLM=0)"

        final_result = heur_preview if not llm_data else final_cleanup(llm_data, raw_text=full_text, source_name=file.filename)
        if (
            final_result.get("document_type") != "unknown"
            and not final_result.get("data_vypuska")
            and _should_try_release_date_scan(full_text)
        ):
            version_dates = _extract_version_dates(full_text)
            release_date = _extract_release_date_from_pages(pages, banned_dates=version_dates)
            if release_date:
                final_result["data_vypuska"] = release_date
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
        "model_candidates": MODEL_CANDIDATES if provider == "ollama" else [],
        "max_llm_images": MAX_LLM_IMAGES,
        "ocr_workers": OCR_WORKERS,
        "enable_llm": ENABLE_LLM,
        "llm_timeout_sec": LLM_REQUEST_TIMEOUT,
        "llm_connect_timeout_sec": LLM_CONNECT_TIMEOUT,
        "llm_max_total_sec": LLM_MAX_TOTAL_SEC,
        "llm_model_failover_tries": LLM_MODEL_FAILOVER_TRIES,
        "llm_primary_max_images": LLM_PRIMARY_MAX_IMAGES,
        "ollama_reachable": bool(_ollama_health_cache.get("alive")),
        "ollama_last_error": _ollama_health_cache.get("last_error", ""),
        "ocr_psms": OCR_PSMS,
        "ocr_tesseract_timeout_sec": OCR_TESSERACT_TIMEOUT_SEC,
        "ocr_date_tesseract_timeout_sec": OCR_DATE_TESSERACT_TIMEOUT_SEC,
        "ocr_release_scan_budget_sec": OCR_RELEASE_SCAN_BUDGET_SEC,
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
        "Качество (%)",
        "Нужна проверка",
        "Пробелы в данных",
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
            r.get("quality_score"),
            "Да" if r.get("needs_review") else "Нет",
            _safe_join(r.get("missing_fields")),
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


