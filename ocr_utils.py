"""OCR y extracción de datos de comprobantes (pago móvil / facturas VE)."""
from __future__ import annotations

import re
from datetime import datetime

PATTERNS_FECHA = [
    r"FECHA\s*\n\s*(\d{2}/\d{2}/\d{4})",
    r"FECHA[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})",
    r"(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}[^\n]*)",
    r"(\d{2}[/-]\d{2}[/-]\d{4})",
]
PATTERNS_MONTO = [
    r"MONTO\s+DE\s+LA\s+OPERACION\s*[^\n]*\nBs\.?\s*([\d\.]+,\d{2})",
    r"MONTO[:\s]*Bs\.?\s*([\d\.,]+\d)",
    r"TOTAL\s+Bs\.?\s*([\d\.]+,\d{2})",
    r"TOTAL[:\s]*([\d\.]+,\d{2})",
    r"Bs\.?\s*([\d\.]+,\d{2})",
]
PATTERNS_REF = [
    r"NUMERO\s+DE\s+REFERENCIA\s+[^\n]*\n(\d+)",
    r"REFERENCIA[:\s]*(\d{6,})",
    r"REF\.?[:\s]*(\d{6,})",
    r"N[°º]\s*(\d{6,})",
]
PATTERNS_PROV = [
    r"CONCEPTO\s*([^\n]+)",
    r"DESTINO\s*([^\n]+)",
    r"BENEFICIARIO\s*([^\n]+)",
    r"IDENTIFICACION\s+RECEPTOR\s[^\n]+\n([^\n]+)",
]


def _first_match(text: str, patterns: list[str]) -> str | None:
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            val = (m.group(1) or "").strip()
            if val:
                return val
    return None


def _clean_fecha(raw: str | None) -> str:
    if not raw:
        return ""
    raw = raw.strip()[:10].replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return raw


def _clean_monto(raw: str | None) -> float:
    if not raw:
        return 0.0
    s = raw.strip().replace(" ", "")
    # Formato venezolano 1.234,56 -> 1234.56
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(re.sub(r"[^\d.]", "", s) or 0), 2)
    except ValueError:
        return 0.0


def extraer_campos(texto: str) -> dict:
    prov = _first_match(texto, PATTERNS_PROV) or "Por identificar"
    prov = " ".join(prov.split())[:60].capitalize()
    fecha = _clean_fecha(_first_match(texto, PATTERNS_FECHA))
    ref = _first_match(texto, PATTERNS_REF) or ""
    # Para montos con varios matches, preferir el último (suele ser el total)
    montos = []
    for p in PATTERNS_MONTO:
        montos += re.findall(p, texto, re.IGNORECASE)
    monto = _clean_monto(montos[-1] if montos else None)
    return {"proveedor": prov, "fecha": fecha or datetime.now().strftime("%d/%m/%Y"), "factura": ref.strip(), "monto_bs": monto, "descripcion": ""}


def ocr_imagen(path_or_bytes, filename: str = "") -> tuple[str, dict | None, str | None]:
    """Devuelve (texto, campos, error). Si tesseract falla, error != None."""
    try:
        from PIL import Image
        import pytesseract
        import io as _io

        if isinstance(path_or_bytes, (bytes, bytearray)):
            img = Image.open(_io.BytesIO(bytes(path_or_bytes)))
        else:
            img = Image.open(path_or_bytes)
        if img.mode != "RGB":
            img = img.convert("RGB")
        # Reescalar imágenes pequeñas para mejorar OCR
        w, h = img.size
        if max(w, h) < 1200:
            scale = 1200 / max(w, h)
            img = img.resize((int(w * scale), int(h * scale)))
        texto = pytesseract.image_to_string(img, lang="spa+eng")
        if not texto.strip():
            return "", None, "OCR vacío: imagen sin texto legible"
        return texto, extraer_campos(texto), None
    except Exception as e:  # noqa: BLE001
        return "", None, str(e)
