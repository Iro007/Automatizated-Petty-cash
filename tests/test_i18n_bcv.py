"""Regression checks for language/number formatting and BCV rate parsing."""
from datetime import date, datetime, timezone
from unittest.mock import patch
import ast
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from bcv_rate import BCVRateError, fetch_bcv_rate, parse_bcv_html
from excel_builder import build_caja_chica
from i18n import TRANSLATIONS, excel_number_format, format_number, normalize_language

checks = []


def check(name, passed, details=None):
    checks.append({"name": name, "passed": bool(passed), "details": details})


fixture_html = """
<section>
  <div id="dolar" class="col-sm-12">
    <span>USD</span><strong class="strong-tb">871,36890000</strong>
  </div>
  <div class="pull-right">Fecha Valor:
    <span class="date-display-single" content="2026-10-05T00:00:00-04:00">Lunes, 05 Octubre 2026</span>
  </div>
</section>
"""
fixed_time = datetime(2026, 10, 4, 4, 0, tzinfo=timezone.utc)
bcv = parse_bcv_html(fixture_html, fetched_at=fixed_time)
check("bcv_parses_eight_decimal_usd", bcv.rate == "871.36890000", bcv.rate)
check("bcv_parses_official_value_date", bcv.value_date == date(2026, 10, 5), bcv.value_date.isoformat())
check("bcv_keeps_fetch_timestamp", bcv.fetched_at == fixed_time, bcv.fetched_at.isoformat())

fallback_date_html = """
<div id="dolar"><strong class="strong-tb">871,36890000</strong></div>
<span class="date-display-single">Lunes, 05 Octubre 2026</span>
"""
check("bcv_parses_visible_spanish_date", parse_bcv_html(fallback_date_html).value_date == date(2026, 10, 5))

for invalid_html in [
    "<div id='dolar'><strong class='strong-tb'>0,00</strong></div><span class='date-display-single' content='2026-10-05'>05 Octubre 2026</span>",
    "<div id='other'><strong class='strong-tb'>871,36890000</strong></div><span class='date-display-single' content='2026-10-05'>05 Octubre 2026</span>",
    "<div id='dolar'><strong class='strong-tb'>not a rate</strong></div><span class='date-display-single' content='2026-10-05'>05 Octubre 2026</span>",
    "<div id='dolar'><strong class='strong-tb'>871,36890000</strong></div>",
]:
    try:
        parse_bcv_html(invalid_html)
        rejected = False
    except BCVRateError:
        rejected = True
    check(f"bcv_rejects_invalid_page_{len(checks)}", rejected)

with patch("bcv_rate.urlopen", side_effect=TimeoutError):
    try:
        fetch_bcv_rate(timeout=0.01)
        timeout_rejected = False
    except BCVRateError:
        timeout_rejected = True
check("bcv_timeout_returns_controlled_error", timeout_rejected)


class HTTPFailure:
    status = 503

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


with patch("bcv_rate.urlopen", return_value=HTTPFailure()):
    try:
        fetch_bcv_rate()
        http_rejected = False
    except BCVRateError:
        http_rejected = True
check("bcv_non_200_response_returns_controlled_error", http_rejected)

check("browser_locale_english", normalize_language("en-US") == "en")
check("browser_locale_spanish_venezuela", normalize_language("es-VE") == "es")
check("unknown_browser_locale_falls_back_to_spanish", normalize_language("fr-FR") == "es")
check("spanish_number_format", format_number(1250.5, "es", 2) == "1.250,50")
check("english_number_format", format_number(1250.5, "en", 2) == "1,250.50")
check("spanish_rate_keeps_eight_decimals", format_number("871.36890000", "es", 8) == "871,36890000")
check("english_rate_keeps_eight_decimals", format_number("871.36890000", "en", 8) == "871.36890000")
check("translation_catalogs_have_same_keys", set(TRANSLATIONS["es"]) == set(TRANSLATIONS["en"]))
used_translation_keys = set()
for source_file in (ROOT / "app.py", ROOT / "excel_builder.py"):
    tree = ast.parse(source_file.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "tr" and len(node.args) > 1:
            if isinstance(node.args[1], ast.Constant) and isinstance(node.args[1].value, str):
                used_translation_keys.add(node.args[1].value)
missing_translation_keys = used_translation_keys - set(TRANSLATIONS["es"])
check("all_literal_ui_and_excel_keys_are_translated", not missing_translation_keys, sorted(missing_translation_keys))
check("excel_spanish_number_format_locale", "0C0A" in excel_number_format("es", 8))
check("excel_english_number_format_locale", "0409" in excel_number_format("en", 8))

row = [{"proveedor": "Transporte de ejemplo", "fecha": "03/10/2026", "factura": "100002", "monto_bs": 749.5, "descripcion": "Dato ficticio"}]
spanish_book = build_caja_chica(
    row, responsable="Persona de ejemplo", cedula="DEMO", fecha_emision=date(2026, 10, 3),
    monto_otorgado_usd=100, tasa_bcv=871.3689, n_reporte="DEMO-ES", idioma="es",
    fecha_valor_bcv=date(2026, 10, 5), origen_tasa="bcv",
)
spanish_sheet = spanish_book.active
check("excel_spanish_labels_and_rate_source", spanish_sheet["B5"].value == "RELACIÓN DE GASTOS — CAJA CHICA" and "Fecha Valor: 05/10/2026" in spanish_sheet["B9"].value, {"title":spanish_sheet["B5"].value,"rate_note":spanish_sheet["B9"].value})
check("excel_rate_retains_eight_decimal_format", spanish_sheet["F8"].number_format == excel_number_format("es", 8), spanish_sheet["F8"].number_format)
check("excel_formula_reference_preserved_es", spanish_sheet["G11"].value == "=IF($F$8=0,0,F11/$F$8)", spanish_sheet["G11"].value)

english_book = build_caja_chica(
    row, responsable="Example person", cedula="DEMO", fecha_emision=date(2026, 10, 3),
    monto_otorgado_usd=100, tasa_bcv=871.3689, n_reporte="DEMO-EN", idioma="en",
    fecha_valor_bcv=date(2026, 10, 5), origen_tasa="ajuste_manual", tasa_bcv_original="871.36890000",
)
english_sheet = english_book.active
check("excel_english_labels_and_manual_reference", english_sheet["B5"].value == "PETTY CASH EXPENSE REPORT" and "manual adjustment" in english_sheet["B9"].value.lower() and "Original BCV reference: 871.36890000" in english_sheet["B9"].value and "Value date: 05/10/2026" in english_sheet["B9"].value, {"title":english_sheet["B5"].value,"rate_note":english_sheet["B9"].value})
check("excel_formula_reference_preserved_en", english_sheet["G11"].value == "=IF($F$8=0,0,F11/$F$8)", english_sheet["G11"].value)

adjusted_failed_book = build_caja_chica(
    row, responsable="Example person", cedula="DEMO", fecha_emision=date(2026, 10, 3),
    monto_otorgado_usd=100, tasa_bcv=50, n_reporte="DEMO-EN-STALE", idioma="en",
    fecha_valor_bcv=date(2026, 10, 5), origen_tasa="ajuste_manual",
    tasa_bcv_original="871.36890000", bcv_actualizacion_fallida=True,
)
adjusted_failed_note = adjusted_failed_book.active["B9"].value
check(
    "excel_adjustment_warns_when_latest_bcv_refresh_failed",
    "manual adjustment" in adjusted_failed_note.lower()
    and "871.36890000" in adjusted_failed_note
    and "latest BCV refresh failed" in adjusted_failed_note,
    adjusted_failed_note,
)

stale_book = build_caja_chica(
    row, responsable="Example person", cedula="DEMO", fecha_emision=date(2026, 10, 3),
    monto_otorgado_usd=100, tasa_bcv=871.3689, n_reporte="DEMO-EN-STALE", idioma="en",
    fecha_valor_bcv=date(2026, 10, 5), origen_tasa="bcv_stale",
)
stale_note = stale_book.active["B9"].value
check(
    "excel_does_not_present_stale_bcv_rate_as_current",
    "last value received from bcv" in stale_note.lower()
    and "latest refresh failed" in stale_note
    and "05/10/2026" in stale_note,
    stale_note,
)

result = {"passed": sum(item["passed"] for item in checks), "total": len(checks), "failed": [item for item in checks if not item["passed"]], "checks": checks}
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
if result["failed"]:
    raise SystemExit(1)
