"""Reproducible real Streamlit AppTest checks; no mocks for workbook generation."""
from pathlib import Path
import sys
import json
import math
from datetime import date
import pandas as pd
import openpyxl
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from excel_builder import build_caja_chica
OUT = ROOT / 'evidence/demo_20261003'
OUT.mkdir(parents=True, exist_ok=True)
checks = []
def record(name, result, details):
    checks.append({"name": name, "passed": bool(result), "details": details})
def app():
    return AppTest.from_file(str(ROOT / "app.py"), default_timeout=30).run()
at = app()
record("app_initial_render", len(at.exception) == 0, [str(x.value) for x in at.exception])
record("required_generation_inputs", len(at.error) == 5, [x.value for x in at.error])
next(x for x in at.button if x.label == "Agregar a la tabla").click().run()
record("empty_manual_entry_rejected", any("obligatorios" in x.value for x in at.error), [x.value for x in at.error])
for x in at.text_input:
    if x.label == "Proveedor / Comercio*": x.set_value("Proveedor ficticio")
    if x.label == "Fecha (dd/mm/aaaa)": x.set_value("31/02/2026")
next(x for x in at.number_input if x.label == "Monto (Bs)*").set_value(10.0)
next(x for x in at.button if x.label == "Agregar a la tabla").click().run()
record("invalid_manual_date_rejected", len(at.session_state["gastos"]) == 0 and any("Fecha inválida" in x.value for x in at.error), [x.value for x in at.error])
rows = [
    {"proveedor":"Papelería de ejemplo", "fecha":"03/10/2026", "factura":"100001", "monto_bs":1250.50, "descripcion":"Materiales ficticios"},
    {"proveedor":"Transporte de ejemplo", "fecha":"03/10/2026", "factura":"100002", "monto_bs":749.50, "descripcion":"Monto revisado contra comprobante sintético"},
    {"proveedor":"Servicio de ejemplo", "fecha":"03/10/2026", "factura":"100003", "monto_bs":500.0, "descripcion":"Gasto ficticio agregado manualmente"},
]
at = app()
at.session_state["gastos"] = pd.DataFrame(rows)
for x in at.text_input:
    values = {"Nombre":"DEMOSTRACIÓN FICTICIA", "RIF":"SIN VALIDEZ FISCAL", "Dirección":"Datos sintéticos", "Teléfonos":"", "Responsable":"Responsable de ejemplo", "Cédula":"DEMO-000001", "N° Reporte":"DEMO-PRUEBAS-20261003"}
    if x.label in values: x.set_value(values[x.label])
for x in at.number_input:
    if x.label == "Monto otorgado ($)": x.set_value(100.0)
    if x.label == "Tasa Bs/$ del día": x.set_value(50.0)
at.run()
at.session_state["gastos"] = pd.DataFrame([{**rows[0], "fecha":"31/02/2026"}])
at.session_state["editor_version"] += 1
at.run()
record("invalid_edited_date_blocks_generation", any("Gasto 1: Fecha inválida" in x.value for x in at.error) and not any(x.label == "📄 Generar archivo Excel" for x in at.button), [x.value for x in at.error])
at.session_state["gastos"] = pd.DataFrame(rows)
at.session_state["editor_version"] += 1
at.run()
record("corrected_date_reenables_generation", not any("Fecha inválida" in x.value for x in at.error) and any(x.label == "📄 Generar archivo Excel" for x in at.button), [x.value for x in at.error])
record("valid_inputs_enable_generation", any(x.label == "📄 Generar archivo Excel" for x in at.button), [x.value for x in at.error])
next(x for x in at.button if x.label == "📄 Generar archivo Excel").click().run()
record("app_generates_workbook", len(at.exception) == 0 and at.session_state["last_file"] is not None, [str(x.value) for x in at.exception])
name, data, source_path = at.session_state["last_file"]
example = OUT / "caja_chica_ejemplo.xlsx"
example.write_bytes(data)
wb = openpyxl.load_workbook(example, data_only=False)
ws = wb.active
record("actual_workbook_values", [ws[f"F{r}"].value for r in range(11,14)] == [1250.50,749.50,500.0] and ws["F8"].value == 50 and ws["C8"].value == 100, {"amounts":[ws[f"F{r}"].value for r in range(11,14)], "rate":ws["F8"].value, "fund":ws["C8"].value})
formulas = {coord:ws[coord].value for coord in ["G11","G12","G13","F14","G14","G15"]}
record("formula_references", formulas == {"G11":"=IF($F$8=0,0,F11/$F$8)","G12":"=IF($F$8=0,0,F12/$F$8)","G13":"=IF($F$8=0,0,F13/$F$8)","F14":"=SUM(F11:F13)","G14":"=SUM(G11:G13)","G15":"=$C$8-G14"}, formulas)
record("independent_totals", math.isclose(sum(r["monto_bs"] for r in rows),2500) and math.isclose(sum(r["monto_bs"] for r in rows)/50,50), {"bs":2500,"usd":50,"balance":50,"note":"Independent Python calculation; native Excel recalculation not yet verified"})
for count in [0,1,30,1000]:
    w = build_caja_chica(rows[:1]*count,tasa_bcv=50,monto_otorgado_usd=100)
    s=w.active; end=10+max(count,1); total=end+1
    record(f"row_boundary_{count}", s[f"F{total}"].value == f"=SUM(F11:F{end})" and s[f"G{end}"].value == f"=IF($F$8=0,0,F{end}/$F$8)", {"actual_rows":count,"total_row":total,"sum_formula":s[f"F{total}"].value})
next(x for x in at.number_input if x.label == "Tasa Bs/$ del día").set_value(0).run()
record("zero_rate_blocks_generation", not any(x.label == "📄 Generar archivo Excel" for x in at.button), [x.value for x in at.error])
at.session_state["gastos"] = pd.DataFrame([{**rows[0],"monto_bs":0.0}])
at.session_state["editor_version"] += 1
next(x for x in at.number_input if x.label == "Tasa Bs/$ del día").set_value(50).run()
record("zero_row_only_warns_limitation", any("monto 0" in x.value for x in at.warning) and any(x.label == "📄 Generar archivo Excel" for x in at.button), "Zero amount row warns, but does not block export")
result={"checks":checks,"all_checks_passed":all(c["passed"] for c in checks),"example":example.name,"source_generated_path":str(source_path),"native_excel_recalculation":"not_verified","rate":"Synthetic example, not a current exchange rate"}
(OUT / "test_results.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(result,ensure_ascii=False,indent=2))
if not result["all_checks_passed"]: raise SystemExit(1)
