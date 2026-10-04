"""Regression tests for date rejection and literal XLSX user text."""
from pathlib import Path
from datetime import date, datetime
import io
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from amount_utils import parse_monto_usuario
from excel_builder import build_caja_chica, validar_fecha_gasto
from ocr_utils import extraer_campos
OUT=ROOT / 'evidence/demo_20261003'
OUT.mkdir(parents=True, exist_ok=True)
checks=[]
def check(name,passed,details):
    checks.append({"name":name,"passed":bool(passed),"details":details})
def rejects(call):
    try: call()
    except ValueError as exc: return str(exc)
    return None
row={"proveedor":"Proveedor de ejemplo","fecha":"03/10/2026","factura":"000001","monto_bs":100.0,"descripcion":"Datos ficticios"}
ocr_text = "CONCEPTO Transporte de ejemplo\nTOTAL Bs 749,50\nCOMISION INCLUIDA Bs 25,00"
ocr_fields = extraer_campos(ocr_text)
check("prefer_labeled_total_over_generic_commission", ocr_fields["monto_bs"] == 749.50, ocr_fields)
for amount_text, expected in [
    ("749,50", 749.50), ("749.50", 749.50),
    ("1.250,50", 1250.50), ("1,250.50", 1250.50),
    ("1.250", 1250.0), ("1,250", 1250.0),
    ("749.50 Bs", 749.50), ("", 0.0),
]:
    actual = parse_monto_usuario(amount_text)
    check(f"amount_separator_{amount_text or 'blank'}", actual == expected, {"input":amount_text,"actual":actual,"expected":expected})
for invalid_amount in ["abc", "1,23,45", "1.234,567", "-749,50"]:
    try:
        parse_monto_usuario(invalid_amount)
        rejected = False
    except ValueError:
        rejected = True
    check(f"amount_invalid_{invalid_amount}", rejected, {"input":invalid_amount,"rejected":rejected})
for value in ["31/02/2026","31/04/2026","29/02/2025","00/10/2026","03/13/2026","03/10/0000","","   ",None,"2026-10-03","3/10/2026",pd.NaT]:
    message=rejects(lambda:validar_fecha_gasto(value))
    check(f"date_reject_{value!s}",message is not None and "Fecha inválida" in message,{"input":str(value),"error":message})
    for kind,values in [("list",[{**row,"fecha":value}]),("DataFrame",pd.DataFrame([{**row,"fecha":value}]))]:
        message=rejects(lambda:build_caja_chica(values,tasa_bcv=50,monto_otorgado_usd=100))
        check(f"builder_{kind}_reject_{value!s}",message is not None and "Gasto 1:" in message,{"input":str(value),"error":message})
for value,expected in [("29/02/2024","29/02/2024"),("03/10/2026","03/10/2026"),(" 03/10/2026 ","03/10/2026"),(date(2026,10,3),"03/10/2026"),(datetime(2026,10,3,12,0),"03/10/2026")]:
    actual=validar_fecha_gasto(value)
    check(f"date_accept_{value!s}",actual==expected,{"input":str(value),"actual":actual})
message=rejects(lambda:build_caja_chica([row],fecha_emision="31/02/2026"))
check("invalid_issue_date_rejected",message is not None and "Fecha de emisión:" in message,message)
ns={"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
for text in ['=1+1','=HYPERLINK("https://example.invalid","demo")','+1+1','-1+1','@SUM(1,1)',' =1+1','\t=1+1',"'=1+1",'#N/A','Texto normal con acentos: revisión']:
    source={**row,"proveedor":text,"factura":text,"descripcion":text}
    company={"nombre":text,"rif":text,"direccion":text,"telefonos":text}
    w=build_caja_chica([source],empresa=company,responsable=text,cedula=text,titulo=text,fecha_emision=date(2026,10,3),tasa_bcv=50,monto_otorgado_usd=100)
    buf=io.BytesIO();w.save(buf);raw=buf.getvalue()
    reloaded=openpyxl.load_workbook(io.BytesIO(raw),data_only=False)
    s=reloaded.active
    targets={"B1":text.upper(),"B2":text,"B3":text,"B4":text,"B5":text,"C7":text,"F7":text,"C11":text,"E11":text,"H11":text}
    literal_ok=all(s[c].value==v and s[c].data_type=="s" for c,v in targets.items())
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        xml=ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    cells={c.attrib["r"]:c for c in xml.findall(".//s:c",ns)}
    xml_ok=all(cells[c].attrib.get("t")=="inlineStr" and cells[c].find("s:f",ns) is None for c in targets)
    expected={"G11":"=IF($F$8=0,0,F11/$F$8)","F12":"=SUM(F11:F11)","G12":"=SUM(G11:G11)","G13":"=$C$8-G12"}
    formulas_ok=all(s[c].value==v and s[c].data_type=="f" for c,v in expected.items())
    check(f"literal_text_{text}",literal_ok and xml_ok and formulas_ok,{"text":text,"literal_roundtrip":literal_ok,"xml_has_strings_no_formulas":xml_ok,"legitimate_formulas_preserved":formulas_ok})
    if text=='=1+1':
        (OUT/'texto_literal_ejemplo.xlsx').write_bytes(raw)
        (OUT/'literal_export_check.json').write_text(json.dumps({"user_cells":{c:{"value":s[c].value,"type":s[c].data_type,"xml_type":cells[c].attrib.get("t"),"has_formula":cells[c].find("s:f",ns) is not None} for c in targets},"legitimate_formulas":{c:{"value":s[c].value,"type":s[c].data_type} for c in expected}},ensure_ascii=False,indent=2),encoding='utf-8')
result={"all_checks_passed":all(c["passed"] for c in checks),"checks":checks,"scope":"Receipt amount priority, comma/dot amount input, date validity and literal string storage in XLSX. Not a general security assessment.","native_excel":"not_verified"}
(OUT/'regression_results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({"checks":len(checks),"passed":sum(c['passed'] for c in checks),"failed":[c for c in checks if not c['passed']]},ensure_ascii=False,indent=2))
if not result['all_checks_passed']:raise SystemExit(1)
