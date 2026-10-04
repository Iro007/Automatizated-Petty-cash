"""Caja Chica Pro — front mejorado + plantilla Excel 100% propia.

- Sin dependencia de base/caja_chica_base.xlsx
- La planilla se genera desde cero con excel_builder.build_caja_chica
- OCR tolerante a fallos + edición manual en tabla
"""
from __future__ import annotations

import io
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

from amount_utils import parse_monto_usuario
from excel_builder import build_caja_chica, validar_fecha_gasto
from ocr_utils import extraer_campos, ocr_imagen

# ---------------- Config ----------------
st.set_page_config(
    page_title="Caja Chica Pro",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "### Caja Chica Pro\nGenera tu relación de gastos sin plantillas externas."},
)

os.makedirs("invoices", exist_ok=True)
os.makedirs("caja_chica", exist_ok=True)

CSS = """
<style>
#MainMenu, footer, header {visibility: hidden;}
.block-container {padding-top: 1.2rem; max-width: 1250px;}
.hero {
  background: linear-gradient(135deg, #1F3864 0%, #2E75B6 60%, #5B9BD5 100%);
  border-radius: 18px; padding: 26px 28px; color: white;
  box-shadow: 0 10px 30px rgba(31,56,100,.25); margin-bottom: 18px;
}
.hero h1 {margin: 0; font-size: 2rem;}
.hero p {margin: 6px 0 0 0; opacity: .92;}
.card {
  background: white; border: 1px solid #E5E7EB; border-radius: 14px;
  padding: 16px 18px; box-shadow: 0 4px 14px rgba(0,0,0,.05);
}
.stButton > button {
  border-radius: 10px; font-weight: 700; border: none;
  background: linear-gradient(135deg, #1F3864, #2E75B6); color: white;
  padding: .6rem 1.2rem;
}
.stButton > button:hover {filter: brightness(1.08);}
.stDownloadButton > button {
  border-radius: 10px; font-weight: 700;
  background: #16a34a; color: white; border: none; padding: .6rem 1.2rem;
}
[data-testid="stMetric"] {
  background: white; border: 1px solid #E5E7EB; border-radius: 14px; padding: 12px;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

COLS = ["proveedor", "fecha", "factura", "monto_bs", "descripcion"]
PRETTY = {"proveedor": "Proveedor / Comercio", "fecha": "Fecha", "factura": "N° Factura / Ref.", "monto_bs": "Monto (Bs)", "descripcion": "Descripción"}

if "gastos" not in st.session_state:
    st.session_state.gastos = pd.DataFrame(columns=COLS)
if "last_file" not in st.session_state:
    st.session_state.last_file = None
if "editor_version" not in st.session_state:
    st.session_state.editor_version = 0
if "monto_editor_invalid_rows" not in st.session_state:
    st.session_state.monto_editor_invalid_rows = []


def df_gastos() -> pd.DataFrame:
    df = st.session_state.gastos
    if df.empty:
        return pd.DataFrame(columns=COLS)
    df = df.copy()
    df["monto_bs"] = pd.to_numeric(df["monto_bs"], errors="coerce").fillna(0.0)
    return df


def totales(df: pd.DataFrame, tasa: float):
    total_bs = float(df["monto_bs"].sum()) if not df.empty else 0.0
    total_usd = (total_bs / tasa) if tasa else 0.0
    return total_bs, total_usd


# ---------------- Sidebar ----------------
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    with st.expander("🏢 Empresa (sale en el Excel)", expanded=False):
        emp_nombre = st.text_input("Nombre", value="MI EMPRESA C.A.")
        emp_rif = st.text_input("RIF", value="RIF: J-00000000-0")
        emp_dir = st.text_input("Dirección", value="Av. Principal, Ciudad, País")
        emp_tel = st.text_input("Teléfonos", value="Tel: +58 000 000 0000")
    with st.expander("👤 Responsable y montos", expanded=True):
        responsable = st.text_input("Responsable", placeholder="Nombre y apellido")
        cedula = st.text_input("Cédula", placeholder="V-12345678", max_chars=12)
        monto_otorgado = st.number_input("Monto otorgado ($)", min_value=0.0, format="%.2f")
        tasa = st.number_input("Tasa Bs/$ del día", min_value=0.0, format="%.2f", help="Se usa para convertir cada gasto a $ dentro del Excel con fórmulas vivas.")
        fecha_emision = st.date_input("Fecha de emisión", value=date.today())
        n_reporte = st.text_input("N° Reporte", value=datetime.now().strftime("%Y%m%d"))
    st.divider()
    st.caption("💡 La plantilla es propia: se genera por código, sin archivos base. Cambia los datos de empresa aquí y saldrán en el Excel.")
    if st.button("🗑️ Vaciar tabla de gastos"):
        st.session_state.gastos = pd.DataFrame(columns=COLS)
        st.session_state.editor_version += 1
        st.session_state.monto_editor_invalid_rows = []
        st.rerun()

# ---------------- Hero ----------------
st.markdown(
    """<div class="hero">
    <h1>🧾 Caja Chica Pro</h1>
    <p>Sube tus comprobantes → revisa y corrige → genera tu planilla Excel propia con totales y firmas.</p>
    </div>""",
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs(["📤 1. Cargar", "🧾 2. Revisar y editar", "📊 3. Resumen", "📥 4. Generar Excel"])

# ---------------- TAB 1 ----------------
with tab1:
    c1, c2 = st.columns([1.2, 1], gap="large")
    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Sube comprobantes")
        files = st.file_uploader(
            "Arrastra imágenes de pago móvil / facturas",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            help="También puedes agregar gastos manualmente abajo si el OCR falla.",
        )
        monto_invalido = bool(st.session_state.get("monto_editor_invalid_rows"))
        if monto_invalido:
            st.info("Corrige los importes inválidos en la pestaña **Revisar y editar** antes de procesar más comprobantes.")
        if files:
            st.write(f"**{len(files)}** archivo(s) listos para procesar:")
            cols = st.columns(3)
            for i, f in enumerate(files):
                with cols[i % 3]:
                    st.image(f, caption=f.name, use_container_width=True)
        procesar = st.button("🔍 Procesar con OCR y agregar", disabled=not files or monto_invalido)
        st.markdown("</div>", unsafe_allow_html=True)

        if procesar and files:
            nuevos, fallos = [], []
            prog = st.progress(0, text="Leyendo imágenes…")
            for i, f in enumerate(files):
                raw = f.getvalue()
                texto, campos, err = ocr_imagen(raw, f.name)
                if campos and (campos.get("monto_bs") or campos.get("factura")):
                    campos["descripcion"] = f.name
                    nuevos.append(campos)
                else:
                    fallos.append(f.name)
                    nuevos.append(
                        {"proveedor": "Revisar manual", "fecha": date.today().strftime("%d/%m/%Y"), "factura": "", "monto_bs": 0.0, "descripcion": f.name}
                    )
                prog.progress((i + 1) / len(files), text=f"Procesando {i+1}/{len(files)}…")
            prog.empty()
            if nuevos:
                st.session_state.gastos = pd.concat([df_gastos(), pd.DataFrame(nuevos, columns=COLS)], ignore_index=True)
                st.session_state.editor_version += 1
            st.success(f"✅ {len(nuevos)} gasto(s) agregados. Ve a la pestaña **Revisar y editar**.")
            if fallos:
                st.warning(f"⚠️ OCR incompleto en: {', '.join(fallos)}. Complétalos manualmente en la pestaña 2.")

    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("➕ Agregar gasto manual")
        with st.form("manual", clear_on_submit=True):
            m_prov = st.text_input("Proveedor / Comercio*")
            m_fec = st.text_input("Fecha (dd/mm/aaaa)", value=date.today().strftime("%d/%m/%Y"))
            m_fac = st.text_input("N° Factura / Referencia")
            m_monto = st.text_input("Monto (Bs)*", placeholder="749,50 o 749.50", help="Acepta coma o punto decimal; también formatos como 1.250,50 o 1,250.50.")
            m_desc = st.text_input("Descripción (opcional)")
            monto_invalido = bool(st.session_state.get("monto_editor_invalid_rows"))
            if monto_invalido:
                st.warning("Corrige los importes inválidos de la tabla antes de agregar otro gasto.")
            if st.form_submit_button("Agregar a la tabla", disabled=monto_invalido):
                if not m_prov or not m_monto:
                    st.error("Proveedor y monto son obligatorios.")
                else:
                    try:
                        fecha_manual = validar_fecha_gasto(m_fec)
                        monto_manual = parse_monto_usuario(m_monto)
                        if monto_manual <= 0:
                            raise ValueError("El monto debe ser mayor que cero.")
                    except ValueError as exc:
                        st.error(str(exc))
                    else:
                        row = pd.DataFrame([{"proveedor": m_prov.strip().capitalize(), "fecha": fecha_manual, "factura": m_fac.strip(), "monto_bs": monto_manual, "descripcion": m_desc.strip()}])
                        st.session_state.gastos = pd.concat([df_gastos(), row], ignore_index=True)
                        st.session_state.editor_version += 1
                        st.success("Gasto agregado.")
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------- TAB 2 ----------------
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Edita antes de generar")
    st.caption("Puedes corregir proveedor, fecha, referencia, monto y descripción. Para decimales puedes usar coma o punto: 749,50 o 749.50. También acepta separadores de miles: 1.250,50 o 1,250.50.")
    df = df_gastos()
    if df.empty:
        st.session_state.monto_editor_invalid_rows = []
        st.info("Aún no hay gastos. Carga imágenes en la pestaña 1 o agrega uno manual.")
    else:
        editor_key = f"editor_{st.session_state.editor_version}"
        if st.session_state.get("editor_base_version") != st.session_state.editor_version:
            st.session_state.editor_base = df.copy()
            st.session_state.editor_base_version = st.session_state.editor_version
        editor_data = st.session_state.editor_base.copy()
        editor_data["monto_bs"] = editor_data["monto_bs"].map(
            lambda value: "" if pd.isna(value) else f"{float(value):.2f}"
        )
        edited_input = st.data_editor(
            editor_data,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "proveedor": st.column_config.TextColumn(PRETTY["proveedor"], required=True),
                "fecha": st.column_config.TextColumn(PRETTY["fecha"], help="dd/mm/aaaa"),
                "factura": st.column_config.TextColumn(PRETTY["factura"]),
                "monto_bs": st.column_config.TextColumn(
                    PRETTY["monto_bs"],
                    help="Acepta coma o punto decimal, por ejemplo 749,50 o 749.50.",
                ),
                "descripcion": st.column_config.TextColumn(PRETTY["descripcion"]),
            },
            key=editor_key,
        )
        edited = edited_input.copy()
        importes, filas_invalidas = [], []
        for fila, raw_amount in enumerate(edited_input["monto_bs"].tolist(), start=1):
            try:
                importes.append(parse_monto_usuario(raw_amount))
            except ValueError:
                importes.append(0.0)
                filas_invalidas.append(fila)
        edited["monto_bs"] = importes
        st.session_state.monto_editor_invalid_rows = filas_invalidas
        # Keep the widget baseline fixed while edits accumulate or are reverted.
        st.session_state.gastos = edited.reset_index(drop=True)
        if filas_invalidas:
            filas = ", ".join(str(fila) for fila in filas_invalidas)
            st.error(f"Monto inválido en la(s) fila(s) {filas}. Usa coma o punto decimal y corrige el valor antes de generar.")
        total_bs, total_usd = totales(edited, tasa or 0)
        m1, m2, m3 = st.columns(3)
        m1.metric("N° gastos", len(edited))
        m2.metric("Total Bs", f"{total_bs:,.2f}")
        m3.metric("Total $ (ref.)", f"{total_usd:,.2f}")
        if (edited["monto_bs"] == 0).any():
            st.warning("Hay filas con monto 0. Revísalas antes de generar.")
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------- TAB 3 ----------------
with tab3:
    df = df_gastos()
    if df.empty:
        st.info("Sin datos para mostrar.")
    else:
        total_bs, total_usd = totales(df, tasa or 0)
        saldo = float(monto_otorgado or 0) - total_usd
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Otorgado ($)", f"{float(monto_otorgado or 0):,.2f}")
        c2.metric("Gastado (Bs)", f"{total_bs:,.2f}")
        c3.metric("Gastado ($)", f"{total_usd:,.2f}")
        c4.metric("Saldo ($)", f"{saldo:,.2f}", delta=f"{saldo:,.2f}")
        st.divider()
        g1, g2 = st.columns([1.4, 1])
        with g1:
            st.subheader("Gasto por proveedor (Bs)")
            chart_df = df.groupby("proveedor", as_index=False)["monto_bs"].sum().sort_values("monto_bs", ascending=False).head(15)
            st.bar_chart(chart_df.set_index("proveedor"))
        with g2:
            st.subheader("Detalle")
            show = df.copy()
            show.columns = [PRETTY[c] for c in COLS]
            st.dataframe(show, use_container_width=True, hide_index=True)

# ---------------- TAB 4 ----------------
with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Generar planilla propia")
    df = df_gastos()
    errores = []
    if not responsable:
        errores.append("Falta el **Responsable** (barra lateral).")
    if not cedula:
        errores.append("Falta la **Cédula**.")
    if not (monto_otorgado or 0):
        errores.append("Falta el **Monto otorgado ($)**.")
    if not (tasa or 0):
        errores.append("Falta la **Tasa Bs/$** (se usa en las fórmulas del Excel).")
    if df.empty:
        errores.append("La tabla de gastos está vacía.")
    else:
        if st.session_state.get("monto_editor_invalid_rows"):
            errores.append("Hay montos inválidos en la tabla. Corrígelos en **Revisar y editar** antes de generar el archivo.")
        for idx, valor in enumerate(df["fecha"], start=1):
            try:
                validar_fecha_gasto(valor)
            except ValueError as exc:
                errores.append(f"Gasto {idx}: {exc} Corrige la fecha en **Revisar y editar**.")
    if errores:
        for e in errores:
            st.error(e)
    else:
        total_bs, total_usd = totales(df, tasa)
        st.success(f"Listo: {len(df)} gasto(s) • Total Bs {total_bs:,.2f} • Total $ {total_usd:,.2f} • Saldo $ {float(monto_otorgado)-total_usd:,.2f}")
        if st.button("📄 Generar archivo Excel", type="primary"):
            empresa = {"nombre": emp_nombre, "rif": emp_rif, "direccion": emp_dir, "telefonos": emp_tel}
            wb = build_caja_chica(
                df,
                empresa=empresa,
                responsable=responsable,
                cedula=cedula,
                fecha_emision=fecha_emision,
                monto_otorgado_usd=float(monto_otorgado),
                tasa_bcv=float(tasa),
                n_reporte=n_reporte,
            )
            fname = f"caja_chica_{n_reporte}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
            out_path = os.path.join("caja_chica", fname)
            wb.save(out_path)
            buf = io.BytesIO()
            wb.save(buf)
            st.session_state.last_file = (fname, buf.getvalue(), out_path)
            st.success(f"Guardado en: `{out_path}`")
        if st.session_state.last_file:
            fname, data, out_path = st.session_state.last_file
            st.download_button(
                "⬇️ Descargar caja chica (.xlsx)",
                data=data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.caption("El Excel trae fórmulas vivas: si cambias la tasa en la celda F8, los $ se recalculan solos. Revisa y firma al final.")
    st.markdown("</div>", unsafe_allow_html=True)

st.divider()
st.caption("Caja Chica Pro • plantilla propia generada por código (sin base externa) • verifica siempre los montos antes de firmar.")
