"""Caja Chica Pro — front mejorado + plantilla Excel 100% propia.

- Sin dependencia de base/caja_chica_base.xlsx
- La planilla se genera desde cero con excel_builder.build_caja_chica
- OCR tolerante a fallos + edición manual en tabla
"""
from __future__ import annotations

import io
import os
from datetime import date, datetime
from decimal import Decimal

import pandas as pd
import streamlit as st

from amount_utils import parse_monto_usuario, parse_tasa_usuario
from bcv_rate import BCV_URL, fetch_bcv_rate
from excel_builder import build_caja_chica, validar_fecha_gasto
from i18n import format_number, normalize_language, tr
from ocr_utils import ocr_imagen
from ui import app_css, brand_header, section_header, workspace_intro, workspace_summary, workflow_guide


def _detect_browser_language() -> str:
    try:
        return normalize_language(st.context.locale)
    except Exception:  # Streamlit context may be unavailable in test harnesses.
        return "es"


initial_page_language = normalize_language(
    st.session_state.get("language") or _detect_browser_language()
)

# ---------------- Config ----------------
st.set_page_config(
    page_title=tr(initial_page_language, "app_title"),
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

os.makedirs("invoices", exist_ok=True)
os.makedirs("caja_chica", exist_ok=True)

COLS = ["proveedor", "fecha", "factura", "monto_bs", "descripcion"]


def pretty_columns(language: str) -> dict[str, str]:
    return {
        "proveedor": tr(language, "supplier_column"),
        "fecha": tr(language, "date_column"),
        "factura": tr(language, "invoice_column"),
        "monto_bs": tr(language, "amount_column"),
        "descripcion": tr(language, "description_column"),
    }


@st.cache_data(ttl=900, show_spinner=False)
def _cached_bcv_rate() -> dict[str, str]:
    result = fetch_bcv_rate()
    return {
        "rate": result.rate,
        "value_date": result.value_date.isoformat(),
        "fetched_at": result.fetched_at.isoformat(),
        "source_url": result.source_url,
    }


if "gastos" not in st.session_state:
    st.session_state.gastos = pd.DataFrame(columns=COLS)
if "last_file" not in st.session_state:
    st.session_state.last_file = None
if "editor_version" not in st.session_state:
    st.session_state.editor_version = 0
if "monto_editor_invalid_rows" not in st.session_state:
    st.session_state.monto_editor_invalid_rows = []
if "language" not in st.session_state:
    st.session_state.language = _detect_browser_language()
if "language_last" not in st.session_state:
    st.session_state.language_last = st.session_state.language
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False
if "bcv_lookup_attempted" not in st.session_state:
    st.session_state.bcv_lookup_attempted = False
if "rate_source" not in st.session_state:
    st.session_state.rate_source = "manual"
if "rate_input" not in st.session_state:
    st.session_state.rate_input = ""
if "monto_otorgado_input" not in st.session_state:
    st.session_state.monto_otorgado_input = format_number(0, st.session_state.language, 2)
if "bcv_value_date" not in st.session_state:
    st.session_state.bcv_value_date = None
if "bcv_reference_rate" not in st.session_state:
    st.session_state.bcv_reference_rate = None
if "bcv_last_error" not in st.session_state:
    st.session_state.bcv_last_error = False


def set_bcv_rate(info: dict[str, str], language: str) -> None:
    st.session_state.bcv_reference_rate = info["rate"]
    st.session_state.bcv_value_date = info["value_date"]
    st.session_state.bcv_fetched_at = info["fetched_at"]
    st.session_state.bcv_source_url = info["source_url"]
    st.session_state.rate_input = format_number(Decimal(info["rate"]), language, 8)
    st.session_state.rate_source = "bcv"
    st.session_state.bcv_last_error = False


def initialize_bcv_rate(language: str) -> None:
    if st.session_state.bcv_lookup_attempted:
        return
    st.session_state.bcv_lookup_attempted = True
    with st.spinner(tr(language, "rate_loading")):
        try:
            set_bcv_rate(_cached_bcv_rate(), language)
        except Exception:
            st.session_state.bcv_last_error = True
            if not st.session_state.rate_input:
                st.session_state.rate_source = "manual"


def refresh_bcv_rate() -> None:
    # Callbacks run before widgets are recreated, so rate_input remains editable.
    _cached_bcv_rate.clear()
    try:
        set_bcv_rate(_cached_bcv_rate(), st.session_state.language)
    except Exception:
        st.session_state.bcv_last_error = True


def remember_input_method(widget_key: str) -> None:
    method = st.session_state.get(widget_key)
    if method in ("receipts", "manual"):
        st.session_state.input_method = method


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


def expense_count_caption(language: str, count: int) -> str:
    return tr(language, "expense_count_single" if count == 1 else "expense_count_short", count=count)


# ---------------- Compact preferences ----------------
if st.session_state.get("input_method") not in ("receipts", "manual"):
    old_method = st.session_state.get("upload_method")
    st.session_state.input_method = old_method if old_method in ("receipts", "manual") else "receipts"
if "manual_expense_date" not in st.session_state:
    st.session_state.manual_expense_date = date.today().strftime("%d/%m/%Y")
if st.session_state.pop("manual_reset_pending", False):
    for draft_key in ("manual_supplier", "manual_invoice_ref", "manual_amount", "manual_description"):
        st.session_state[draft_key] = ""
    st.session_state.manual_expense_date = date.today().strftime("%d/%m/%Y")
# Keep an unfinished manual entry when the user switches input methods.
for draft_key in ("manual_supplier", "manual_expense_date", "manual_invoice_ref", "manual_amount", "manual_description"):
    if draft_key in st.session_state:
        st.session_state[draft_key] = st.session_state[draft_key]

with st.container(key="app_header"):
    title_slot, language_slot, theme_slot = st.columns([5, 1.2, 1.5], gap="small")
    with language_slot:
        language_before = st.session_state.language_last
        language = st.radio(
            tr(st.session_state.language, "language"),
            options=["es", "en"],
            format_func=lambda code: code.upper(),
            horizontal=True,
            key="language",
            label_visibility="collapsed",
        )
        if language != language_before:
            try:
                old_grant = parse_monto_usuario(st.session_state.get("monto_otorgado_input", ""), language_before)
                st.session_state.monto_otorgado_input = format_number(old_grant, language, 2)
            except ValueError:
                pass
            try:
                old_rate = parse_tasa_usuario(st.session_state.get("rate_input", ""), language_before)
                st.session_state.rate_input = format_number(old_rate, language, 8)
            except ValueError:
                pass
            st.session_state.language_last = language
            if not st.session_state.get("monto_editor_invalid_rows"):
                st.session_state.editor_version += 1
            st.session_state.last_file = None

    with theme_slot:
        dark_mode = st.toggle(tr(language, "dark_mode"), key="dark_mode")
    with title_slot:
        st.markdown(
            brand_header(language),
            unsafe_allow_html=True,
        )

st.markdown(app_css(dark_mode), unsafe_allow_html=True)
PRETTY = pretty_columns(language)
initialize_bcv_rate(language)

tab1, tab2, tab3, tab4 = st.tabs([
    tr(language, "tab_upload"), tr(language, "tab_review"),
    tr(language, "tab_summary"), tr(language, "tab_export"),
])

# Render report widgets before calculating totals, but place them in Export.
with tab4:
    with st.container(key="report_panel"):
        st.markdown(section_header(language, "report_details", "receipt"), unsafe_allow_html=True)
        st.caption(tr(language, "export_setup_intro"))
        person_col, id_col, grant_col = st.columns(3)
        with person_col:
            responsable = st.text_input(tr(language, "responsible"), key="responsable")
        with id_col:
            cedula = st.text_input(tr(language, "id_number"), placeholder="V-12345678", max_chars=12, key="cedula")
        with grant_col:
            monto_otorgado_raw = st.text_input(tr(language, "amount_granted"), key="monto_otorgado_input", help=tr(language, "amount_help"))
        rate_col, date_col, report_col = st.columns(3)
        with rate_col:
            tasa_raw = st.text_input(tr(language, "bcv_rate"), key="rate_input", help=tr(language, "rate_help"))
        with date_col:
            fecha_emision_raw = st.text_input(tr(language, "issue_date"), value=date.today().strftime("%d/%m/%Y"), key="issue_date_input")
        with report_col:
            n_reporte = st.text_input(tr(language, "report_number"), value=datetime.now().strftime("%Y%m%d"), key="n_reporte")
        try:
            applied_rate = parse_tasa_usuario(tasa_raw, language)
        except ValueError:
            applied_rate = Decimal("0")
        original_rate = st.session_state.get("bcv_reference_rate")
        is_adjusted = original_rate is not None and applied_rate != Decimal(original_rate)
        with st.expander(tr(language, "rate_details")):
            st.button(tr(language, "rate_refresh"), key="refresh_bcv_rate", on_click=refresh_bcv_rate)
            st.link_button(tr(language, "bcv_source_link"), BCV_URL)
            if st.session_state.get("bcv_value_date"):
                effective_date = date.fromisoformat(st.session_state.bcv_value_date).strftime("%d/%m/%Y")
                original = format_number(Decimal(st.session_state.bcv_reference_rate), language, 8)
                st.caption(tr(language, "rate_adjusted_reference", rate=original, date=effective_date))
        if st.session_state.bcv_last_error:
            st.warning(tr(language, "rate_fetch_failed"))
            if st.session_state.get("bcv_value_date"):
                st.caption(tr(language, "rate_bcv_stale_reference", date=date.fromisoformat(st.session_state.bcv_value_date).strftime("%d/%m/%Y")))
        if applied_rate > 0 and is_adjusted:
            st.caption(tr(language, "rate_source_adjusted"))
        elif applied_rate > 0 and original_rate is None:
            st.caption(tr(language, "rate_source_manual"))
        elif st.session_state.get("bcv_value_date"):
            if not st.session_state.bcv_last_error and applied_rate > 0:
                st.caption(tr(language, "rate_bcv_reference", date=date.fromisoformat(st.session_state.bcv_value_date).strftime("%d/%m/%Y")))
        with st.expander(tr(language, "company_details")):
            company_col, contact_col = st.columns(2)
            with company_col:
                emp_nombre = st.text_input(tr(language, "company_name"), value=tr(language, "default_company_name"), key="empresa_nombre")
                emp_rif = st.text_input(tr(language, "rif"), value=tr(language, "default_rif"), key="empresa_rif")
            with contact_col:
                emp_dir = st.text_input(tr(language, "address"), value=tr(language, "default_address"), key="empresa_direccion")
                emp_tel = st.text_input(tr(language, "phones"), value=tr(language, "default_phones"), key="empresa_telefonos")
        export_actions = st.container()
try:
    monto_otorgado = parse_monto_usuario(monto_otorgado_raw, language)
    grant_error = False
except ValueError:
    monto_otorgado, grant_error = None, True
try:
    tasa_decimal = parse_tasa_usuario(tasa_raw, language)
    rate_error = False
except ValueError:
    tasa_decimal, rate_error = Decimal("0"), True
tasa = float(tasa_decimal)
original_rate = st.session_state.get("bcv_reference_rate")
if tasa_decimal > 0 and original_rate is not None:
    st.session_state.rate_source = (
        ("bcv_stale" if st.session_state.bcv_last_error else "bcv")
        if tasa_decimal == Decimal(original_rate) else "ajuste_manual"
    )
else:
    st.session_state.rate_source = "manual"
try:
    fecha_emision = validar_fecha_gasto(fecha_emision_raw, language)
    issue_date_error = False
except ValueError:
    fecha_emision, issue_date_error = fecha_emision_raw, True

# ---------------- Add expenses ----------------
with tab1:
    intro_column, status_column = st.columns([2.4, 1], gap="medium")
    with intro_column:
        st.markdown(workspace_intro(language), unsafe_allow_html=True)
    with status_column:
        status_placeholder = st.empty()
    with st.container(key="input_switch"):
        method_key = f"upload_method_{language}"
        method_labels = {
            "receipts": tr(language, "upload_method_receipts"),
            "manual": tr(language, "upload_method_manual"),
        }
        selected_method = st.radio(
            tr(language, "upload_method"), ["receipts", "manual"], horizontal=True,
            index=0 if st.session_state.input_method == "receipts" else 1,
            format_func=lambda method: method_labels.get(method, str(method)),
            label_visibility="collapsed", key=method_key,
            on_change=remember_input_method, args=(method_key,),
        )
        method = selected_method if selected_method in method_labels else st.session_state.input_method
    if method == "receipts":
        upload_column, guide_column = st.columns([1.75, 1], gap="medium")
        with upload_column:
            with st.container(key="receipt_panel"):
                st.markdown(section_header(language, "upload_title", "upload"), unsafe_allow_html=True)
                st.caption(tr(language, "upload_intro"))
                files = st.file_uploader(
                    tr(language, "upload_label"),
                    type=["png", "jpg", "jpeg", "webp"],
                    accept_multiple_files=True,
                    help=tr(language, "upload_help"),
                    key="receipt_uploader",
                )
                monto_invalido = bool(st.session_state.get("monto_editor_invalid_rows"))
                if monto_invalido:
                    st.info(tr(language, "fix_invalid_before_ocr"))
                if files:
                    with st.expander(tr(language, "files_ready", count=len(files))):
                        cols = st.columns(3)
                        for i, f in enumerate(files):
                            with cols[i % 3]:
                                st.image(f, caption=f.name, width=160)
                procesar = st.button(tr(language, "process_ocr"), type="primary", use_container_width=True, disabled=not files or monto_invalido)
                if procesar and files:
                    nuevos, fallos = [], []
                    prog = st.progress(0, text=tr(language, "processing_ocr"))
                    for i, f in enumerate(files):
                        raw = f.getvalue()
                        texto, campos, err = ocr_imagen(raw, f.name)
                        if campos and (campos.get("monto_bs") or campos.get("factura")):
                            campos["descripcion"] = f.name
                            nuevos.append(campos)
                        else:
                            fallos.append(f.name)
                            nuevos.append(
                                {"proveedor": tr(language, "review_manually"), "fecha": date.today().strftime("%d/%m/%Y"), "factura": "", "monto_bs": 0.0, "descripcion": f.name}
                            )
                        prog.progress((i + 1) / len(files), text=tr(language, "processing_file", current=i + 1, total=len(files)))
                    prog.empty()
                    if nuevos:
                        st.session_state.gastos = pd.concat([df_gastos(), pd.DataFrame(nuevos, columns=COLS)], ignore_index=True)
                        st.session_state.editor_version += 1
                    st.success(tr(language, "ocr_added", count=len(nuevos)))
                    if fallos:
                        st.warning(tr(language, "ocr_incomplete", files=", ".join(fallos)))
        with guide_column:
            st.markdown(workflow_guide(language), unsafe_allow_html=True)

    else:
        with st.container(key="manual_panel"):
            st.markdown(section_header(language, "upload_method_manual", "edit"), unsafe_allow_html=True)
            st.caption(tr(language, "manual_intro"))
            if st.session_state.pop("manual_feedback", False):
                st.success(tr(language, "manual_added"))
            with st.container(key="manual_fields"):
                supplier_col, expense_date_col, amount_col = st.columns([2, 1.2, 1.2])
                with supplier_col:
                    m_prov = st.text_input(tr(language, "supplier"), key="manual_supplier")
                with expense_date_col:
                    m_fec = st.text_input(tr(language, "expense_date"), key="manual_expense_date")
                with amount_col:
                    m_monto = st.text_input(tr(language, "amount_bs"), placeholder="749,50" if language == "es" else "749.50", help=tr(language, "amount_help"), key="manual_amount")
                with st.expander(tr(language, "optional_details")):
                    reference_col, description_col = st.columns(2)
                    with reference_col:
                        m_fac = st.text_input(tr(language, "invoice_ref"), key="manual_invoice_ref")
                    with description_col:
                        m_desc = st.text_input(tr(language, "description_optional"), key="manual_description")
                monto_invalido = bool(st.session_state.get("monto_editor_invalid_rows"))
                if monto_invalido:
                    st.warning(tr(language, "fix_invalid_before_manual"))
                if st.button(tr(language, "manual_add_button"), disabled=monto_invalido, type="primary"):
                    if not m_prov.strip() or not m_monto.strip():
                        st.error(tr(language, "required_supplier_amount"))
                    else:
                        try:
                            fecha_manual = validar_fecha_gasto(m_fec, language)
                            monto_manual = parse_monto_usuario(m_monto, language)
                            if monto_manual <= 0:
                                raise ValueError(tr(language, "amount_positive"))
                        except ValueError as exc:
                            st.error(str(exc))
                        else:
                            row = pd.DataFrame([{"proveedor": m_prov.strip(), "fecha": fecha_manual, "factura": m_fac.strip(), "monto_bs": monto_manual, "descripcion": m_desc.strip()}])
                            st.session_state.gastos = pd.concat([df_gastos(), row], ignore_index=True)
                            st.session_state.editor_version += 1
                            st.session_state.manual_reset_pending = True
                            st.session_state.manual_feedback = True
                            st.rerun()

    if not st.session_state.gastos.empty:
        st.caption(expense_count_caption(language, len(st.session_state.gastos)))

# ---------------- TAB 2 ----------------
with tab2:
    st.subheader(tr(language, "review_title"))
    st.caption(tr(language, "review_hint_short"))
    df = df_gastos()
    if df.empty:
        st.session_state.monto_editor_invalid_rows = []
        st.info(tr(language, "empty_expenses_upload"))
    else:
        editor_key = f"editor_{st.session_state.editor_version}"
        if st.session_state.get("editor_base_version") != st.session_state.editor_version:
            st.session_state.editor_base = df.copy()
            st.session_state.editor_base_version = st.session_state.editor_version
        editor_data = st.session_state.editor_base.copy()
        editor_data["monto_bs"] = editor_data["monto_bs"].map(
            lambda value: "" if pd.isna(value) else format_number(float(value), language, 2)
        )
        edited_input = st.data_editor(
            editor_data,
            num_rows="dynamic",
            use_container_width=True,
            height=min(360, 74 + len(editor_data) * 35),
            hide_index=True,
            column_config={
                "proveedor": st.column_config.TextColumn(PRETTY["proveedor"], required=True),
                "fecha": st.column_config.TextColumn(PRETTY["fecha"], help=tr(language, "date_format_help")),
                "factura": st.column_config.TextColumn(PRETTY["factura"]),
                "monto_bs": st.column_config.TextColumn(
                    PRETTY["monto_bs"],
                    help=tr(language, "amount_help"),
                ),
                "descripcion": st.column_config.TextColumn(PRETTY["descripcion"]),
            },
            key=editor_key,
        )
        edited = edited_input.copy()
        importes, filas_invalidas = [], []
        for fila, raw_amount in enumerate(edited_input["monto_bs"].tolist(), start=1):
            try:
                importes.append(parse_monto_usuario(raw_amount, language))
            except ValueError:
                importes.append(0.0)
                filas_invalidas.append(fila)
        edited["monto_bs"] = importes
        invalid_state_changed = bool(st.session_state.monto_editor_invalid_rows) != bool(filas_invalidas)
        st.session_state.monto_editor_invalid_rows = filas_invalidas
        # Keep the widget baseline fixed while edits accumulate or are reverted.
        st.session_state.gastos = edited.reset_index(drop=True)
        if invalid_state_changed:
            st.rerun()
        if filas_invalidas:
            filas = ", ".join(str(fila) for fila in filas_invalidas)
            st.error(tr(language, "invalid_amount_rows", rows=filas))
        total_bs, total_usd = totales(edited, tasa or 0)
        m1, m2, m3 = st.columns(3)
        m1.metric(tr(language, "metric_count"), len(edited))
        m2.metric(tr(language, "metric_spent_bs"), format_number(total_bs, language, 2))
        m3.metric(tr(language, "metric_spent_usd"), format_number(total_usd, language, 2) if tasa > 0 else "—")
        if (edited["monto_bs"] == 0).any():
            st.warning(tr(language, "zero_amount_warning"))

    if not df.empty:
        with st.expander(tr(language, "clear_expenses")):
            st.caption(tr(language, "reset_hint"))
            if st.button(tr(language, "reset_confirm")):
                st.session_state.gastos = pd.DataFrame(columns=COLS)
                st.session_state.editor_version += 1
                st.session_state.monto_editor_invalid_rows = []
                st.session_state.last_file = None
                st.rerun()

# ---------------- TAB 3 ----------------
with tab3:
    df = df_gastos()
    if df.empty:
        st.info(tr(language, "summary_empty"))
    else:
        total_bs, total_usd = totales(df, tasa or 0)
        saldo = float(monto_otorgado or 0) - total_usd
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(tr(language, "metric_granted"), format_number(float(monto_otorgado or 0), language, 2))
        c2.metric(tr(language, "metric_spent_bs"), format_number(total_bs, language, 2))
        c3.metric(tr(language, "metric_spent_usd"), format_number(total_usd, language, 2) if tasa > 0 else "—")
        c4.metric(tr(language, "metric_balance"), format_number(saldo, language, 2) if tasa > 0 else "—")
        if tasa <= 0:
            st.caption(tr(language, "rate_status_pending"))
        st.subheader(tr(language, "chart_supplier"))
        chart_df = df.groupby("proveedor", as_index=False)["monto_bs"].sum().sort_values("monto_bs", ascending=False).head(15)
        st.bar_chart(chart_df.set_index("proveedor"), height=260)
        with st.expander(tr(language, "details")):
            show = df.copy()
            show.columns = [PRETTY[c] for c in COLS]
            st.dataframe(show, use_container_width=True, hide_index=True)

# ---------------- TAB 4 ----------------
with export_actions:
    df = df_gastos()
    errores = []
    if issue_date_error:
        errores.append(tr(language, "invalid_issue_date"))
    if not responsable:
        errores.append(tr(language, "missing_responsible"))
    if not cedula:
        errores.append(tr(language, "missing_id"))
    if grant_error or not (monto_otorgado or 0):
        errores.append(tr(language, "missing_grant"))
    if rate_error or not (tasa or 0):
        errores.append(tr(language, "missing_rate"))
    if df.empty:
        errores.append(tr(language, "empty_expenses_export"))
    else:
        if st.session_state.get("monto_editor_invalid_rows"):
            errores.append(tr(language, "invalid_amounts_export"))
        for idx, valor in enumerate(df["fecha"], start=1):
            try:
                validar_fecha_gasto(valor, language)
            except ValueError:
                errores.append(tr(language, "invalid_expense_date", number=idx))
    current_signature = (
        df.to_json(), language, emp_nombre, emp_rif, emp_dir, emp_tel,
        responsable, cedula, str(monto_otorgado), str(tasa_decimal),
        str(fecha_emision), n_reporte, st.session_state.rate_source,
        st.session_state.get("bcv_value_date"), st.session_state.get("bcv_reference_rate"), st.session_state.bcv_last_error,
    )
    if st.session_state.get("last_file_signature") != current_signature:
        st.session_state.last_file = None
    if df.empty:
        st.info(tr(language, "empty_expenses_export"))
    else:
        total_bs, total_usd = totales(df, tasa)
        if not errores:
            st.caption(tr(language, "ready_summary", count=len(df), bs=format_number(total_bs, language, 2), usd=format_number(total_usd, language, 2), balance=format_number(float(monto_otorgado or 0) - total_usd, language, 2)))
        else:
            st.caption(f"{expense_count_caption(language, len(df))} · {format_number(total_bs, language, 2)} Bs")
    generate_requested = False
    if not st.session_state.last_file:
        generate_requested = st.button(tr(language, "generate_excel"), type="primary", disabled=df.empty)
    if generate_requested:
        if errores:
            st.warning("\n\n".join(errores))
        else:
            empresa = {"nombre": emp_nombre, "rif": emp_rif, "direccion": emp_dir, "telefonos": emp_tel}
            wb = build_caja_chica(
                df,
                empresa=empresa,
                responsable=responsable,
                cedula=cedula,
                fecha_emision=fecha_emision,
                monto_otorgado_usd=float(monto_otorgado),
                tasa_bcv=float(tasa_decimal),
                n_reporte=n_reporte,
                idioma=language,
                fecha_valor_bcv=(date.fromisoformat(st.session_state.bcv_value_date) if st.session_state.get("bcv_value_date") else None),
                origen_tasa=st.session_state.get("rate_source", "manual"),
                tasa_bcv_original=st.session_state.get("bcv_reference_rate"),
                bcv_actualizacion_fallida=st.session_state.get("bcv_last_error", False),
            )
            fname = f"caja_chica_{n_reporte}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
            out_path = os.path.join("caja_chica", fname)
            wb.save(out_path)
            buf = io.BytesIO()
            wb.save(buf)
            st.session_state.last_file = (fname, buf.getvalue(), out_path)
            st.session_state.last_file_signature = current_signature
            st.rerun()
    if st.session_state.last_file:
        fname, data, out_path = st.session_state.last_file
        st.download_button(
            tr(language, "download_excel"),
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(tr(language, "export_ready_hint"))
        with st.expander(tr(language, "excel_formula_help")):
            st.caption(tr(language, "excel_formula_note"))

current_expenses = df_gastos()
status_amount = "—" if st.session_state.get("monto_editor_invalid_rows") else format_number(float(current_expenses["monto_bs"].sum()) if not current_expenses.empty else 0, language, 2)
status_placeholder.markdown(workspace_summary(language, len(current_expenses), status_amount), unsafe_allow_html=True)

with st.container(key="session_footer"):
    st.caption(tr(language, "footer"))
