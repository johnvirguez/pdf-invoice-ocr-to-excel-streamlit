import io
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st

# =========================
# CONFIG UI
# =========================
st.set_page_config(page_title="SED | Facturas → Excel (Finanzas)", layout="wide")

st.markdown(
    """
    <style>
      /* Contenedor scrolleable para lista de archivos */
      .file-scroll {
        max-height: 220px;
        overflow-y: auto;
        padding: 8px 10px;
        border: 1px solid rgba(49, 51, 63, 0.2);
        border-radius: 10px;
        background: rgba(250, 250, 250, 0.5);
      }
      /* Tablas con scroll */
      .table-scroll {
        max-height: 520px;
        overflow-y: auto;
        border: 1px solid rgba(49, 51, 63, 0.2);
        border-radius: 10px;
        padding: 6px;
        background: white;
      }
      /* Reduce padding vertical en dataframes */
      div[data-testid="stDataFrame"] { margin-top: 0.25rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📄➡️📊 SED | Facturas PDF → Excel (Contabilidad/Finanzas)")
st.caption("Carga PDFs → Procesa → Selecciona una factura para ver detalle → Descarga Excel final.")

# =========================
# Helpers UI
# =========================
def df_to_display(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    present = [c for c in cols if c in df.columns]
    return df[present].copy() if present else df.copy()


# =========================
# EXTRACCIÓN (INTEGRAR TU BLOQUE ACTUAL AQUÍ)
# =========================
@st.cache_data(show_spinner=False)
def run_processing(files) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    ⚠️ AQUÍ debes pegar TU lógica actual de procesamiento:
    - leer PDFs
    - construir df_fin (FINANZAS_FACTURAS)
    - construir df_lines (LINEAS_FACTURA)
    - construir df_audit (AUDITORIA_TEXTO)
    y retornar (df_fin, df_lines, df_audit)

    Nota: st.cache_data acelera re-consultas del detalle (no reprocesa si no cambian los archivos)
    """
    # === INICIO: EJEMPLO MINIMO (BORRAR) ===
    df_fin = pd.DataFrame()
    df_lines = pd.DataFrame()
    df_audit = pd.DataFrame()
    # === FIN: EJEMPLO MINIMO (BORRAR) ===
    return df_fin, df_lines, df_audit


# =========================
# Upload + Control Panel
# =========================
with st.container():
    c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")

    with c1:
        uploaded_files = st.file_uploader(
            "Sube uno o varios PDFs (facturas)",
            type=["pdf"],
            accept_multiple_files=True,
            help="Tip: si el PDF es escaneado (imagen) y no tiene texto seleccionable, sin OCR no se podrá extraer detalle.",
        )

        # Lista scrolleable de archivos cargados
        if uploaded_files:
            st.markdown('<div class="file-scroll">', unsafe_allow_html=True)
            for f in uploaded_files:
                st.write(f"📄 {f.name} — {round(f.size/1024, 1)} KB")
            st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        show_audit = st.checkbox("Mostrar auditoría", value=False)
        show_raw = st.checkbox("Mostrar texto", value=False, help="Solo en el detalle de la factura seleccionada.")

    with c3:
        process = st.button("🚀 Procesar", type="primary", use_container_width=True)

st.divider()

# =========================
# Process
# =========================
if process:
    if not uploaded_files:
        st.error("Sube al menos un PDF.")
        st.stop()

    with st.spinner("Procesando facturas…"):
        df_fin, df_lines, df_audit = run_processing(uploaded_files)

    if df_fin is None or df_fin.empty:
        st.warning("No se generó información de encabezado (df_fin vacío). Revisa tu bloque de extracción.")
        st.stop()

    # Guardar en sesión para navegación sin recalcular
    st.session_state["df_fin"] = df_fin
    st.session_state["df_lines"] = df_lines
    st.session_state["df_audit"] = df_audit

    st.success(f"✅ Proceso completado: {len(df_fin)} factura(s).")

# =========================
# Reporte Unificado + Zoom Detalle
# =========================
if "df_fin" in st.session_state and not st.session_state["df_fin"].empty:
    df_fin = st.session_state["df_fin"]
    df_lines = st.session_state.get("df_lines", pd.DataFrame())
    df_audit = st.session_state.get("df_audit", pd.DataFrame())

    # Filtros superiores
    f1, f2, f3 = st.columns([1.5, 1, 1])
    with f1:
        search = st.text_input("🔎 Buscar (Proveedor, NIT, Factura, Documento)", value="")
    with f2:
        paises = sorted([p for p in df_fin.get("Pais", pd.Series(dtype=str)).dropna().unique().tolist() if str(p).strip()])
        pais_filter = st.multiselect("🌎 País", options=paises, default=paises)
    with f3:
        metodo_vals = sorted([m for m in df_fin.get("Metodo_Extraccion", pd.Series(dtype=str)).dropna().unique().tolist() if str(m).strip()])
        metodo_filter = st.multiselect("🧠 Método", options=metodo_vals, default=metodo_vals)

    # Aplicar filtros
    view = df_fin.copy()
    if pais_filter and "Pais" in view.columns:
        view = view[view["Pais"].isin(pais_filter)]
    if metodo_filter and "Metodo_Extraccion" in view.columns:
        view = view[view["Metodo_Extraccion"].isin(metodo_filter)]

    if search.strip():
        s = search.strip().lower()
        cols = [c for c in ["Proveedor_Razon_Social", "Proveedor_Id_Tributaria", "Factura_Numero", "Documento", "Cliente_Razon_Social"] if c in view.columns]
        if cols:
            mask = False
            for c in cols:
                mask = mask | view[c].astype(str).str.lower().str.contains(s, na=False)
            view = view[mask]

    # Layout: lista + detalle
    left, right = st.columns([1.1, 1.4], gap="large")

    # -------- LEFT: Tabla scrolleable (selección de factura)
    with left:
        st.subheader("📌 Facturas (Reporte)")
        st.caption("Selecciona una fila para ver el detalle a la derecha.")

        display_cols = [
            "Documento", "Pais", "Proveedor_Razon_Social", "Proveedor_Id_Tributaria",
            "Factura_Numero", "Fecha_Emision", "Total_Factura", "Moneda", "Metodo_Extraccion"
        ]
        view_disp = df_to_display(view, display_cols).reset_index(drop=True)

        # Selector compacto (listbox) para “zoom” sin depender de clic en tabla
        # (Streamlit no soporta click-row nativo estable; esto es más UX-friendly)
        options = []
        key_col = "Documento" if "Documento" in view_disp.columns else view_disp.columns[0]
        for i, r in view_disp.iterrows():
            factura = r.get("Factura_Numero", "")
            prov = (r.get("Proveedor_Razon_Social", "") or "")[:28]
            total = r.get("Total_Factura", "")
            options.append(f"{i+1:03d} | {factura} | {prov} | {total}")

        if not options:
            st.info("No hay resultados con los filtros actuales.")
            st.stop()

        selected = st.selectbox("📄 Selección", options=options, index=0)

        sel_idx = int(selected.split("|")[0].strip()) - 1
        selected_row = view_disp.iloc[sel_idx].to_dict()

        # Tabla resumen scrolleable (para listas largas)
        st.markdown('<div class="table-scroll">', unsafe_allow_html=True)
        st.dataframe(view_disp, use_container_width=True, height=520)
        st.markdown("</div>", unsafe_allow_html=True)

    # -------- RIGHT: Detalle “Zoom”
    with right:
        st.subheader("🔍 Detalle de la factura")
        st.caption("Encabezado + Líneas + Auditoría (opcional).")

        # Identificadores
        doc = selected_row.get("Documento", "")
        facnum = selected_row.get("Factura_Numero", "")

        tabs = st.tabs(["Encabezado", "Líneas", "Auditoría"])

        with tabs[0]:
            # Encabezado completo de esa factura
            header_df = df_fin[df_fin["Documento"] == doc].copy() if "Documento" in df_fin.columns else df_fin.copy()
            st.dataframe(header_df, use_container_width=True)

        with tabs[1]:
            if df_lines is None or df_lines.empty:
                st.info("No hay líneas detectadas.")
            else:
                # Preferir filtro por Documento y si existe Factura_Numero también
                lines_df = df_lines.copy()
                if "Documento" in lines_df.columns:
                    lines_df = lines_df[lines_df["Documento"] == doc]
                elif "Factura_Numero" in lines_df.columns and facnum:
                    lines_df = lines_df[lines_df["Factura_Numero"] == facnum]

                st.markdown('<div class="table-scroll">', unsafe_allow_html=True)
                st.dataframe(lines_df, use_container_width=True, height=520)
                st.markdown("</div>", unsafe_allow_html=True)

        with tabs[2]:
            if not show_audit:
                st.info("Activa 'Mostrar auditoría' en el panel superior para ver esta sección.")
            else:
                if df_audit is None or df_audit.empty:
                    st.info("No hay auditoría disponible.")
                else:
                    aud = df_audit[df_audit["Documento"] == doc].copy() if "Documento" in df_audit.columns else df_audit.copy()
                    st.dataframe(aud, use_container_width=True)

                    if show_raw and not aud.empty and "Texto" in aud.columns:
                        st.text_area("Texto extraído (preview)", value=str(aud.iloc[0]["Texto"]), height=320)

    st.divider()

    # Acción final: Descargar Excel (tu función actual)
    st.subheader("⬇️ Exportar")
    st.caption("Descarga el Excel final con formato contable (agrupado, sin cuadrícula, fuente Century Gothic 10).")

    # Aquí debes usar TU función actual build_excel_bytes(df_fin, df_lines, df_audit)
    # Para no romper tu lógica existente, dejo un placeholder:
    def build_excel_bytes_placeholder(df_fin, df_lines, df_audit) -> bytes:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_fin.to_excel(writer, index=False, sheet_name="FINANZAS_FACTURAS")
            df_lines.to_excel(writer, index=False, sheet_name="LINEAS_FACTURA")
            df_audit.to_excel(writer, index=False, sheet_name="AUDITORIA_TEXTO")
        return out.getvalue()

    excel_bytes = build_excel_bytes_placeholder(df_fin, df_lines, df_audit)
    filename = f"SED_Facturas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.download_button(
        "📥 Descargar Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
