import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader


# =========================
# Configuración
# =========================
st.set_page_config(page_title="SED | Facturas PDF → Excel", layout="wide")

APP_TITLE = "📄➡️📊 SED | Facturas PDF → Excel (sin Azure)"
APP_SUBTITLE = "Extrae texto de PDFs digitales y detecta campos típicos de facturas para exportar a Excel."


# =========================
# Utilidades de texto
# =========================
def normalize_text(t: str) -> str:
    t = t.replace("\u00a0", " ")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def extract_text_pypdf(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    parts: List[str] = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return normalize_text("\n".join(parts))


def looks_scanned(text: str) -> bool:
    # Heurística simple: muy poco texto => probablemente escaneado (imagen)
    return len(text.strip()) < 50


def parse_number_co(s: str) -> Optional[float]:
    """
    Intenta interpretar números estilo CO:
    - 1.234.567,89
    - 1,234,567.89
    - 1234567.89
    Retorna float o None.
    """
    if not s:
        return None
    raw = s.strip()
    raw = re.sub(r"[^\d,.\-]", "", raw)
    if not raw:
        return None

    # Si tiene coma y punto, definimos separador decimal por el último símbolo
    last_comma = raw.rfind(",")
    last_dot = raw.rfind(".")

    if last_comma > last_dot:
        # coma decimal, puntos miles
        raw = raw.replace(".", "")
        raw = raw.replace(",", ".")
    else:
        # punto decimal, comas miles
        raw = raw.replace(",", "")

    try:
        return float(raw)
    except:
        return None


def find_first(patterns: List[str], text: str, flags=re.IGNORECASE) -> Optional[str]:
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            # si hay grupos, usa el 1; si no, todo
            return m.group(1).strip() if m.groups() else m.group(0).strip()
    return None


@dataclass
class InvoiceExtract:
    documento: str
    es_probable_escaneado: bool
    proveedor: Optional[str]
    nit: Optional[str]
    factura_numero: Optional[str]
    fecha: Optional[str]
    subtotal: Optional[float]
    iva: Optional[float]
    total: Optional[float]
    moneda: Optional[str]
    confidence_hint: str  # solo para indicar "heurístico"


def detect_invoice_fields(text: str, filename: str) -> InvoiceExtract:
    """
    Detección heurística (sin OCR) basada en texto extraído.
    Ajustable a formatos de factura en Colombia.
    """
    scanned = looks_scanned(text)

    # Proveedor (muy heurístico): intenta capturar una línea cercana a "Señores", "Proveedor", "Razón Social"
    proveedor = find_first(
        [
            r"Raz[oó]n\s+Social[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
            r"Proveedor[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
            r"Emisor[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
        ],
        text,
    )
    if proveedor:
        # corta a 1 línea
        proveedor = proveedor.split("\n")[0].strip()

    # NIT / Tax ID
    nit = find_first(
        [
            r"NIT[:\s]*([0-9\.\-]{6,20})",
            r"Tax\s*Id[:\s]*([0-9\.\-]{6,20})",
            r"N\.I\.T\.[:\s]*([0-9\.\-]{6,20})",
        ],
        text,
    )

    # Número de factura
    factura_num = find_first(
        [
            r"Factura\s*(No\.|Nro\.|N°|#)?\s*[:\s]*([A-Z0-9\-]{3,})",
            r"No\.\s*Factura[:\s]*([A-Z0-9\-]{3,})",
            r"Invoice\s*(No\.|Number)?\s*[:\s]*([A-Z0-9\-]{3,})",
        ],
        text,
    )
    # Algunas regex tienen 2 grupos; si capturó el grupo 1 del find_first, puede quedar mal.
    # Reintento específico por grupos:
    if factura_num and re.search(r"Factura\s*(No\.|Nro\.|N°|#)?", factura_num, re.IGNORECASE):
        m = re.search(r"Factura\s*(?:No\.|Nro\.|N°|#)?\s*[:\s]*([A-Z0-9\-]{3,})", factura_num, re.IGNORECASE)
        if m:
            factura_num = m.group(1)

    # Fecha (dd/mm/yyyy o dd-mm-yyyy)
    fecha = find_first(
        [
            r"Fecha\s*(de\s*Emisi[oó]n)?[:\s]*([0-3]?\d[\/\-][01]?\d[\/\-][12]\d{3})",
            r"Fecha[:\s]*([0-3]?\d[\/\-][01]?\d[\/\-][12]\d{3})",
        ],
        text,
    )
    # Reintento por 2 grupos
    if fecha and "Fecha" in fecha:
        m = re.search(r"Fecha\s*(?:de\s*Emisi[oó]n)?[:\s]*([0-3]?\d[\/\-][01]?\d[\/\-][12]\d{3})", text, re.IGNORECASE)
        if m:
            fecha = m.group(1)

    # Moneda (muy básico)
    moneda = find_first(
        [
            r"Moneda[:\s]*([A-Z]{3})",
            r"(COP|USD|EUR)\b",
        ],
        text,
    )

    # Totales: buscamos valores cercanos a palabras clave
    # Subtotal
    subtotal_str = find_first(
        [
            r"Subtotal[:\s\$]*([0-9\.\,]+)",
            r"Sub\s*total[:\s\$]*([0-9\.\,]+)",
        ],
        text,
    )
    subtotal = parse_number_co(subtotal_str) if subtotal_str else None

    # IVA / Impuesto
    iva_str = find_first(
        [
            r"IVA[:\s\$]*([0-9\.\,]+)",
            r"Impuesto[s]?\s*(?:IVA)?[:\s\$]*([0-9\.\,]+)",
        ],
        text,
    )
    iva = parse_number_co(iva_str) if iva_str else None

    # Total
    total_str = find_first(
        [
            r"Total\s*(?:a\s*pagar)?[:\s\$]*([0-9\.\,]+)",
            r"TOTAL[:\s\$]*([0-9\.\,]+)",
            r"Valor\s*Total[:\s\$]*([0-9\.\,]+)",
        ],
        text,
    )
    total = parse_number_co(total_str) if total_str else None

    return InvoiceExtract(
        documento=filename,
        es_probable_escaneado=scanned,
        proveedor=proveedor,
        nit=nit,
        factura_numero=factura_num,
        fecha=fecha,
        subtotal=subtotal,
        iva=iva,
        total=total,
        moneda=moneda,
        confidence_hint="heurístico (sin OCR)",
    )


def build_excel_bytes(resumen: pd.DataFrame, raw: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
        raw.to_excel(writer, index=False, sheet_name="Texto_Extraido")
    return out.getvalue()


# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

with st.sidebar:
    st.header("⚙️ Opciones")
    show_raw_preview = st.checkbox("Mostrar preview del texto extraído", value=True)
    max_preview_chars = st.slider("Preview (caracteres)", 200, 5000, 1200, 100)

st.divider()

uploaded_files = st.file_uploader(
    "Sube uno o varios PDFs (facturas)",
    type=["pdf"],
    accept_multiple_files=True,
)

process = st.button("🚀 Procesar PDFs", type="primary", use_container_width=True)

if process:
    if not uploaded_files:
        st.error("Sube al menos un PDF.")
        st.stop()

    resumen_rows: List[Dict] = []
    raw_rows: List[Dict] = []

    prog = st.progress(0)
    status = st.empty()

    total_files = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        status.write(f"Procesando **{uf.name}** ({idx}/{total_files})…")
        pdf_bytes = uf.read()

        try:
            text = extract_text_pypdf(pdf_bytes)
            inv = detect_invoice_fields(text, uf.name)

            resumen_rows.append({
                "Documento": inv.documento,
                "Probable_Escaneado": "SI" if inv.es_probable_escaneado else "NO",
                "Proveedor": inv.proveedor or "",
                "NIT": inv.nit or "",
                "Factura_Numero": inv.factura_numero or "",
                "Fecha": inv.fecha or "",
                "Subtotal": inv.subtotal,
                "IVA": inv.iva,
                "Total": inv.total,
                "Moneda": inv.moneda or "",
                "Metodo": inv.confidence_hint,
            })

            raw_rows.append({
                "Documento": uf.name,
                "Longitud_Texto": len(text),
                "Texto": text[:32000],  # evita celdas gigantes
            })

        except Exception as e:
            resumen_rows.append({
                "Documento": uf.name,
                "Probable_Escaneado": "",
                "Proveedor": "",
                "NIT": "",
                "Factura_Numero": "",
                "Fecha": "",
                "Subtotal": None,
                "IVA": None,
                "Total": None,
                "Moneda": "",
                "Metodo": "ERROR",
                "Error": str(e),
            })
            raw_rows.append({
                "Documento": uf.name,
                "Longitud_Texto": 0,
                "Texto": f"ERROR: {e}",
            })

        prog.progress(int((idx / total_files) * 100))

    status.success("✅ Proceso finalizado. Generando Excel…")

    df_resumen = pd.DataFrame(resumen_rows)
    df_raw = pd.DataFrame(raw_rows)

    # Vista en pantalla
    st.subheader("Resumen detectado")
    st.dataframe(df_resumen, use_container_width=True)

    # Alertas: escaneados
    scanned_count = (df_resumen["Probable_Escaneado"] == "SI").sum() if "Probable_Escaneado" in df_resumen.columns else 0
    if scanned_count:
        st.warning(
            f"Detecté **{scanned_count}** PDF(s) con muy poco texto (probablemente escaneados). "
            "Sin OCR, esos archivos no podrán extraer campos con precisión."
        )

    if show_raw_preview:
        st.subheader("Texto extraído (preview)")
        for r in raw_rows:
            with st.expander(f"{r['Documento']} — {r['Longitud_Texto']} chars"):
                st.text(r["Texto"][:max_preview_chars])

    excel_bytes = build_excel_bytes(df_resumen, df_raw)
    filename = f"SED_facturas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.download_button(
        "⬇️ Descargar Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
