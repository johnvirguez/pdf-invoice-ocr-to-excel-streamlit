import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader


# =========================
# App Config
# =========================
st.set_page_config(page_title="SED | Facturas → Excel (Finanzas)", layout="wide")
APP_TITLE = "📄➡️📊 SED | Facturas PDF → Excel (Contabilidad / Finanzas)"
APP_SUBTITLE = (
    "Extracción desde PDFs digitales (sin OCR). Formato fijo para Contabilidad/Finanzas. "
    "La columna 'Metodo_Extraccion' NO es error; indica qué parser se usó."
)

FIN_COLS = [
    "Documento",
    "Pais",
    "Tipo_Documento",
    "Proveedor_Razon_Social",
    "Proveedor_Id_Tributaria",
    "Cliente_Razon_Social",
    "Cliente_Id_Tributaria",
    "Factura_Numero",
    "Fecha_Emision",
    "Condicion_Venta",
    "Medio_Pago",
    "Moneda",
    "Simbolo_Moneda",
    "Subtotal",
    "Impuesto_IVA",
    "Total_Factura",
    "Anticipo",
    "Saldo",
    "Clave_Numerica",
    "Codigo_Unico_Consulta",
    "Probable_Escaneado",
    "Metodo_Extraccion",
    "Error",
]


# =========================
# Data Structures
# =========================
@dataclass
class FinanceInvoice:
    Documento: str
    Pais: str = ""
    Tipo_Documento: str = "Factura"
    Proveedor_Razon_Social: str = ""
    Proveedor_Id_Tributaria: str = ""
    Cliente_Razon_Social: str = ""
    Cliente_Id_Tributaria: str = ""
    Factura_Numero: str = ""
    Fecha_Emision: str = ""
    Condicion_Venta: str = ""
    Medio_Pago: str = ""
    Moneda: str = ""
    Simbolo_Moneda: str = ""
    Subtotal: Optional[float] = None
    Impuesto_IVA: Optional[float] = None
    Total_Factura: Optional[float] = None
    Anticipo: Optional[float] = None
    Saldo: Optional[float] = None
    Clave_Numerica: str = ""
    Codigo_Unico_Consulta: str = ""
    Probable_Escaneado: str = ""
    Metodo_Extraccion: str = ""
    Error: str = ""


# =========================
# Text Utilities
# =========================
def normalize_text(t: str) -> str:
    if t is None:
        return ""
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
    return len((text or "").strip()) < 50


def safe_group(m: Optional[re.Match], idx: int = 1) -> str:
    if not m:
        return ""
    try:
        val = m.group(idx) if m.lastindex and idx <= m.lastindex else m.group(0)
    except Exception:
        val = m.group(0) if m else ""
    return (val or "").strip()


def find_first(patterns: List[str], text: str, flags=re.IGNORECASE) -> str:
    if not text:
        return ""
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            if m.lastindex and m.lastindex >= 1:
                return safe_group(m, 1)
            return safe_group(m, 0)
    return ""


def parse_number_latam(s: str) -> Optional[float]:
    """
    - CO común: 1,541,384.13 (miles con coma, decimal con punto)
    - CR común: 1.541.384,13 o 1,541,384.13 (varía por generador)
    """
    if not s:
        return None
    raw = s.strip()
    raw = re.sub(r"[^\d,.\-]", "", raw)
    if not raw:
        return None

    last_comma = raw.rfind(",")
    last_dot = raw.rfind(".")

    if last_comma > last_dot:
        raw = raw.replace(".", "")
        raw = raw.replace(",", ".")
    else:
        raw = raw.replace(",", "")

    try:
        return float(raw)
    except:
        return None


def detect_currency(text: str) -> Tuple[str, str]:
    t = text or ""
    if " CRC" in t or "Moneda: CRC" in t or "Código Moneda........ CRC" in t:
        return "CRC", "¢"
    if " COP" in t or "pesos m/cte" in (t.lower()) or "Bogotá - Colombia" in t:
        return "COP", "$"
    # símbolos
    if "¢" in t:
        return "CRC", "¢"
    if "$" in t:
        return "COP", "$"
    return "", ""


# =========================
# Format Detectors
# =========================
def is_forlan_co(text: str) -> bool:
    t = (text or "").upper()
    return "FERRETERIA FORLAN" in t and "FACTURA ELECTR" in t and "TOTAL A PAGAR" in t


def is_navatec_cr(text: str) -> bool:
    t = (text or "").upper()
    return ("FACTURA ELECTRÓNICA N°" in t or "FACTURA ELECTRONICA N°" in t) and "FACTURAELECTRONICA.CR" in t


def is_gustavo_cr(text: str) -> bool:
    t = (text or "").upper()
    return "GUSTAVO GAMBOA VILLALOBOS" in t and "FACTURA ELECTRÓNICA #:" in t


def is_hacienda_tribu_cr(text: str) -> bool:
    t = (text or "").upper()
    return "WWW.HACIENDA.GO.CR" in t and "TRIBU-CR" in t and "COMPROBANTE" in t


# =========================
# Parsers (Finanzas)
# =========================
def parse_forlan_co_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first([r"(?m)^(FERRETERIA\s+FORLAN\s+SAS)\s*$"], text)
    proveedor_id = find_first([r"NIT\s*([0-9\.\-]+)"], text)

    cliente = find_first([r"Señores\s+([A-ZÁÉÍÓÚÑ0-9\.\s&\-]+)"], text)
    cliente_nit = find_first([r"Señores.*?\nNIT\s*([0-9\.\-]+)"], text)

    # "No. FO 16498" -> prefijo FO + número
    prefijo = find_first([r"No\.\s*([A-Z]{1,5})\s*\n*\s*([0-9]{3,})"], text)
    # El find_first no maneja 2 grupos; hacemos match directo:
    m = re.search(r"No\.\s*([A-Z]{1,5})\s*\n*\s*([0-9]{3,})", text, re.IGNORECASE)
    factura_num = f"{m.group(1)} {m.group(2)}".strip() if m else ""

    # Fecha: "Generación 16/02/2026, 14:53"
    fecha = find_first([r"Generaci[oó]n\s*([0-3]\d\/[01]\d\/[12]\d{3},\s*[0-2]\d:[0-5]\d)"], text)

    condicion = find_first([r"Forma\s+de\s+pago:\s*\n*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)
    medio = find_first([r"Medio\s+de\s+pago:\s*\n*([A-Za-zÁÉÍÓÚÑ\s\-]+)"], text)

    # Montos
    subtotal_str = find_first([r"Total\s+Bruto\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"IVA\s*19%\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total\s+a\s+Pagar\s*([0-9\.,]+)"], text)

    # OC / CUFE
    oc = find_first([r"Oc:\s*(OC[0-9]+)"], text)
    cufe = find_first([r"CUFE:\s*([a-f0-9]{20,})"], text)

    inv = FinanceInvoice(
        Documento=filename,
        Pais="CO",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social=cliente,
        Cliente_Id_Tributaria=cliente_nit,
        Factura_Numero=factura_num,
        Fecha_Emision=fecha,
        Condicion_Venta=condicion,
        Medio_Pago=medio,
        Moneda=moneda or "COP",
        Simbolo_Moneda=simbolo or "$",
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Anticipo=None,
        Saldo=None,
        Clave_Numerica=oc,                 # en CO usamos este campo para OC (para finanzas es útil)
        Codigo_Unico_Consulta=cufe,        # guardamos CUFE aquí
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="FORLAN CO (reglas finanzas)",
        Error="",
    )
    return inv


def parse_navatec_cr_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first([r"(?m)^(NAVATEC\s+INGENIERIA\s+S\.A\.)\s*$", r"(?m)^(NAVATECO)\s*$"], text)
    ids = re.findall(r"Ident\.\s*Jur[ií]dica:\s*([0-9\-]+)", text, flags=re.IGNORECASE)
    proveedor_id = (ids[0] if len(ids) >= 1 else "").strip()
    cliente_id = (ids[1] if len(ids) >= 2 else "").strip()

    cliente = find_first([r"Receptor\s+([A-ZÁÉÍÓÚÑ0-9\.\s&\-]+)"], text)

    factura_num = find_first([r"Factura\s+Electr[oó]nica\s+N°\s*([0-9]+)"], text)
    fecha = find_first([r"Fecha\s+de\s+Emisi[oó]n:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s+[0-2]\d:[0-5]\d\s*[ap]\.m\.)"], text)
    condicion = find_first([r"Condici[oó]n\s+de\s+venta:\s*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)
    medio = find_first([r"Medio\s+de\s+Pago:\s*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)

    clave = find_first([r"Clave\s+Num[eé]rica:\s*\n*([0-9]{30,})"], text)
    cod_unico = find_first([r"C[oó]digo\s+Único\s+de\s+Consulta:\s*([A-Z0-9]+)"], text)

    subtotal_str = find_first([r"Subtotal\s+Neto\s*¢\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"Total\s+Impuesto\s*¢\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total\s+Factura:\s*¢\s*([0-9\.,]+)"], text)
    anticipo_str = find_first([r"ANTICIPO\s*¢\s*([0-9\.,]+)"], text)
    saldo_str = find_first([r"SALDO\s*¢\s*([0-9\.,]+)"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais="CR",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social=cliente,
        Cliente_Id_Tributaria=cliente_id,
        Factura_Numero=factura_num,
        Fecha_Emision=fecha,
        Condicion_Venta=condicion,
        Medio_Pago=medio,
        Moneda=moneda or "CRC",
        Simbolo_Moneda=simbolo or "¢",
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Anticipo=parse_number_latam(anticipo_str),
        Saldo=parse_number_latam(saldo_str),
        Clave_Numerica=clave,
        Codigo_Unico_Consulta=cod_unico,
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="NAVATEC CR (reglas finanzas)",
        Error="",
    )


def parse_gustavo_cr_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first([r"Raz[oó]n\s+Social:\s*([^\n,]+)"], text)
    proveedor_id = find_first([r"C[eé]dula\s+F[ií]sica:\s*([0-9]+)"], text)

    cliente = find_first([r"Señor\(es\):\s*([A-ZÁÉÍÓÚÑ0-9\.\s&\-]+)"], text)
    cliente_id = find_first([r"Identificaci[oó]n:.*?\n*([0-9]{8,})"], text)

    factura_num = find_first([r"Factura\s+Electr[oó]nica\s*#:\s*([0-9]+)"], text)
    clave = find_first([r"Clave\s+Num[eé]rica\s*#:\s*([0-9]{30,})"], text)
    fecha = find_first([r"Fecha\s+y\s+Hora\s+de\s+Emisi[oó]n:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s+[0-2]\d:[0-5]\d:[0-5]\d\s*[AP]M)"], text)

    condicion = find_first([r"Condici[oó]n\s+de\s+Venta:\s*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)

    subtotal_str = find_first([r"Subtotal\s+Neto:\s*CRC\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"Total\s+Impuestos:\s*CRC\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total\s+Factura:\s*CRC\s*([0-9\.,]+)"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais="CR",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social=cliente,
        Cliente_Id_Tributaria=cliente_id,
        Factura_Numero=factura_num,
        Fecha_Emision=fecha,
        Condicion_Venta=condicion,
        Medio_Pago="",
        Moneda=moneda or "CRC",
        Simbolo_Moneda=simbolo or "¢",
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Anticipo=None,
        Saldo=None,
        Clave_Numerica=clave,
        Codigo_Unico_Consulta="",
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="GUSTAVO CR (reglas finanzas)",
        Error="",
    )


def parse_hacienda_tribu_cr_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first([r"Nombre:\s*([A-ZÁÉÍÓÚÑ\s]+)\nNombre comercial:"], text)
    proveedor_id = find_first([r"C[eé]dula:\s*([0-9]+)"], text)

    cliente = find_first([r"DATOS\s+DEL\s+CLIENTE\s+Nombre:\s*([A-ZÁÉÍÓÚÑ\s]+)"], text)
    cliente_id = find_first([r"DATOS\s+DEL\s+CLIENTE.*?C[eé]dula:\s*([0-9]+)"], text)

    consecutivo = find_first([r"Consecutivo:\s*([0-9]+)"], text)
    clave = find_first([r"Clave:\s*([0-9]{30,})"], text)
    fecha = find_first([r"Fecha:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s+[0-2]\d:[0-5]\d:[0-5]\d)"], text)

    condicion = find_first([r"Condici[oó]n\s+de\s+Venta:\s*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)
    medio = find_first([r"Medio\s+de\s+Pago:\s*([A-Za-zÁÉÍÓÚÑ\s\-]+)"], text)

    # Totales (en página 2 suele estar)
    subtotal_str = find_first([r"Total\s+venta\s+neta\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"Total\s+impuestos\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total\s+comprobante\s*([0-9\.,]+)"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais="CR",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social=cliente,
        Cliente_Id_Tributaria=cliente_id,
        Factura_Numero=consecutivo,
        Fecha_Emision=fecha,
        Condicion_Venta=condicion,
        Medio_Pago=medio,
        Moneda=moneda or "CRC",
        Simbolo_Moneda=simbolo or "¢",
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Anticipo=None,
        Saldo=None,
        Clave_Numerica=clave,
        Codigo_Unico_Consulta="",
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="HACIENDA TRIBU-CR (reglas finanzas)",
        Error="",
    )


def parse_generic_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first(
        [
            r"Raz[oó]n\s+Social[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
            r"Proveedor[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
            r"Emisor[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
        ],
        text,
    ).split("\n")[0].strip()

    proveedor_id = find_first(
        [
            r"NIT[:\s]*([0-9\.\-]{6,20})",
            r"N\.I\.T\.[:\s]*([0-9\.\-]{6,20})",
            r"Ident\.\s*Jur[ií]dica:\s*([0-9\-]+)",
            r"C[eé]dula:\s*([0-9]+)",
        ],
        text,
    )

    factura_num = find_first(
        [
            r"Factura\s*(?:No\.|Nro\.|N°|#)?\s*[:\s]*([A-Z0-9\-]{3,})",
            r"Invoice\s*(?:No\.|Number)?\s*[:\s]*([A-Z0-9\-]{3,})",
            r"Consecutivo:\s*([0-9]+)",
        ],
        text,
    )

    fecha = find_first(
        [
            r"Fecha\s*(?:de\s*Emisi[oó]n)?[:\s]*([0-3]?\d[\/\-][01]?\d[\/\-][12]\d{3})",
            r"Fecha:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s+[0-2]\d:[0-5]\d:[0-5]\d)",
        ],
        text,
    )

    subtotal_str = find_first([r"Subtotal[:\s\$]*([0-9\.\,]+)", r"Total\s+Bruto\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"IVA[:\s\$]*([0-9\.\,]+)", r"Total\s+impuestos\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total[:\s\$]*([0-9\.\,]+)", r"Total\s+a\s+Pagar\s*([0-9\.,]+)"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais="",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social="",
        Cliente_Id_Tributaria="",
        Factura_Numero=factura_num,
        Fecha_Emision=fecha,
        Condicion_Venta="",
        Medio_Pago="",
        Moneda=moneda,
        Simbolo_Moneda=simbolo,
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Anticipo=None,
        Saldo=None,
        Clave_Numerica="",
        Codigo_Unico_Consulta="",
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="Genérico (heurístico finanzas)",
        Error="",
    )


# =========================
# Excel Builder
# =========================
def build_excel_bytes(df_fin: pd.DataFrame, df_audit: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_fin.to_excel(writer, index=False, sheet_name="FINANZAS_FACTURAS")
        df_audit.to_excel(writer, index=False, sheet_name="AUDITORIA_TEXTO")
    return out.getvalue()


# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

with st.sidebar:
    st.header("⚙️ Opciones")
    show_audit = st.checkbox("Incluir hoja de auditoría (texto extraído)", value=True)
    max_audit_chars = st.slider("Máximo texto por documento (auditoría)", 2000, 32000, 12000, 500)

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

    fin_rows: List[Dict[str, Any]] = []
    audit_rows: List[Dict[str, Any]] = []

    prog = st.progress(0)
    status = st.empty()
    total_files = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        status.write(f"Procesando **{uf.name}** ({idx}/{total_files})…")
        pdf_bytes = uf.read()

        try:
            text = extract_text_pypdf(pdf_bytes)

            if is_forlan_co(text):
                inv = parse_forlan_co_finance(text, uf.name)
            elif is_navatec_cr(text):
                inv = parse_navatec_cr_finance(text, uf.name)
            elif is_gustavo_cr(text):
                inv = parse_gustavo_cr_finance(text, uf.name)
            elif is_hacienda_tribu_cr(text):
                inv = parse_hacienda_tribu_cr_finance(text, uf.name)
            else:
                inv = parse_generic_finance(text, uf.name)

            row = inv.__dict__
            fin_rows.append({c: row.get(c, "") for c in FIN_COLS})

            if show_audit:
                audit_rows.append(
                    {
                        "Documento": uf.name,
                        "Longitud_Texto": len(text),
                        "Texto": (text or "")[:max_audit_chars],
                    }
                )

        except Exception as e:
            err_row = {c: "" for c in FIN_COLS}
            err_row["Documento"] = uf.name
            err_row["Metodo_Extraccion"] = "ERROR"
            err_row["Error"] = str(e)
            fin_rows.append(err_row)

            if show_audit:
                audit_rows.append({"Documento": uf.name, "Longitud_Texto": 0, "Texto": f"ERROR: {e}"})

        prog.progress(int((idx / total_files) * 100))

    status.success("✅ Proceso finalizado. Generando Excel…")

    df_fin = pd.DataFrame(fin_rows, columns=FIN_COLS)
    df_audit = pd.DataFrame(audit_rows) if show_audit else pd.DataFrame()

    st.subheader("Vista previa (FINANZAS_FACTURAS)")
    st.dataframe(df_fin, use_container_width=True)

    scanned_count = (df_fin["Probable_Escaneado"] == "SI").sum()
    if scanned_count:
        st.warning(
            f"Detecté **{scanned_count}** PDF(s) con muy poco texto (probablemente escaneados). "
            "Sin OCR, esos archivos no se capturan bien."
        )

    excel_bytes = build_excel_bytes(df_fin, df_audit) if show_audit else build_excel_bytes(df_fin, pd.DataFrame())
    filename = f"SED_Facturas_Finanzas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.download_button(
        "⬇️ Descargar Excel (Finanzas)",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
