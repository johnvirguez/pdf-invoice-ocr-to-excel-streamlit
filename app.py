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
st.set_page_config(page_title="SED | Facturas → Excel (Finanzas CO + Líneas)", layout="wide")
APP_TITLE = "📄➡️📊 SED | Facturas PDF → Excel (Contabilidad Colombia + Líneas)"
APP_SUBTITLE = (
    "Extracción desde PDFs digitales (sin OCR). "
    "Genera encabezado (totales) + todas las líneas de factura. "
    "Incluye columnas para marcar costo."
)

# Encabezado (1 fila por factura) - Contabilidad CO + compatibilidad CR
FIN_COLS = [
    "Documento",
    "Pais",
    "Tipo_Documento",
    "Proveedor_Razon_Social",
    "Proveedor_Id_Tributaria",
    "Cliente_Razon_Social",
    "Cliente_Id_Tributaria",
    "Prefijo",
    "Factura_Numero",
    "Consecutivo",
    "Fecha_Emision",
    "Condicion_Venta",
    "Forma_Pago",
    "Medio_Pago",
    "Moneda",
    "Simbolo_Moneda",
    "Subtotal",
    "Impuesto_IVA",
    "Total_Factura",
    "OC",
    "CUFE",
    "Resolucion_DIAN",
    "QR_o_Codigo",
    "Clave_Numerica",
    "Codigo_Unico_Consulta",
    "Anticipo",
    "Saldo",
    # NUEVO: marcar costo a nivel factura
    "Costo_Factura_Marcado",
    "Probable_Escaneado",
    "Metodo_Extraccion",
    "Error",
]

# Líneas (múltiples filas por factura)
LINE_COLS = [
    "Documento",
    "Linea",
    "Codigo_Item",
    "Descripcion",
    "Cantidad",
    "Unidad",
    "Precio_Unitario",
    "Descuento",
    "Subtotal_Linea",
    "Impuesto_Linea",
    "Total_Linea",
    "Moneda",
    "Pais",
    # NUEVO: columnas para marcar costo por línea
    "Marca_Costo",
    "Cuenta_Costo",
    "Descripcion_Raw",  # respaldo por si falla el parsing fino
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
    Prefijo: str = ""
    Factura_Numero: str = ""
    Consecutivo: str = ""
    Fecha_Emision: str = ""
    Condicion_Venta: str = ""
    Forma_Pago: str = ""
    Medio_Pago: str = ""
    Moneda: str = ""
    Simbolo_Moneda: str = ""
    Subtotal: Optional[float] = None
    Impuesto_IVA: Optional[float] = None
    Total_Factura: Optional[float] = None
    OC: str = ""
    CUFE: str = ""
    Resolucion_DIAN: str = ""
    QR_o_Codigo: str = ""
    Clave_Numerica: str = ""
    Codigo_Unico_Consulta: str = ""
    Anticipo: Optional[float] = None
    Saldo: Optional[float] = None
    Costo_Factura_Marcado: str = ""  # para diligenciar después
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
    if not s:
        return None
    raw = s.strip()
    raw = re.sub(r"[^\d,.\-]", "", raw)
    if not raw:
        return None

    last_comma = raw.rfind(",")
    last_dot = raw.rfind(".")

    # decimal separator = the last of comma/dot
    if last_comma > last_dot:
        # comma decimal, dots thousands
        raw = raw.replace(".", "")
        raw = raw.replace(",", ".")
    else:
        # dot decimal, commas thousands
        raw = raw.replace(",", "")

    try:
        return float(raw)
    except:
        return None


def detect_currency(text: str) -> Tuple[str, str]:
    t = text or ""
    if " CRC" in t or "Moneda: CRC" in t or "Código Moneda........ CRC" in t or "¢" in t:
        return "CRC", "¢"
    if " COP" in t or "pesos" in t.lower() or "Bogotá - Colombia" in t or "$" in t:
        return "COP", "$"
    return "", ""


def consolidate_wrapped_lines(lines: List[str]) -> List[str]:
    """
    Une líneas partidas (muy común en PDFs) para mejorar parsing de ítems.
    Heurística:
      - Si una línea NO empieza con código de ítem, y la anterior parece parte de ítem,
        se concatena.
    """
    out: List[str] = []
    for line in lines:
        s = (line or "").strip()
        if not s:
            continue

        starts_item = bool(re.match(r"^\d{3}\s+\d", s))  # NAVATEC: 001 1.00 ...
        # Para Colombia, a veces el ítem no inicia con 001; dejamos fallback abajo.
        if out and not starts_item:
            # Si la anterior parece ítem (tiene números al final), une
            prev = out[-1]
            if re.search(r"([0-9][0-9\.,]+)\s*$", prev) or re.match(r"^\d{3}\s+\d", prev.strip()):
                out[-1] = prev + " " + s
            else:
                out.append(s)
        else:
            out.append(s)
    return out


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
    return "GUSTAVO GAMBOA VILLALOBOS" in t and ("RESUMEN DEL DOCUMENTO" in t or "FACTURA ELECTRÓNICA #:" in t)


def is_hacienda_tribu_cr(text: str) -> bool:
    t = (text or "").upper()
    return "WWW.HACIENDA.GO.CR" in t and "TRIBU-CR" in t and "COMPROBANTE" in t


# =========================
# Line Items Extractors
# =========================
def extract_items_navatec(text: str, doc: str, pais: str, moneda: str) -> List[Dict[str, Any]]:
    """
    NAVATEC / facturaelectronica.cr: líneas estilo
      001 1.00 Unid MTR001 EL COCO ALAJUELA 741,300.84 0.00 741.300,84 96,369.11
    Captura:
      - Linea (001)
      - Cantidad (1.00)
      - Unidad (Unid)
      - Codigo_Item (MTR001)
      - Descripcion (...)
      - Precio_Unitario
      - Descuento
      - Subtotal_Linea
      - Impuesto_Linea
    """
    raw_lines = consolidate_wrapped_lines((text or "").splitlines())
    items: List[Dict[str, Any]] = []

    pat = re.compile(
        r"^(?P<linea>\d{3})\s+"
        r"(?P<cantidad>\d+(?:\.\d+)?)\s+"
        r"(?P<unidad>\w+)\s+"
        r"(?P<codigo>[A-Z0-9]+)\s+"
        r"(?P<desc>.+?)\s+"
        r"(?P<precio>[0-9\.,]+)\s+"
        r"(?P<descuento>[0-9\.,]+)\s+"
        r"(?P<subtotal>[0-9\.,]+)\s+"
        r"(?P<imp>[0-9\.,]+)\s*$"
    )

    for ln in raw_lines:
        m = pat.match(ln)
        if not m:
            continue

        cantidad = parse_number_latam(m.group("cantidad"))
        precio = parse_number_latam(m.group("precio"))
        descuento = parse_number_latam(m.group("descuento"))
        subtotal_linea = parse_number_latam(m.group("subtotal"))
        impuesto_linea = parse_number_latam(m.group("imp"))
        total_linea = None
        if subtotal_linea is not None and impuesto_linea is not None:
            total_linea = subtotal_linea + impuesto_linea

        items.append(
            {
                "Documento": doc,
                "Linea": m.group("linea"),
                "Codigo_Item": m.group("codigo"),
                "Descripcion": m.group("desc").strip(),
                "Cantidad": cantidad,
                "Unidad": m.group("unidad"),
                "Precio_Unitario": precio,
                "Descuento": descuento,
                "Subtotal_Linea": subtotal_linea,
                "Impuesto_Linea": impuesto_linea,
                "Total_Linea": total_linea,
                "Moneda": moneda,
                "Pais": pais,
                "Marca_Costo": "",
                "Cuenta_Costo": "",
                "Descripcion_Raw": ln,
            }
        )

    return items


def extract_items_generic(text: str, doc: str, pais: str, moneda: str) -> List[Dict[str, Any]]:
    """
    Extractor genérico (best-effort) para PDFs CO u otros donde no haya formato fijo.
    Objetivo: no perder info.
    Heurística: líneas que terminen con 1 o más valores numéricos.
    Si detecta 2 valores al final, asume (Subtotal_Linea, Impuesto_Linea) o (Precio, Total).
    """
    raw_lines = consolidate_wrapped_lines((text or "").splitlines())
    items: List[Dict[str, Any]] = []

    # Captura hasta 3 números al final de la línea
    tail_nums = re.compile(r"^(?P<body>.*?)(?P<n1>[0-9][0-9\.,]+)\s+(?P<n2>[0-9][0-9\.,]+)(?:\s+(?P<n3>[0-9][0-9\.,]+))?\s*$")

    line_no = 0
    for ln in raw_lines:
        s = (ln or "").strip()

        # Filtra encabezados comunes
        if len(s) < 8:
            continue
        if any(k in s.upper() for k in ["SUBTOTAL", "TOTAL", "IVA", "CUFE", "RESOLU", "CLAVE", "PÁGINA", "PAGINA", "AUTORIZADO"]):
            continue

        m = tail_nums.match(s)
        if not m:
            continue

        line_no += 1
        body = m.group("body").strip()
        n1 = parse_number_latam(m.group("n1"))
        n2 = parse_number_latam(m.group("n2"))
        n3 = parse_number_latam(m.group("n3")) if m.group("n3") else None

        # Asignación tentativa:
        # Si hay 3 números: precio, subtotal, impuesto (muy tentativo)
        precio = n1 if n3 is not None else None
        subtotal_linea = n2 if n3 is not None else n1
        impuesto_linea = n3 if n3 is not None else None
        total_linea = None
        if subtotal_linea is not None and impuesto_linea is not None:
            total_linea = subtotal_linea + impuesto_linea
        elif n2 is not None and n1 is not None and total_linea is None:
            # fallback: el último podría ser total
            total_linea = n2

        items.append(
            {
                "Documento": doc,
                "Linea": str(line_no),
                "Codigo_Item": "",
                "Descripcion": body[:250],
                "Cantidad": None,
                "Unidad": "",
                "Precio_Unitario": precio,
                "Descuento": None,
                "Subtotal_Linea": subtotal_linea,
                "Impuesto_Linea": impuesto_linea,
                "Total_Linea": total_linea,
                "Moneda": moneda,
                "Pais": pais,
                "Marca_Costo": "",
                "Cuenta_Costo": "",
                "Descripcion_Raw": s,
            }
        )

    return items


# =========================
# Header Parsers (Finanzas)
# =========================
def parse_forlan_co_finance(text: str, filename: str) -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first([r"(?m)^(FERRETERIA\s+FORLAN\s+SAS)\s*$"], text)
    proveedor_id = find_first([r"NIT\s*([0-9\.\-]+)"], text)

    cliente = find_first([r"Señores\s+([A-ZÁÉÍÓÚÑ0-9\.\s&\-]+)"], text)
    cliente_nit = find_first([r"Señores.*?\nNIT\s*([0-9\.\-]+)"], text)

    m = re.search(r"No\.\s*([A-Z]{1,5})\s*\n*\s*([0-9]{3,})", text, re.IGNORECASE)
    prefijo = (m.group(1) or "").strip() if m else ""
    consecutivo = (m.group(2) or "").strip() if m else ""
    factura_num = f"{prefijo} {consecutivo}".strip() if prefijo or consecutivo else ""

    fecha = find_first([r"Generaci[oó]n\s*([0-3]\d\/[01]\d\/[12]\d{3},\s*[0-2]\d:[0-5]\d)"], text)

    forma_pago = find_first([r"Forma\s+de\s+pago:\s*\n*([A-Za-zÁÉÍÓÚÑ\s]+)"], text)
    medio_pago = find_first([r"Medio\s+de\s+pago:\s*\n*([A-Za-zÁÉÍÓÚÑ\s\-]+)"], text)

    subtotal_str = find_first([r"Total\s+Bruto\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"IVA\s*19%\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total\s+a\s+Pagar\s*([0-9\.,]+)"], text)

    oc = find_first([r"Oc:\s*(OC[0-9]+)"], text)
    cufe = find_first([r"CUFE:\s*([a-f0-9]{20,})"], text)
    resol = find_first([r"Resoluci[oó]n\s*(?:DIAN)?\s*[:\-]?\s*([0-9\-\/]+)"], text)
    qr_hint = find_first([r"(CUFE:\s*[a-f0-9]{20,})"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais="CO",
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Cliente_Razon_Social=cliente,
        Cliente_Id_Tributaria=cliente_nit,
        Prefijo=prefijo,
        Factura_Numero=factura_num,
        Consecutivo=consecutivo,
        Fecha_Emision=fecha,
        Condicion_Venta="",
        Forma_Pago=forma_pago,
        Medio_Pago=medio_pago,
        Moneda=moneda or "COP",
        Simbolo_Moneda=simbolo or "$",
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        OC=oc,
        CUFE=cufe,
        Resolucion_DIAN=resol,
        QR_o_Codigo=qr_hint,
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="FORLAN CO (contabilidad + líneas)",
        Error="",
    )


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
        Metodo_Extraccion="NAVATEC CR (contabilidad + líneas)",
        Error="",
    )


def parse_generic_finance(text: str, filename: str, pais_hint: str = "") -> FinanceInvoice:
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
            r"Generaci[oó]n\s*([0-3]\d\/[01]\d\/[12]\d{3},\s*[0-2]\d:[0-5]\d)",
        ],
        text,
    )

    subtotal_str = find_first([r"Subtotal[:\s\$]*([0-9\.\,]+)", r"Total\s+Bruto\s*([0-9\.,]+)"], text)
    iva_str = find_first([r"IVA[:\s\$]*([0-9\.\,]+)", r"Total\s+impuestos\s*([0-9\.,]+)"], text)
    total_str = find_first([r"Total[:\s\$]*([0-9\.\,]+)", r"Total\s+a\s+Pagar\s*([0-9\.,]+)"], text)

    return FinanceInvoice(
        Documento=filename,
        Pais=pais_hint,
        Tipo_Documento="Factura",
        Proveedor_Razon_Social=proveedor,
        Proveedor_Id_Tributaria=proveedor_id,
        Factura_Numero=factura_num,
        Fecha_Emision=fecha,
        Moneda=moneda,
        Simbolo_Moneda=simbolo,
        Subtotal=parse_number_latam(subtotal_str),
        Impuesto_IVA=parse_number_latam(iva_str),
        Total_Factura=parse_number_latam(total_str),
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="Genérico (contabilidad + líneas)",
        Error="",
    )


# =========================
# Excel Builder
# =========================
def build_excel_bytes(df_fin: pd.DataFrame, df_lines: pd.DataFrame, df_audit: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_fin.to_excel(writer, index=False, sheet_name="FINANZAS_FACTURAS")
        df_lines.to_excel(writer, index=False, sheet_name="LINEAS_FACTURA")
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
    line_rows: List[Dict[str, Any]] = []
    audit_rows: List[Dict[str, Any]] = []

    prog = st.progress(0)
    status = st.empty()
    total_files = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        status.write(f"Procesando **{uf.name}** ({idx}/{total_files})…")
        pdf_bytes = uf.read()

        try:
            text = extract_text_pypdf(pdf_bytes)
            moneda, _ = detect_currency(text)

            # Encabezado + Líneas según tipo
            if is_forlan_co(text):
                inv = parse_forlan_co_finance(text, uf.name)
                items = extract_items_generic(text, uf.name, inv.Pais, inv.Moneda)  # Forlan: best-effort
            elif is_navatec_cr(text):
                inv = parse_navatec_cr_finance(text, uf.name)
                items = extract_items_navatec(text, uf.name, inv.Pais, inv.Moneda)  # NAVATEC: sólido
            elif is_gustavo_cr(text) or is_hacienda_tribu_cr(text):
                # Para estos CR, por ahora líneas genéricas (si aparecen)
                inv = parse_generic_finance(text, uf.name, pais_hint="CR")
                inv.Metodo_Extraccion = "CR (genérico) contabilidad + líneas"
                items = extract_items_generic(text, uf.name, inv.Pais or "CR", inv.Moneda or (moneda or "CRC"))
            else:
                inv = parse_generic_finance(text, uf.name)
                items = extract_items_generic(text, uf.name, inv.Pais, inv.Moneda)

            # Garantiza columnas fijas del encabezado
            row = inv.__dict__
            fin_rows.append({c: row.get(c, "") for c in FIN_COLS})

            # Líneas
            if items:
                for it in items:
                    line_rows.append({c: it.get(c, "") for c in LINE_COLS})
            else:
                # si no detecta líneas, deja una fila "vacía" para no perder trazabilidad
                line_rows.append(
                    {
                        "Documento": uf.name,
                        "Linea": "",
                        "Codigo_Item": "",
                        "Descripcion": "",
                        "Cantidad": None,
                        "Unidad": "",
                        "Precio_Unitario": None,
                        "Descuento": None,
                        "Subtotal_Linea": None,
                        "Impuesto_Linea": None,
                        "Total_Linea": None,
                        "Moneda": inv.Moneda,
                        "Pais": inv.Pais,
                        "Marca_Costo": "",
                        "Cuenta_Costo": "",
                        "Descripcion_Raw": "NO_SE_DETECTARON_LINEAS",
                    }
                )

            # Auditoría
            if show_audit:
                audit_rows.append(
                    {"Documento": uf.name, "Longitud_Texto": len(text), "Texto": (text or "")[:max_audit_chars]}
                )

        except Exception as e:
            err_row = {c: "" for c in FIN_COLS}
            err_row["Documento"] = uf.name
            err_row["Metodo_Extraccion"] = "ERROR"
            err_row["Error"] = str(e)
            fin_rows.append(err_row)

            line_rows.append(
                {
                    "Documento": uf.name,
                    "Linea": "",
                    "Codigo_Item": "",
                    "Descripcion": "",
                    "Cantidad": None,
                    "Unidad": "",
                    "Precio_Unitario": None,
                    "Descuento": None,
                    "Subtotal_Linea": None,
                    "Impuesto_Linea": None,
                    "Total_Linea": None,
                    "Moneda": "",
                    "Pais": "",
                    "Marca_Costo": "",
                    "Cuenta_Costo": "",
                    "Descripcion_Raw": f"ERROR: {e}",
                }
            )

            if show_audit:
                audit_rows.append({"Documento": uf.name, "Longitud_Texto": 0, "Texto": f"ERROR: {e}"})

        prog.progress(int((idx / total_files) * 100))

    status.success("✅ Proceso finalizado. Generando Excel…")

    df_fin = pd.DataFrame(fin_rows, columns=FIN_COLS)
    df_lines = pd.DataFrame(line_rows, columns=LINE_COLS)
    df_audit = pd.DataFrame(audit_rows) if show_audit else pd.DataFrame(columns=["Documento", "Longitud_Texto", "Texto"])

    st.subheader("Vista previa (FINANZAS_FACTURAS)")
    st.dataframe(df_fin, use_container_width=True)

    st.subheader("Vista previa (LINEAS_FACTURA)")
    st.dataframe(df_lines.head(200), use_container_width=True)

    scanned_count = (df_fin["Probable_Escaneado"] == "SI").sum()
    if scanned_count:
        st.warning(
            f"Detecté **{scanned_count}** PDF(s) con muy poco texto (probablemente escaneados). "
            "Sin OCR, las líneas podrían salir vacías."
        )

    excel_bytes = build_excel_bytes(df_fin, df_lines, df_audit)
    filename = f"SED_Facturas_ContabilidadCO_Lineas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.download_button(
        "⬇️ Descargar Excel (Encabezado + Líneas)",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
