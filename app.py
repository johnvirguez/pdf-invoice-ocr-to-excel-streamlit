import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# App Config
# =========================
st.set_page_config(page_title="SED | Facturas → Excel (Finanzas CO + Líneas)", layout="wide")
APP_TITLE = "📄➡️📊 SED | Facturas PDF → Excel (Contabilidad Colombia + Líneas)"
APP_SUBTITLE = (
    "Extracción desde PDFs digitales (sin OCR). "
    "Genera encabezado (totales) + todas las líneas. "
    "Excel final: sin cuadrícula, fuente Century Gothic 10 y líneas agrupadas por factura."
)

# Encabezado (1 fila por factura)
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
    # para marcar costo a nivel factura
    "Costo_Factura_Marcado",
    "Probable_Escaneado",
    "Metodo_Extraccion",
    "Error",
]

# Líneas (múltiples filas por factura) — incluir Factura_Numero para agrupar
LINE_COLS = [
    "Factura_Numero",
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
    # columnas para marcar costo por línea
    "Marca_Costo",
    "Cuenta_Costo",
    "Descripcion_Raw",
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

    Costo_Factura_Marcado: str = ""
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
    if " CRC" in t or "Moneda: CRC" in t or "Código Moneda........ CRC" in t or "¢" in t:
        return "CRC", "¢"
    if " COP" in t or "pesos" in t.lower() or "Bogotá - Colombia" in t or "$" in t:
        return "COP", "$"
    return "", ""


def lines(text: str) -> List[str]:
    return [ln.strip() for ln in (text or "").splitlines() if (ln or "").strip()]


# =========================
# Format Detectors
# =========================
def is_forlan_co(text: str) -> bool:
    t = (text or "").upper()
    return "FERRETERIA FORLAN" in t and "FACTURA ELECTR" in t and "VR." in t and "TOTAL A PAGAR" in t


def is_navatec_cr(text: str) -> bool:
    t = (text or "").upper()
    return ("FACTURA ELECTRÓNICA N°" in t or "FACTURA ELECTRONICA N°" in t) and "FACTURAELECTRONICA.CR" in t


def is_tribu_cr_hacienda(text: str) -> bool:
    t = (text or "").upper()
    return "WWW.HACIENDA.GO.CR" in t and "TRIBU-CR" in t and "COMPROBANTE" in t


def is_ciclo_huracan(text: str) -> bool:
    t = (text or "").upper()
    return "CICLO HURACAN" in t and "NO COD PRODUCTO" in t and "TOTAL DE LÍNEA" in t


def is_brujo_caribeno(text: str) -> bool:
    t = (text or "").upper()
    return "EL BRUJO CARIBEÑO" in t and "CÓDIGO UNIDAD CANTIDAD PRECIO" in t


def is_erial_office_depot(text: str) -> bool:
    t = (text or "").upper()
    return "ERIAL BQ" in t and "LINEA SKU" in t and "PRECIO" in t and "IMPUESTO" in t


def is_gustavo_gamboa(text: str) -> bool:
    t = (text or "").upper()
    return "GUSTAVO GAMBOA VILLALOBOS" in t and "# DESCRIPCIÓN / CÓDIGO" in t


# =========================
# Header Parsers
# =========================
def parse_forlan_co_header(text: str, filename: str) -> FinanceInvoice:
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
    resol = find_first([r"Autorizaci[oó]n\s+Electr[oó]nica\s+([0-9]+)"], text)
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
        Metodo_Extraccion="FORLAN CO (header)",
        Error="",
    )


def parse_navatec_cr_header(text: str, filename: str) -> FinanceInvoice:
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
        Metodo_Extraccion="NAVATEC CR (header)",
        Error="",
    )


def parse_generic_header(text: str, filename: str, pais_hint: str = "") -> FinanceInvoice:
    scanned = looks_scanned(text)
    moneda, simbolo = detect_currency(text)

    proveedor = find_first(
        [
            r"Raz[oó]n\s+Social[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
            r"Nombre:\s*([A-ZÁÉÍÓÚÑ0-9\.\s&\-]+)\n",
            r"Emisor[:\s]+([A-ZÁÉÍÓÚÑ0-9&\-\.\s]{4,})",
        ],
        text,
    ).split("\n")[0].strip()

    proveedor_id = find_first(
        [
            r"NIT[:\s]*([0-9\.\-]{6,20})",
            r"Identificaci[oó]n:\s*([0-9]+)",
            r"C[eé]dula:\s*([0-9]+)",
            r"Ident\.\s*Jur[ií]dica:\s*([0-9\-]+)",
        ],
        text,
    )

    factura_num = find_first(
        [
            r"Factura\s+Electr[oó]nica[:\s#]*([0-9]{8,})",
            r"Consecutivo:\s*([0-9]+)",
            r"Factura\s*(?:No\.|Nro\.|N°|#)?\s*[:\s]*([A-Z0-9\-]{3,})",
        ],
        text,
    )

    fecha = find_first(
        [
            r"Fecha:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s+[0-2]\d:[0-5]\d:[0-5]\d)",
            r"Fecha\s+y\s+hora\s+de\s+emisi[oó]n:\s*([0-3]\d[\-\/][01]\d[\-\/][12]\d{3}\s+[0-2]\d:[0-5]\d:[0-5]\d\s*[ap]\.m\.)",
            r"Fecha\s+de\s+emisi[oó]n:\s*([0-3]\d\/[01]\d\/[12]\d{3}\s*[0-2]?\d:[0-5]\d)",
        ],
        text,
    )

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
        Probable_Escaneado="SI" if scanned else "NO",
        Metodo_Extraccion="Genérico (header)",
        Error="",
    )


# =========================
# Line Items Parsers (por formato)
# =========================
def items_forlan_co(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    FORLAN CO: filas tipo:
      1 21147 15.00 SOLDADURA ELECTRICA ... 14,000.00 210,000.00 244,650.00
    Columnas del PDF: Ítem, Código, Cantidad, Descripción, Vr Unitario, Vr Bruto, Vr Total 
    """
    out: List[Dict[str, Any]] = []
    for ln in lines(text):
        # Start: itemNo code qty ... end: unit bruto total
        m = re.match(
            r"^(?P<item>\d+)\s+(?P<codigo>\d{3,})\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<desc>.+?)\s+(?P<unit>[0-9\.,]+)\s+(?P<bruto>[0-9\.,]+)\s+(?P<total>[0-9\.,]+)\s*$",
            ln
        )
        if not m:
            continue

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": m.group("item"),
            "Codigo_Item": m.group("codigo"),
            "Descripcion": m.group("desc").strip(),
            "Cantidad": parse_number_latam(m.group("qty")),
            "Unidad": "",
            "Precio_Unitario": parse_number_latam(m.group("unit")),
            "Descuento": None,
            "Subtotal_Linea": parse_number_latam(m.group("bruto")),
            "Impuesto_Linea": None,  # en Forlan viene el % pero no el impuesto por línea
            "Total_Linea": parse_number_latam(m.group("total")),
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": ln,
        })

    return out


def items_navatec_cr(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    NAVATEC: líneas tipo:
      001 1.00 Unid MTR001 ... 741,300.84 0.00 741.300,84 96,369.11 
    """
    out: List[Dict[str, Any]] = []
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

    for ln in lines(text):
        m = pat.match(ln)
        if not m:
            continue

        subtotal_linea = parse_number_latam(m.group("subtotal"))
        imp = parse_number_latam(m.group("imp"))
        total_linea = (subtotal_linea + imp) if (subtotal_linea is not None and imp is not None) else None

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": m.group("linea"),
            "Codigo_Item": m.group("codigo"),
            "Descripcion": m.group("desc").strip(),
            "Cantidad": parse_number_latam(m.group("cantidad")),
            "Unidad": m.group("unidad"),
            "Precio_Unitario": parse_number_latam(m.group("precio")),
            "Descuento": parse_number_latam(m.group("descuento")),
            "Subtotal_Linea": subtotal_linea,
            "Impuesto_Linea": imp,
            "Total_Linea": total_linea,
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": ln,
        })

    return out


def items_tribu_hacienda_cr(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    TRIBU-CR / Hacienda: en página 2:
      Línea Código Detalle del Producto
      1 8419000000000 REPARACION VEHICULO MECANICO
      1,00 Unidad 20.000,00 ... 0,00 22.600,00 
    """
    out: List[Dict[str, Any]] = []
    ln_list = lines(text)

    i = 0
    while i < len(ln_list):
        ln = ln_list[i]
        m = re.match(r"^(?P<linea>\d+)\s+(?P<codigo>\d{10,})\s+(?P<desc>.+)$", ln)
        if not m:
            i += 1
            continue

        linea = m.group("linea")
        codigo = m.group("codigo")
        desc = m.group("desc").strip()

        # busca la siguiente línea con cantidad/unidad/precio/total
        j = i + 1
        blob = ""
        while j < len(ln_list):
            if ln_list[j].upper().startswith("OBSERVACIONES"):
                break
            # siguiente item?
            if re.match(r"^\d+\s+\d{10,}\s+", ln_list[j]):
                break
            blob += " " + ln_list[j]
            j += 1

        # En blob buscamos: qty unit price monto descuento total (el orden varía por saltos)
        # Nos quedamos con los últimos 4 números como: precio/monto/descuento/total (heurístico)
        nums = re.findall(r"[0-9]{1,3}(?:[0-9\.,]*[0-9])", blob)
        # qty tiene coma decimal (1,00)
        qty = find_first([r"(\d+,\d+)\s+Unidad", r"(\d+,\d+)\s+Servicios", r"(\d+,\d+)\s+\w+"], blob)

        # toma los últimos 4 números del blob como [precio, monto, descuento, total] si existen
        precio = monto = descuento = total = None
        if len(nums) >= 4:
            precio = parse_number_latam(nums[-4])
            monto = parse_number_latam(nums[-3])
            descuento = parse_number_latam(nums[-2])
            total = parse_number_latam(nums[-1])

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": linea,
            "Codigo_Item": codigo,
            "Descripcion": desc,
            "Cantidad": parse_number_latam(qty),
            "Unidad": "Unidad",
            "Precio_Unitario": precio,
            "Descuento": descuento,
            "Subtotal_Linea": monto,
            "Impuesto_Linea": None,
            "Total_Linea": total,
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": (ln + " | " + blob.strip())[:1000],
        })

        i = j
    return out


def items_ciclo_huracan(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    CICLO HURACAN:
      1 04692 BINOCULAR 8X21
      1.00 Unid CRC 10,619.46903 ... CRC 12,000.00 :contentReference[oaicite:9]{index=9}
    """
    out: List[Dict[str, Any]] = []
    ln_list = lines(text)

    i = 0
    while i < len(ln_list):
        ln = ln_list[i]
        m = re.match(r"^(?P<linea>\d+)\s+(?P<codigo>\d+)\s+(?P<desc>.+)$", ln)
        if not m:
            i += 1
            continue

        linea = m.group("linea")
        codigo = m.group("codigo")
        desc = m.group("desc").strip()

        # Captura siguiente bloque con cantidades y valores
        blob = ""
        j = i + 1
        while j < len(ln_list):
            if ln_list[j].upper().startswith("COMENTARIO") or ln_list[j].upper().startswith("COMENTARIO:"):
                break
            if re.match(r"^\d+\s+\d+\s+", ln_list[j]):
                break
            blob += " " + ln_list[j]
            j += 1

        qty = find_first([r"(\d+(?:\.\d+)?)\s+Unid"], blob)
        # valores: precio, impuesto, total (en texto aparece IVA 13% CRC 1,380.53 CRC 12,000.00)
        precio = find_first([r"CRC\s*([0-9\.,]+)"], blob)
        total = find_first([r"CRC\s*([0-9\.,]+)\s*$"], blob)

        # impuesto: busca "CRC x,xxx.xx" antes del total
        imp = None
        m_imp = re.search(r"IVA\s*13%.*?CRC\s*([0-9\.,]+)", blob, re.IGNORECASE)
        if m_imp:
            imp = parse_number_latam(m_imp.group(1))

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": linea,
            "Codigo_Item": codigo,
            "Descripcion": desc,
            "Cantidad": parse_number_latam(qty),
            "Unidad": "Unid",
            "Precio_Unitario": parse_number_latam(precio),
            "Descuento": 0.0,
            "Subtotal_Linea": None,
            "Impuesto_Linea": imp,
            "Total_Linea": parse_number_latam(total),
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": (ln + " | " + blob.strip())[:1000],
        })

        i = j
    return out


def items_brujo_caribeno(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    EL BRUJO CARIBEÑO:
      C01 Al 1.00 300,000.00 0.00 300,000.00 39,000.00 (descr arriba) :contentReference[oaicite:10]{index=10}
    """
    out: List[Dict[str, Any]] = []
    ln_list = lines(text)

    # intenta capturar descripción larga anterior a la línea numérica
    last_desc = ""
    for ln in ln_list:
        if "Servicios de alquiler" in ln:
            last_desc = ln.strip()

        m = re.match(r"^(?P<codigo>[A-Z0-9]+)\s+(?P<unidad>\w+)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<precio>[0-9\.,]+)\s+(?P<descnt>[0-9\.,]+)\s+(?P<subt>[0-9\.,]+)\s+(?P<imp>[0-9\.,]+)\s*$", ln)
        if not m:
            continue

        subtotal = parse_number_latam(m.group("subt"))
        imp = parse_number_latam(m.group("imp"))
        total = (subtotal + imp) if (subtotal is not None and imp is not None) else None

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": "1",
            "Codigo_Item": m.group("codigo"),
            "Descripcion": last_desc or "SERVICIO",
            "Cantidad": parse_number_latam(m.group("qty")),
            "Unidad": m.group("unidad"),
            "Precio_Unitario": parse_number_latam(m.group("precio")),
            "Descuento": parse_number_latam(m.group("descnt")),
            "Subtotal_Linea": subtotal,
            "Impuesto_Linea": imp,
            "Total_Linea": total,
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": ln,
        })
    return out


def items_erial_office_depot(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    ERIAL BQ (Office Depot): líneas tipo:
      1 3212900039900 ... 1.00 Unid 876.11 876.11 113.89 13.00 0.00 990.00 :contentReference[oaicite:11]{index=11}
    """
    out: List[Dict[str, Any]] = []
    for ln in lines(text):
        m = re.match(
            r"^(?P<linea>\d+)\s+(?P<sku>\d{10,})\s+(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<uni>\w+)\s+(?P<pu>[0-9\.,]+)\s+(?P<subt>[0-9\.,]+)\s+(?P<imp>[0-9\.,]+)\s+(?P<pct>[0-9\.,]+)\s+(?P<descnt>[0-9\.,]+)\s+(?P<total>[0-9\.,]+)\s*$",
            ln
        )
        if not m:
            continue

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": m.group("linea"),
            "Codigo_Item": m.group("sku"),
            "Descripcion": m.group("desc").strip(),
            "Cantidad": parse_number_latam(m.group("qty")),
            "Unidad": m.group("uni"),
            "Precio_Unitario": parse_number_latam(m.group("pu")),
            "Descuento": parse_number_latam(m.group("descnt")),
            "Subtotal_Linea": parse_number_latam(m.group("subt")),
            "Impuesto_Linea": parse_number_latam(m.group("imp")),
            "Total_Linea": parse_number_latam(m.group("total")),
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": ln,
        })
    return out


def items_gustavo_gamboa(text: str, inv: FinanceInvoice) -> List[Dict[str, Any]]:
    """
    Gustavo Gamboa: líneas tipo:
      1 RU1487 SECTOR CONCEPCION 1.00 837,405.00 Serv Prof 0.00 13.00 % I.V.A 108,862.65 946,267.65 :contentReference[oaicite:12]{index=12}
    """
    out: List[Dict[str, Any]] = []
    for ln in lines(text):
        m = re.match(
            r"^(?P<linea>\d+)\s+(?P<codigo>[A-Z0-9]+)\s+(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<precio>[0-9\.,]+)\s+(?P<uni>Serv\s+Prof|\w+)\s+(?P<descnt>[0-9\.,]+)\s+(?P<pct>[0-9\.,]+)\s+%.*?\s+(?P<imp>[0-9\.,]+)\s+(?P<total>[0-9\.,]+)\s*$",
            ln,
            re.IGNORECASE
        )
        if not m:
            continue

        total = parse_number_latam(m.group("total"))
        imp = parse_number_latam(m.group("imp"))
        subtotal = (total - imp) if (total is not None and imp is not None) else None

        out.append({
            "Factura_Numero": inv.Factura_Numero,
            "Documento": inv.Documento,
            "Linea": m.group("linea"),
            "Codigo_Item": m.group("codigo"),
            "Descripcion": m.group("desc").strip(),
            "Cantidad": parse_number_latam(m.group("qty")),
            "Unidad": m.group("uni").strip(),
            "Precio_Unitario": parse_number_latam(m.group("precio")),
            "Descuento": parse_number_latam(m.group("descnt")),
            "Subtotal_Linea": subtotal,
            "Impuesto_Linea": imp,
            "Total_Linea": total,
            "Moneda": inv.Moneda,
            "Pais": inv.Pais,
            "Marca_Costo": "",
            "Cuenta_Costo": "",
            "Descripcion_Raw": ln,
        })
    return out


# =========================
# Excel Formatting + Grouping
# =========================
def autosize_columns(ws):
    # autosize simple (cap 60)
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def apply_global_excel_formatting(wb):
    font = Font(name="Century Gothic", size=10)

    for ws in wb.worksheets:
        # Disable gridlines
        ws.sheet_view.showGridLines = False

        # Apply font to used range
        for row in ws.iter_rows():
            for cell in row:
                cell.font = font

        # Freeze header row
        ws.freeze_panes = "A2"

        # Autofilter
        ws.auto_filter.ref = ws.dimensions

        # Autosize
        autosize_columns(ws)


def group_line_items_by_invoice(ws, factura_col_letter: str = "A"):
    """
    Agrupa filas por Factura_Numero (col A en LINEAS_FACTURA)
    """
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_view.showOutlineSymbols = True

    # lee valores desde fila 2
    current = None
    start_row = None

    max_row = ws.max_row
    for r in range(2, max_row + 1):
        val = ws[f"{factura_col_letter}{r}"].value
        if val != current:
            # cerrar grupo anterior
            if current is not None and start_row is not None:
                end_row = r - 1
                if end_row > start_row:
                    ws.row_dimensions.group(start_row, end_row, outline_level=1, hidden=False)
            # abrir nuevo
            current = val
            start_row = r

    # cerrar último
    if current is not None and start_row is not None:
        end_row = max_row
        if end_row > start_row:
            ws.row_dimensions.group(start_row, end_row, outline_level=1, hidden=False)


def build_excel_bytes(df_fin: pd.DataFrame, df_lines: pd.DataFrame, df_audit: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_fin.to_excel(writer, index=False, sheet_name="FINANZAS_FACTURAS")
        df_lines.to_excel(writer, index=False, sheet_name="LINEAS_FACTURA")
        df_audit.to_excel(writer, index=False, sheet_name="AUDITORIA_TEXTO")

    out.seek(0)
    wb = load_workbook(out)

    apply_global_excel_formatting(wb)

    # Agrupar en LINEAS_FACTURA por Factura_Numero (columna A)
    if "LINEAS_FACTURA" in wb.sheetnames:
        ws = wb["LINEAS_FACTURA"]
        group_line_items_by_invoice(ws, factura_col_letter="A")

    final = io.BytesIO()
    wb.save(final)
    return final.getvalue()


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

            # Header
            if is_forlan_co(text):
                inv = parse_forlan_co_header(text, uf.name)
                inv.Metodo_Extraccion = "FORLAN CO (header + líneas)"
                items = items_forlan_co(text, inv)
            elif is_navatec_cr(text):
                inv = parse_navatec_cr_header(text, uf.name)
                inv.Metodo_Extraccion = "NAVATEC CR (header + líneas)"
                items = items_navatec_cr(text, inv)
            elif is_tribu_cr_hacienda(text):
                inv = parse_generic_header(text, uf.name, pais_hint="CR")
                inv.Pais = "CR"
                inv.Metodo_Extraccion = "TRIBU-CR (header + líneas)"
                items = items_tribu_hacienda_cr(text, inv)
            elif is_ciclo_huracan(text):
                inv = parse_generic_header(text, uf.name, pais_hint="CR")
                inv.Pais = "CR"
                inv.Metodo_Extraccion = "CICLO HURACAN (header + líneas)"
                items = items_ciclo_huracan(text, inv)
            elif is_brujo_caribeno(text):
                inv = parse_generic_header(text, uf.name, pais_hint="CR")
                inv.Pais = "CR"
                inv.Metodo_Extraccion = "EL BRUJO CARIBEÑO (header + líneas)"
                items = items_brujo_caribeno(text, inv)
            elif is_erial_office_depot(text):
                inv = parse_generic_header(text, uf.name, pais_hint="CR")
                inv.Pais = "CR"
                inv.Metodo_Extraccion = "ERIAL BQ (header + líneas)"
                items = items_erial_office_depot(text, inv)
            elif is_gustavo_gamboa(text):
                inv = parse_generic_header(text, uf.name, pais_hint="CR")
                inv.Pais = "CR"
                inv.Metodo_Extraccion = "GUSTAVO GAMBOA (header + líneas)"
                items = items_gustavo_gamboa(text, inv)
            else:
                inv = parse_generic_header(text, uf.name)
                inv.Metodo_Extraccion = "Genérico (header + líneas)"
                items = []  # si es genérico, evitamos crear filas vacías sin intentar algo riesgoso

            # Encabezado (si faltan moneda/símbolo por genérico, intenta inferir)
            if not inv.Moneda:
                inv.Moneda, inv.Simbolo_Moneda = detect_currency(text)

            fin_row = inv.__dict__
            fin_rows.append({c: fin_row.get(c, "") for c in FIN_COLS})

            # Líneas: si no hay líneas, solo 1 fila con SIN_LINEAS_DETECTADAS (pero con Factura_Numero)
            if items:
                for it in items:
                    line_rows.append({c: it.get(c, "") for c in LINE_COLS})
            else:
                line_rows.append({
                    "Factura_Numero": inv.Factura_Numero,
                    "Documento": inv.Documento,
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
                    "Descripcion_Raw": "SIN_LINEAS_DETECTADAS",
                })

            # Auditoría
            if show_audit:
                audit_rows.append({
                    "Documento": uf.name,
                    "Longitud_Texto": len(text),
                    "Texto": (text or "")[:max_audit_chars],
                })

        except Exception as e:
            err_row = {c: "" for c in FIN_COLS}
            err_row["Documento"] = uf.name
            err_row["Metodo_Extraccion"] = "ERROR"
            err_row["Error"] = str(e)
            fin_rows.append(err_row)

            line_rows.append({
                "Factura_Numero": "",
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
            })

            if show_audit:
                audit_rows.append({"Documento": uf.name, "Longitud_Texto": 0, "Texto": f"ERROR: {e}"})

        prog.progress(int((idx / total_files) * 100))

    status.success("✅ Proceso finalizado. Generando Excel…")

    df_fin = pd.DataFrame(fin_rows, columns=FIN_COLS)

    # Ordenamos líneas por Factura_Numero para que el agrupado quede perfecto
    df_lines = pd.DataFrame(line_rows, columns=LINE_COLS)
    df_lines["Factura_Numero"] = df_lines["Factura_Numero"].fillna("")
    df_lines = df_lines.sort_values(by=["Factura_Numero", "Documento", "Linea"], kind="stable").reset_index(drop=True)

    df_audit = pd.DataFrame(audit_rows) if show_audit else pd.DataFrame(columns=["Documento", "Longitud_Texto", "Texto"])

    st.subheader("Vista previa (FINANZAS_FACTURAS)")
    st.dataframe(df_fin, use_container_width=True)

    st.subheader("Vista previa (LINEAS_FACTURA)")
    st.dataframe(df_lines.head(200), use_container_width=True)

    excel_bytes = build_excel_bytes(df_fin, df_lines, df_audit)
    filename = f"SED_Facturas_ContabilidadCO_Lineas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    st.download_button(
        "⬇️ Descargar Excel (Encabezado + Líneas, Formato Contabilidad)",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
