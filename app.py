import io
import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# --- Logging (visible en logs de Streamlit Cloud) ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("pdf_invoice_extractor")


# =========================
# Helpers: Secrets / Config
# =========================
def get_azure_settings() -> Tuple[Optional[str], Optional[str]]:
    """
    Reads Azure Document Intelligence endpoint and key from st.secrets or env-like patterns.
    In Streamlit Cloud, define these keys in the app Secrets.
    """
    endpoint = None
    key = None

    # Preferred: st.secrets
    try:
        endpoint = st.secrets.get("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT", None)
        key = st.secrets.get("AZURE_DOCUMENT_INTELLIGENCE_KEY", None)
    except Exception:
        pass

    return endpoint, key


def safe_str(x: Any) -> str:
    if x is None:
        return ""
    return str(x)


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M%S")


# ==========================================
# Extract text from digital PDF (no OCR)
# ==========================================
def extract_text_pymupdf(pdf_bytes: bytes) -> str:
    """
    Fast extraction for text-based PDFs. If PDF is scanned (images),
    this usually returns little/empty text.
    """
    try:
        import fitz  # PyMuPDF
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            parts = []
            for page in doc:
                parts.append(page.get_text("text") or "")
        return "\n".join(parts).strip()
    except Exception as e:
        logger.exception("PyMuPDF extraction failed")
        raise RuntimeError(f"Error extrayendo texto con PyMuPDF: {e}") from e


# ==========================================
# Azure Document Intelligence (Invoice)
# ==========================================
def analyze_invoice_azure(pdf_bytes: bytes, endpoint: str, key: str) -> Dict[str, Any]:
    """
    Uses Azure AI Document Intelligence (prebuilt-invoice) to extract fields & line items.
    """
    try:
        from azure.core.credentials import AzureKeyCredential
        from azure.ai.documentintelligence import DocumentIntelligenceClient
        from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
    except Exception as e:
        raise RuntimeError(
            "No se pudo importar Azure SDK. Verifica requirements.txt."
        ) from e

    try:
        client = DocumentIntelligenceClient(endpoint=endpoint, credential=AzureKeyCredential(key))
        poller = client.begin_analyze_document(
            model_id="prebuilt-invoice",
            analyze_request=AnalyzeDocumentRequest(bytes_source=pdf_bytes),
        )
        result = poller.result()
        return result.as_dict()  # serializable-ish dict
    except Exception as e:
        logger.exception("Azure Document Intelligence failed")
        raise RuntimeError(f"Error analizando con Azure Document Intelligence: {e}") from e


def flatten_invoice_result(result_dict: Dict[str, Any]) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """
    From Azure result dict -> (header_fields, line_items).
    Output is normalized to simplify Excel export.
    """
    header: Dict[str, Any] = {}
    items: List[Dict[str, Any]] = []

    documents = result_dict.get("documents") or []
    if not documents:
        return header, items

    doc0 = documents[0]
    fields = doc0.get("fields") or {}

    def read_field(name: str) -> Any:
        f = fields.get(name) or {}
        # Prefer content/value depending on type
        return f.get("value") if "value" in f else f.get("content")

    # Common invoice fields (may vary by country/template)
    header_map = {
        "VendorName": "Proveedor",
        "VendorTaxId": "NIT_Proveedor",
        "CustomerName": "Cliente",
        "CustomerTaxId": "NIT_Cliente",
        "InvoiceId": "Factura_Numero",
        "InvoiceDate": "Factura_Fecha",
        "DueDate": "Vencimiento_Fecha",
        "PurchaseOrder": "Orden_Compra",
        "Subtotal": "Subtotal",
        "TotalTax": "Impuestos",
        "InvoiceTotal": "Total_Factura",
        "AmountDue": "Saldo_Pendiente",
        "BillingAddress": "Direccion_Facturacion",
        "ShippingAddress": "Direccion_Envio",
        "PaymentTerm": "Termino_Pago",
        "Currency": "Moneda",
    }

    for az_name, col_name in header_map.items():
        header[col_name] = read_field(az_name)

    # Also keep confidence if available (optional)
    header["_confidence"] = doc0.get("confidence")

    # Line items
    line_items_field = fields.get("Items") or {}
    li_value = line_items_field.get("value") or []
    for idx, it in enumerate(li_value, start=1):
        # Each item is a dict with "value" containing fields
        item_fields = (it.get("value") or {})
        row = {
            "Item": idx,
            "Descripcion": (item_fields.get("Description") or {}).get("value") or (item_fields.get("Description") or {}).get("content"),
            "Cantidad": (item_fields.get("Quantity") or {}).get("value") or (item_fields.get("Quantity") or {}).get("content"),
            "PrecioUnitario": (item_fields.get("UnitPrice") or {}).get("value") or (item_fields.get("UnitPrice") or {}).get("content"),
            "TotalLinea": (item_fields.get("Amount") or {}).get("value") or (item_fields.get("Amount") or {}).get("content"),
            "SKU": (item_fields.get("ProductCode") or {}).get("value") or (item_fields.get("ProductCode") or {}).get("content"),
            "Unidad": (item_fields.get("Unit") or {}).get("value") or (item_fields.get("Unit") or {}).get("content"),
        }
        items.append(row)

    return header, items


def build_excel_bytes(
    summary_rows: List[Dict[str, Any]],
    detected_fields_rows: List[Dict[str, Any]],
    line_items_rows: List[Dict[str, Any]],
) -> bytes:
    """
    Creates an Excel in-memory with 3 sheets.
    """
    out = io.BytesIO()

    df_summary = pd.DataFrame(summary_rows)
    df_fields = pd.DataFrame(detected_fields_rows)
    df_items = pd.DataFrame(line_items_rows)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Resumen")
        df_fields.to_excel(writer, index=False, sheet_name="Campos_Detectados")
        df_items.to_excel(writer, index=False, sheet_name="LineItems")

    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="SED | PDF (Factura) → Excel (OCR)", layout="wide")

st.title("📄➡️📊 SED | OCR de Facturas (PDF) a Excel")
st.caption(
    "Carga uno o varios PDFs. Si el PDF es escaneado, se recomienda Azure Document Intelligence (prebuilt-invoice) "
    "para OCR + extracción estructurada."
)

with st.sidebar:
    st.header("⚙️ Configuración")
    endpoint, key = get_azure_settings()

    st.write("**Modo de extracción**")
    mode = st.radio(
        "Selecciona el modo",
        options=["Automático", "Solo texto (PDF digital)", "Azure OCR (Factura)"],
        index=0,
        help=(
            "Automático: intenta texto directo, y si no hay texto suficiente usa Azure (si está configurado)."
        ),
    )

    st.divider()
    st.subheader("🔐 Estado de Azure")
    if endpoint and key:
        st.success("Azure configurado por Secrets ✅")
    else:
        st.warning("Azure NO configurado (solo funcionará extracción de texto digital).")

    with st.expander("Cómo configurar Secrets (Streamlit Cloud)"):
        st.code(
            """AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT="https://<tu-recurso>.cognitiveservices.azure.com/"
AZURE_DOCUMENT_INTELLIGENCE_KEY="<tu-key>" """,
            language="bash",
        )

st.divider()

uploaded_files = st.file_uploader(
    "Sube uno o varios PDFs (facturas)",
    type=["pdf"],
    accept_multiple_files=True,
)

colA, colB = st.columns([1, 2], vertical_alignment="center")
with colA:
    process = st.button("🚀 Procesar PDFs", type="primary", use_container_width=True)
with colB:
    st.info("Tip: Si tu factura es un escaneo (imagen), usa **Azure OCR (Factura)** y configura Secrets.")

if process:
    if not uploaded_files:
        st.error("Por favor, sube al menos un archivo PDF.")
        st.stop()

    summary_rows: List[Dict[str, Any]] = []
    detected_fields_rows: List[Dict[str, Any]] = []
    line_items_rows: List[Dict[str, Any]] = []

    prog = st.progress(0)
    status = st.empty()

    total = len(uploaded_files)

    for i, uf in enumerate(uploaded_files, start=1):
        status.write(f"Procesando **{uf.name}** ({i}/{total})…")
        try:
            pdf_bytes = uf.read()

            extracted_text = ""
            azure_used = False
            azure_result: Dict[str, Any] = {}

            # Decide mode
            if mode == "Solo texto (PDF digital)":
                extracted_text = extract_text_pymupdf(pdf_bytes)

            elif mode == "Azure OCR (Factura)":
                if not (endpoint and key):
                    raise RuntimeError("Azure no está configurado. Define Secrets en Streamlit Cloud.")
                azure_result = analyze_invoice_azure(pdf_bytes, endpoint, key)
                azure_used = True

            else:  # Automático
                extracted_text = extract_text_pymupdf(pdf_bytes)
                # heuristic: if too little text, try Azure if available
                if len(extracted_text) < 50 and endpoint and key:
                    azure_result = analyze_invoice_azure(pdf_bytes, endpoint, key)
                    azure_used = True

            # Normalize outputs
            doc_id = f"{uf.name}"

            if azure_used:
                header, items = flatten_invoice_result(azure_result)

                # Summary row: 1 per document
                row = {"Documento": doc_id, "Metodo": "Azure Document Intelligence"}
                row.update({k: safe_str(v) for k, v in header.items()})
                summary_rows.append(row)

                # Key/value fields sheet (store all fields returned)
                documents = azure_result.get("documents") or []
                fields = ((documents[0].get("fields") or {}) if documents else {})
                for k, v in fields.items():
                    detected_fields_rows.append(
                        {
                            "Documento": doc_id,
                            "Campo": k,
                            "Valor": safe_str(v.get("value") if isinstance(v, dict) else v),
                            "Contenido": safe_str(v.get("content") if isinstance(v, dict) else ""),
                            "Confianza": safe_str(v.get("confidence") if isinstance(v, dict) else ""),
                        }
                    )

                # Line items sheet
                if items:
                    for it in items:
                        it_row = {"Documento": doc_id}
                        it_row.update({k: safe_str(v) for k, v in it.items()})
                        line_items_rows.append(it_row)
                else:
                    # Keep at least a marker row
                    line_items_rows.append({"Documento": doc_id, "Item": "", "Descripcion": "", "Cantidad": "", "PrecioUnitario": "", "TotalLinea": ""})

            else:
                # Text-only fallback
                summary_rows.append(
                    {
                        "Documento": doc_id,
                        "Metodo": "Texto directo (PDF digital)",
                        "Texto_Extraido_Longitud": len(extracted_text),
                        "Texto_Extraido_Preview": extracted_text[:500],
                    }
                )
                detected_fields_rows.append(
                    {
                        "Documento": doc_id,
                        "Campo": "RAW_TEXT",
                        "Valor": extracted_text[:32000],  # avoid gigantic cells
                        "Contenido": "",
                        "Confianza": "",
                    }
                )

        except Exception as e:
            logger.exception("Error processing file")
            summary_rows.append({"Documento": uf.name, "Metodo": "ERROR", "Detalle_Error": safe_str(e)})
            st.warning(f"⚠️ {uf.name}: {e}")

        prog.progress(int((i / total) * 100))

    status.success("Proceso finalizado. Generando Excel…")

    try:
        excel_bytes = build_excel_bytes(summary_rows, detected_fields_rows, line_items_rows)
        filename = f"SED_facturas_OCR_{now_stamp()}.xlsx"

        st.subheader("✅ Resultado")
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("Ver detalle (Campos_Detectados)"):
            st.dataframe(pd.DataFrame(detected_fields_rows).head(200), use_container_width=True)

        with st.expander("Ver detalle (LineItems)"):
            st.dataframe(pd.DataFrame(line_items_rows).head(200), use_container_width=True)

    except Exception as e:
        logger.exception("Excel build failed")
        st.error(f"No se pudo generar el Excel: {e}")
