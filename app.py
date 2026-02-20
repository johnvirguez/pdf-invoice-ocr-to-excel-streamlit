import io
from datetime import datetime
import pandas as pd
import streamlit as st
import pdfplumber

st.set_page_config(page_title="PDF Facturas → Excel", layout="wide")

st.title("📄➡️📊 OCR Básico de Facturas (PDF → Excel)")
st.caption("Extrae texto de facturas PDF digitales y genera Excel.")

uploaded_files = st.file_uploader(
    "Sube uno o varios PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("🚀 Procesar PDFs"):

    if not uploaded_files:
        st.error("Sube al menos un archivo PDF.")
        st.stop()

    summary_rows = []
    progress = st.progress(0)

    for idx, file in enumerate(uploaded_files):
        try:
            with pdfplumber.open(file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""

            summary_rows.append({
                "Documento": file.name,
                "Longitud_Texto": len(text),
                "Preview_Texto": text[:500]
            })

        except Exception as e:
            summary_rows.append({
                "Documento": file.name,
                "Error": str(e)
            })

        progress.progress((idx + 1) / len(uploaded_files))

    df = pd.DataFrame(summary_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumen")

    st.success("Proceso finalizado ✅")

    st.download_button(
        "⬇️ Descargar Excel",
        data=output.getvalue(),
        file_name=f"facturas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
