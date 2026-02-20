import io
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="PDF Facturas → Excel", layout="wide")
st.title("📄➡️📊 PDF Facturas → Excel (sin Azure)")

try:
    import pdfplumber
except Exception as e:
    st.error("No se pudo importar pdfplumber. Revisa requirements.txt / Logs.")
    st.exception(e)
    st.stop()

uploaded_files = st.file_uploader(
    "Sube uno o varios PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("🚀 Procesar PDFs", type="primary"):
    if not uploaded_files:
        st.error("Sube al menos un archivo PDF.")
        st.stop()

    rows = []
    prog = st.progress(0)

    for i, f in enumerate(uploaded_files, start=1):
        try:
            with pdfplumber.open(f) as pdf:
                text_parts = []
                for page in pdf.pages:
                    text_parts.append(page.extract_text() or "")
                text = "\n".join(text_parts).strip()

            rows.append({
                "Documento": f.name,
                "Longitud_Texto": len(text),
                "Preview_Texto": text[:800],
                "Error": ""
            })
        except Exception as e:
            rows.append({
                "Documento": f.name,
                "Longitud_Texto": "",
                "Preview_Texto": "",
                "Error": str(e)
            })

        prog.progress(i / len(uploaded_files))

    df = pd.DataFrame(rows)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumen")

    st.success("✅ Listo")
    st.download_button(
        "⬇️ Descargar Excel",
        data=out.getvalue(),
        file_name=f"facturas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(df, use_container_width=True)
