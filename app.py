import streamlit as st
import pandas as pd

st.set_page_config(page_title="Mi primera aplicación", layout="wide")

st.title("🚀 Mi primera aplicación web en Streamlit")

st.write("Escribe tu nombre y genera una gráfica simple.")

# Entrada de usuario
nombre = st.text_input("Escribe tu nombre")

if nombre:
    st.success(f"Hola {nombre}, bienvenido a tu primera app en la nube ☁️")

    # Datos de ejemplo
    data = pd.DataFrame({
        "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo"],
        "Ventas": [100, 150, 80, 200, 170]
    })

    st.subheader("📊 Ejemplo de gráfico (sin matplotlib)")
    st.line_chart(data.set_index("Mes"))

    st.subheader("📋 Datos utilizados")
    st.dataframe(data, use_container_width=True)
