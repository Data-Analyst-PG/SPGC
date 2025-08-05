import streamlit as st
import hashlib
import base64
from supabase import create_client
from PIL import Image

# =========================
# ✅ ENCABEZADO Y MENÚ
# =========================

# Ruta al logo
LOGO_CLARO = "Color PGL MS.png"
LOGO_OSCURO = "White PGL MS.png"

@st.cache_data
def image_to_base64(img_path):
    with open(img_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()

logo_claro_b64 = image_to_base64(LOGO_CLARO)
logo_oscuro_b64 = image_to_base64(LOGO_OSCURO)

# Mostrar encabezado con logo dinámico
st.markdown(f"""
    <div style='text-align: center;'>
        <img src="data:image/png;base64,{logo_claro_b64}" class="logo-light" style="height: 120px; margin-bottom: 20px;">
        <img src="data:image/png;base64,{logo_oscuro_b64}" class="logo-dark" style="height: 120px; margin-bottom: 20px;">
    </div>
    <h1 style='text-align: center; color: #003366;'>SPGC</h1>
    <p style='text-align: center;'>Sistema de Prorratero de Gastos y Costos</p>
    <hr style='margin-top: 20px; margin-bottom: 30px;'>
    <style>
    @media (prefers-color-scheme: dark) {{
        .logo-light {{ display: none; }}
        .logo-dark {{ display: inline; }}
    }}
    @media (prefers-color-scheme: light) {{
        .logo-light {{ display: inline; }}
        .logo-dark {{ display: none; }}
    }}
    </style>
""", unsafe_allow_html=True)

st.info("Selecciona una opción desde el menú lateral para comenzar 🚀")

# Instrucciones de navegación
st.subheader("📂 Módulos disponibles")
st.markdown("""
- **🧾 Módulo 1: Prorrateo de Gastos Generales
""")
