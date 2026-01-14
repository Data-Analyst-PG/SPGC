import streamlit as st
import pandas as pd
from supabase import create_client

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
st.set_page_config(page_title="Distribución CI Clientes", layout="wide")

# ============================================================
# CONEXIÓN SUPABASE
# ============================================================
SUPABASE_URL = st.secrets["supabase"]["url"]
SUPABASE_KEY = st.secrets["supabase"]["key"]
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# ============================================================
# SELECCIÓN DE EMPRESA
# ============================================================
st.title("👥 Distribución CI Clientes")

empresa = st.selectbox(
    "Selecciona la empresa",
    [
        "Lincoln Freight",
        "Set Logis Plus",
        "Picus Carrier",
        "Igloo Carrier",
    ],
)

st.divider()

# ============================================================
# CONFIGURACIÓN POR EMPRESA
# ============================================================
if empresa == "Lincoln Freight":
    COL_FECHA = "Trip Date"
    COL_UNIDAD = "Truck"
    COL_REMOLQUE = "Trailer"
    COL_CLIENTE = "Customer"
    COL_DISTANCIA = "Real Miles"
    USA_OPERADOR = True
    PREFIJO_REMOLQUE = "LF"
    TIPO_DISTANCIA = "MILLAS"

elif empresa == "Set Logis Plus":
    COL_FECHA = "Trip Date"
    COL_UNIDAD = "Truck"
    COL_REMOLQUE = "Trailer"
    COL_CLIENTE = "Customer"
    COL_DISTANCIA = "Real Miles"
    USA_OPERADOR = True
    PREFIJO_REMOLQUE = "STL"
    TIPO_DISTANCIA = "MILLAS"

elif empresa in ["Picus Carrier", "Igloo Carrier"]:
    COL_FECHA = "Fecha"
    COL_UNIDAD = "Unidad"
    COL_REMOLQUE = "Remolque"
    COL_CLIENTE = "Cliente"
    COL_DISTANCIA = "KMS Ruta"
    USA_OPERADOR = False
    PREFIJO_REMOLQUE = None
    TIPO_DISTANCIA = "KILOMETROS"

# ============================================================
# CARGA DE ARCHIVO
# ============================================================
archivo = st.file_uploader("📂 Carga el archivo de viajes", type=["xlsx"])

if not archivo:
    st.stop()

df = pd.read_excel(archivo)

# ============================================================
# VALIDACIONES DE COLUMNAS
# ============================================================
columnas_necesarias = [
    COL_FECHA,
    COL_UNIDAD,
    COL_REMOLQUE,
    COL_CLIENTE,
    COL_DISTANCIA,
]

if USA_OPERADOR:
    columnas_necesarias.append("Logistic Operator")

faltantes = [c for c in columnas_necesarias if c not in df.columns]

if faltantes:
    st.error(f"Faltan columnas en el archivo: {faltantes}")
    st.stop()

# ============================================================
# LIMPIEZA Y NORMALIZACIÓN
# ============================================================
df = df.copy()

df["FECHA"] = pd.to_datetime(df[COL_FECHA])
df["UNIDAD"] = df[COL_UNIDAD].astype(str).str.strip()
df["REMOLQUE"] = df[COL_REMOLQUE].astype(str).str.strip()
df["CLIENTE"] = df[COL_CLIENTE].astype(str).str.strip()
df["DISTANCIA"] = pd.to_numeric(df[COL_DISTANCIA], errors="coerce").fillna(0)

if USA_OPERADOR:
    df["OPERADOR"] = df["Logistic Operator"].astype(str).str.strip()

# ============================================================
# FILTRO REMOLQUES (SI APLICA)
# ============================================================
if PREFIJO_REMOLQUE:
    df = df[df["REMOLQUE"].str.startswith(PREFIJO_REMOLQUE)]

# ============================================================
# RESUMEN BASE
# ============================================================
st.subheader("📊 Resumen de viajes")

st.metric("Total viajes", len(df))
st.metric(f"Total {TIPO_DISTANCIA.lower()}", round(df["DISTANCIA"].sum(), 2))

st.dataframe(df.head(20), use_container_width=True)

# ============================================================
# CATÁLOGO DE COSTOS POR CLIENTE (SUPABASE)
# ============================================================
st.divider()
st.subheader("📚 Catálogo de costos por cliente")

resp = (
    supabase.table("catalogo_costos_clientes")
    .select("*")
    .eq("empresa", empresa)
    .execute()
)

catalogo_df = pd.DataFrame(resp.data)

if catalogo_df.empty:
    catalogo_df = pd.DataFrame(
        columns=["empresa", "concepto", "tipo_distribucion"]
    )

# ============================================================
# DETECCIÓN DE NUEVOS CLIENTES
# ============================================================
clientes_archivo = sorted(df["CLIENTE"].unique())
clientes_catalogo = catalogo_df["concepto"].tolist()

nuevos = [c for c in clientes_archivo if c not in clientes_catalogo]

st.write("### Clientes detectados")

tabla_clientes = pd.DataFrame({
    "empresa": empresa,
    "concepto": clientes_archivo,
    "tipo_distribucion": [
        catalogo_df.set_index("concepto").get("tipo_distribucion", {}).get(c, "")
        for c in clientes_archivo
    ]
})

def highlight_new(row):
    if row["concepto"] in nuevos:
        return ["background-color: #ffcccc"] * len(row)
    return [""] * len(row)

st.dataframe(
    tabla_clientes.style.apply(highlight_new, axis=1),
    use_container_width=True
)

# ============================================================
# GUARDAR CATÁLOGO
# ============================================================
if st.button("💾 Guardar catálogo"):
    registros = tabla_clientes.to_dict("records")

    supabase.table("catalogo_costos_clientes").upsert(
        registros,
        on_conflict="empresa,concepto"
    ).execute()

    st.success("Catálogo guardado correctamente")
