"""
Lincoln — Auditoría de Viajes
App Streamlit para auditar automáticamente el archivo de Automatización Lincoln.

Instalar dependencias:
    pip install streamlit pandas openpyxl xlsxwriter

Correr:
    streamlit run lincoln_auditoria.py
"""

import io
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Lincoln — Auditoría",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo oscuro general */
    .stApp { background-color: #0a0c10; }
    .block-container { padding-top: 1.5rem; }

    /* Sidebar */
    section[data-testid="stSidebar"] { background-color: #12151c; border-right: 1px solid #252a38; }
    section[data-testid="stSidebar"] .stRadio label { color: #94a3b8; font-size: 13px; }

    /* Métricas */
    [data-testid="stMetric"] { background: #12151c; border: 1px solid #252a38; border-radius: 10px; padding: 14px; }
    [data-testid="stMetricLabel"] { color: #64748b !important; font-size: 11px !important; text-transform: uppercase; letter-spacing: .06em; }

    /* Tablas */
    .stDataFrame { border: 1px solid #252a38 !important; border-radius: 8px; }

    /* Chips de estado inline */
    .chip-ok  { background:#10b98120; color:#10b981; border:1px solid #10b98140; border-radius:20px; padding:2px 8px; font-size:11px; font-weight:600; }
    .chip-err { background:#ef444420; color:#ef4444; border:1px solid #ef444440; border-radius:20px; padding:2px 8px; font-size:11px; font-weight:600; }
    .chip-warn{ background:#f59e0b20; color:#f59e0b; border:1px solid #f59e0b40; border-radius:20px; padding:2px 8px; font-size:11px; font-weight:600; }

    /* Separador */
    hr { border-color: #252a38; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# CONSTANTES — EQUIVALENCIAS I→C (basadas en hoja "Equivalencia I y C")
# ─────────────────────────────────────────────────────────────

# Flete USA ─ Regla 3: ingreso1 + ingreso2 deben ≈ costo
FLETE_USA_R3_ING = ["I FREIGHT USATRANSP USA39", "I FUEL CHARGES DIESEL40"]
FLETE_USA_R3_COSTO = "C FREIGHT USACT TRANSP USA77"

# Flete USA ─ columnas sin costo asociado (R1 y R2, unidad propia)
FLETE_USA_SIN_COSTO = {
    "I FREIGHT USATRANSP USA2": None,   # R1 carretera
    "I FUEL CHARGES DIESEL3": None,
    "I FREIGHT USATRANSP USA20": None,  # R2 broker + unidad
    "I FUEL CHARGES DIESEL21": None,
    "I FREIGHT USATRANSP USA56": None,  # sin identificar
}

# Flete MX ─ pares ingreso→costo por regla
FLETE_MEX_PARES = {
    "I FREIGHT MEXTRANSP MEX19": "C FREIGHT MEXCT TRANSP MEX71",   # R2
    "I FREIGHT MEXTRANSP MEX38": "C FREIGHT MEXCT TRANSP MEX76",   # R3 (y R2 cuando no hay tracto)
    "I FREIGHT MEXTRANSP MEX61": "C FREIGHT MEXCT TRANSP MEX84",   # sin identificar
}
# MEX1 sin costo (R1, nunca hay datos)
FLETE_MEX_SIN_COSTO = ["I FREIGHT MEXTRANSP MEX1"]

# Cruce ─ pares ingreso→costo
CRUCE_PARES = {
    "I CROSS BORDER EMPTYCRUCE VACIO6":       "C CROSS BORDER LOADEDCT CRUCE CARGADO66",
    "I CROSS BORDER LOADEDCRUCE CARGADO7":    "C CROSS BORDER LOADEDCT CRUCE CARGADO66",
    "I CROSS BORDER EMPTYCRUCE VACIO24":      "C CROSS BORDER LOADEDCT CRUCE CARGADO68",
    "I CROSS BORDER LOADEDCRUCE CARGADO25":   "C CROSS BORDER LOADEDCT CRUCE CARGADO68",
    "I CROSS BORDER EMPTYCRUCE VACIO43":      "C CROSS BORDER LOADEDCT CRUCE CARGADO73",
    "I CROSS BORDER LOADEDCRUCE CARGADO44":   "C CROSS BORDER LOADEDCT CRUCE CARGADO73",
}

# Extra Stop ─ pares ingreso→costo (R1 sin costo)
EXTRA_STOP_PARES = {
    "I EXTRA STOPPARADA EXTRA5":  None,                         # R1 sin costo
    "I EXTRA STOPPARADA EXTRA23": "C EXTRA STOPCT PARADA EXTRA70",  # R2
    "I EXTRA STOPPARADA EXTRA42": "C EXTRA STOPCT PARADA EXTRA75",  # R3
}

# TNU ─ pares ingreso→costo (R1 y R2 sin costo)
TNU_PARES = {
    "I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO14": None,
    "I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO32": None,
    "I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO51": "C TNU - TRUCK NOT USEDCT MOVIMIENTO EN FALSO90",
}

# Handling ─ pares ingreso→costo
HANDLING_PARES = {
    "I HANDLING CHARGESMANIOBRAS13": None,
    "I HANDLING CHARGESMANIOBRAS31": None,
    "I HANDLING CHARGESMANIOBRAS50": "C HANDLING CHARGESCT MANIOBRAS89",
}

# Umbrales de variación permitida
UMBRAL = {
    "flete_usa": 200,
    "flete_mex": 200,
    "cruce":     200,
    "extra_stop": 50,
    "tnu":        50,
    "handling":   50,
}


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def v(row, col):
    """Valor numérico seguro de una columna, 0 si no existe o es NaN."""
    return float(row.get(col, 0) or 0)


def get_regla(row):
    """Determina la regla de servicio según Servicio y Número Tracto."""
    serv  = str(row.get("Servicio", "")).strip().upper()
    tracto = str(row.get("Número Tracto", "")).strip()
    if "CARRETERA" in serv and tracto:
        return 1
    if "BROKER" in serv and tracto:
        return 2
    if "BROKER" in serv and not tracto:
        return 3
    return 0


def fmt_usd(val):
    if val is None:
        return "—"
    return f"${abs(val):,.2f}"


def estado_chip(ok):
    if ok:
        return "✅ OK"
    return "❌ Anomalía"


# ─────────────────────────────────────────────────────────────
# LÓGICAS DE AUDITORÍA
# ─────────────────────────────────────────────────────────────
def audit_flete_usa(row):
    """
    Devuelve dict con resultados de auditoría de Flete USA para una fila.
    Regla 3: I FREIGHT USA39 + I FUEL40 ≈ C FREIGHT USA77 (±$200).
    Reglas 1 y 2: no debe haber costo en columnas de costo USA.
    """
    regla  = get_regla(row)
    viaje  = row.get("Número De Viaje", "")
    tracto = row.get("Número Tracto", "")
    tipo   = row.get("Tipo Viaje", "")
    ok     = True
    obs    = []

    costos_usa = [
        "C FREIGHT USACT TRANSP USA72",
        "C FREIGHT USACT TRANSP USA77",
        "C FREIGHT USACT TRANSP USA78",
    ]

    if regla == 1:
        i_flete = v(row, "I FREIGHT USATRANSP USA2")
        i_fuel  = v(row, "I FUEL CHARGES DIESEL3")
        costo   = sum(v(row, c) for c in costos_usa)
        i_total = i_flete + i_fuel
        if i_total == 0:
            return None  # sin datos
        if costo > 0:
            ok = False
            obs.append(f"R1: No debe haber costo en flete USA para unidad propia (costo={fmt_usd(costo)}).")

    elif regla == 2:
        i_flete = v(row, "I FREIGHT USATRANSP USA20")
        i_fuel  = v(row, "I FUEL CHARGES DIESEL21")
        costo   = sum(v(row, c) for c in costos_usa)
        i_total = i_flete + i_fuel
        if i_total == 0:
            return None
        if costo > 0:
            ok = False
            obs.append(f"R2: No debe haber costo en flete USA cuando hay unidad capturada (costo={fmt_usd(costo)}).")

    elif regla == 3:
        i_flete = v(row, FLETE_USA_R3_ING[0])
        i_fuel  = v(row, FLETE_USA_R3_ING[1])
        i_total = i_flete + i_fuel
        costo   = v(row, FLETE_USA_R3_COSTO)
        if i_total == 0 and costo == 0:
            return None
        if i_total > 0 and costo == 0:
            ok = False
            obs.append(f"R3: Hay ingreso ({fmt_usd(i_total)}) pero sin costo — tercero debe tener costo.")
        elif i_total == 0 and costo > 0:
            ok = False
            obs.append(f"R3: Hay costo ({fmt_usd(costo)}) pero sin ingreso correspondiente.")
        else:
            diff = abs(i_total - costo)
            if diff > UMBRAL["flete_usa"]:
                ok = False
                obs.append(
                    f"R3: Variación de {fmt_usd(diff)} excede ${UMBRAL['flete_usa']} "
                    f"(I_flete={fmt_usd(i_flete)} + I_fuel={fmt_usd(i_fuel)} = {fmt_usd(i_total)}, "
                    f"C={fmt_usd(costo)})."
                )
    else:
        return None

    return {
        "Número Viaje": viaje,
        "Tracto": tracto,
        "Tipo Viaje": tipo,
        "Regla": f"R{regla}",
        "I Flete": i_flete if regla != 1 else v(row, "I FREIGHT USATRANSP USA2"),
        "I Fuel": i_fuel if regla != 1 else v(row, "I FUEL CHARGES DIESEL3"),
        "I Total": i_total,
        "Costo": costo,
        "Diferencia": i_total - costo,
        "Estado": estado_chip(ok),
        "OK": ok,
        "Observación": " / ".join(obs) if obs else "",
    }


def audit_flete_mex(row):
    """
    Audita todos los pares de flete MX según equivalencias.
    Todo ingreso MX debe tener su costo (siempre lo hace un tercero).
    """
    viaje  = row.get("Número De Viaje", "")
    tracto = row.get("Número Tracto", "")
    tipo   = row.get("Tipo Viaje", "")
    regla  = get_regla(row)
    resultados = []

    for col_i, col_c in FLETE_MEX_PARES.items():
        i_val = v(row, col_i)
        c_val = v(row, col_c) if col_c else 0
        if i_val == 0 and c_val == 0:
            continue
        ok  = True
        obs = []
        if i_val > 0 and c_val == 0:
            ok = False
            obs.append(f"Ingreso MX ({fmt_usd(i_val)}) sin costo — siempre lo hace un tercero.")
        elif i_val == 0 and c_val > 0:
            ok = False
            obs.append(f"Costo MX ({fmt_usd(c_val)}) sin ingreso correspondiente.")
        else:
            diff = abs(i_val - c_val)
            if diff > UMBRAL["flete_mex"]:
                ok = False
                obs.append(
                    f"Variación {fmt_usd(diff)} excede ${UMBRAL['flete_mex']} "
                    f"(I={fmt_usd(i_val)}, C={fmt_usd(c_val)})."
                )
        resultados.append({
            "Número Viaje": viaje,
            "Tracto": tracto,
            "Tipo Viaje": tipo,
            "Regla": f"R{regla}",
            "Col Ingreso": col_i,
            "Col Costo": col_c or "—",
            "Ingreso": i_val,
            "Costo": c_val,
            "Diferencia": i_val - c_val,
            "Estado": estado_chip(ok),
            "OK": ok,
            "Observación": " / ".join(obs),
        })
    return resultados


def audit_cruce(row):
    """
    Audita todos los pares de cruce usando la hoja de equivalencias.
    Los costos de cruce rondan $100–$200; mayor a $200 es anomalía.
    """
    viaje  = row.get("Número De Viaje", "")
    tracto = row.get("Número Tracto", "")
    tipo   = row.get("Tipo Viaje", "")
    regla  = get_regla(row)
    resultados = []

    # Agrupar: varios ingresos comparten la misma columna de costo
    grupos = {}
    for col_i, col_c in CRUCE_PARES.items():
        i_val = v(row, col_i)
        if i_val == 0:
            continue
        if col_c not in grupos:
            grupos[col_c] = {"i_total": 0, "cols_i": [], "col_c": col_c}
        grupos[col_c]["i_total"] += i_val
        grupos[col_c]["cols_i"].append(col_i)

    for col_c, g in grupos.items():
        i_total = g["i_total"]
        c_val   = v(row, col_c)
        ok  = True
        obs = []

        if i_total > 0 and c_val == 0:
            # R1/R2 con unidad propia no necesitan costo en cruce
            # Solo R3 (sin tracto) exige costo
            if regla == 3:
                ok = False
                obs.append(f"Cruce: ingreso ({fmt_usd(i_total)}) sin costo — tercero debe tener costo.")
        elif i_total == 0 and c_val > 0:
            ok = False
            obs.append(f"Cruce: costo ({fmt_usd(c_val)}) sin ingreso.")
        elif i_total > 0 and c_val > 0:
            diff = abs(i_total - c_val)
            if diff > UMBRAL["cruce"]:
                ok = False
                obs.append(
                    f"Variación {fmt_usd(diff)} excede ${UMBRAL['cruce']} "
                    f"(I={fmt_usd(i_total)}, C={fmt_usd(c_val)})."
                )
            if c_val > 400:
                ok = False
                obs.append(f"Costo de cruce ({fmt_usd(c_val)}) fuera del rango de mercado ($100–$200).")

        if i_total > 0 or c_val > 0:
            resultados.append({
                "Número Viaje": viaje,
                "Tracto": tracto,
                "Tipo Viaje": tipo,
                "Regla": f"R{regla}",
                "Col Costo": col_c,
                "I Cruce Total": i_total,
                "Costo": c_val,
                "Diferencia": i_total - c_val,
                "Estado": estado_chip(ok),
                "OK": ok,
                "Observación": " / ".join(obs),
            })
    return resultados


def _audit_simple(row, pares, umbral, nombre):
    """
    Función genérica para Extra Stop, TNU y Handling.
    Para cada par ingreso→costo aplica las reglas estándar.
    """
    viaje  = row.get("Número De Viaje", "")
    tracto = row.get("Número Tracto", "")
    tipo   = row.get("Tipo Viaje", "")
    regla  = get_regla(row)
    resultados = []

    for col_i, col_c in pares.items():
        i_val = v(row, col_i)
        c_val = v(row, col_c) if col_c else 0
        if i_val == 0 and c_val == 0:
            continue
        ok  = True
        obs = []

        if col_c is None:
            # Sin costo asociado — solo validar que no haya valores extremos
            if nombre == "Extra Stop" and i_val > 300:
                ok = False
                obs.append(f"{nombre}: ingreso ({fmt_usd(i_val)}) parece elevado (>$300).")
            if nombre == "Handling" and i_val > 1500:
                ok = False
                obs.append(f"{nombre}: ingreso ({fmt_usd(i_val)}) parece elevado (>$1,500).")
        else:
            if i_val > 0 and c_val == 0:
                ok = False
                obs.append(f"{nombre}: ingreso ({fmt_usd(i_val)}) sin costo — tercero debe tener costo.")
            elif i_val == 0 and c_val > 0:
                ok = False
                obs.append(f"{nombre}: costo ({fmt_usd(c_val)}) sin ingreso — revisar.")
            else:
                diff = abs(i_val - c_val)
                if diff > umbral:
                    ok = False
                    obs.append(
                        f"{nombre}: variación {fmt_usd(diff)} excede ${umbral} "
                        f"(I={fmt_usd(i_val)}, C={fmt_usd(c_val)})."
                    )

        resultados.append({
            "Número Viaje": viaje,
            "Tracto": tracto,
            "Tipo Viaje": tipo,
            "Regla": f"R{regla}",
            "Col Ingreso": col_i,
            "Col Costo": col_c or "—",
            "Ingreso": i_val,
            "Costo": c_val,
            "Diferencia": i_val - c_val,
            "Estado": estado_chip(ok),
            "OK": ok,
            "Observación": " / ".join(obs),
        })
    return resultados


def audit_extra_stop(row):
    return _audit_simple(row, EXTRA_STOP_PARES, UMBRAL["extra_stop"], "Extra Stop")


def audit_tnu(row):
    return _audit_simple(row, TNU_PARES, UMBRAL["tnu"], "TNU")


def audit_handling(row):
    return _audit_simple(row, HANDLING_PARES, UMBRAL["handling"], "Handling")


# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO PRINCIPAL
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def procesar_archivo(file_bytes):
    """Lee el Excel y ejecuta todas las auditorías. Cacheable."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Companies")
    df = df.fillna(0)

    # Normalizar columna clave
    for col in ["Servicio", "Número Tracto", "Número De Viaje", "Tipo Viaje",
                "Estatus", "Cliente", "Importe Ingreso", "Importe Costo",
                "Importe Utilidad", "% Utilidad"]:
        if col not in df.columns:
            df[col] = ""

    df["Número Tracto"] = df["Número Tracto"].apply(
        lambda x: "" if str(x).strip() in ["0", "nan", "0.0"] else str(x).strip()
    )

    resultados = {
        "flete_usa":  [],
        "flete_mex":  [],
        "cruce":      [],
        "extra_stop": [],
        "tnu":        [],
        "handling":   [],
        "cancelados": [],
        "ut_up":      [],
    }

    for _, row in df.iterrows():
        estatus = str(row.get("Estatus", "")).upper()
        if "CANCEL" in estatus:
            resultados["cancelados"].append(row.to_dict())
            continue

        r = audit_flete_usa(row)
        if r:
            resultados["flete_usa"].append(r)

        resultados["flete_mex"].extend(audit_flete_mex(row))
        resultados["cruce"].extend(audit_cruce(row))
        resultados["extra_stop"].extend(audit_extra_stop(row))
        resultados["tnu"].extend(audit_tnu(row))
        resultados["handling"].extend(audit_handling(row))

        # Utilidades
        regla = get_regla(row)
        ingreso  = float(row.get("Importe Ingreso", 0) or 0)
        costo    = float(row.get("Importe Costo", 0) or 0)
        utilidad = float(row.get("Importe Utilidad", 0) or 0)
        pct      = float(row.get("% Utilidad", 0) or 0)
        tracto   = str(row.get("Número Tracto", "")).strip()

        umbral_ut = 0.40 if tracto else 0.20
        alerta_ut = pct < umbral_ut and ingreso > 0

        resultados["ut_up"].append({
            "Número Viaje": row.get("Número De Viaje", ""),
            "Tracto": tracto,
            "Servicio": row.get("Servicio", ""),
            "Cliente": row.get("Cliente", ""),
            "Ingreso": ingreso,
            "Costo": costo,
            "Utilidad": utilidad,
            "% Utilidad": pct,
            "Umbral": umbral_ut,
            "Alerta UT": "⚠️ Baja" if alerta_ut else "✅ OK",
            "OK": not alerta_ut,
        })

    # Convertir a DataFrames
    dfs = {}
    for key, data in resultados.items():
        dfs[key] = pd.DataFrame(data) if data else pd.DataFrame()

    # Estadísticas rápidas
    total = len(df)
    cancelados = len(dfs["cancelados"])
    with_anomaly = set()
    for key in ["flete_usa", "flete_mex", "cruce", "extra_stop", "tnu", "handling"]:
        if not dfs[key].empty and "OK" in dfs[key].columns:
            bad = dfs[key][dfs[key]["OK"] == False]["Número Viaje"].unique()
            with_anomaly.update(bad)

    stats = {
        "total": total,
        "cancelados": cancelados,
        "anomalias": len(with_anomaly),
        "ok": total - cancelados - len(with_anomaly),
    }
    return dfs, stats, with_anomaly


def to_excel_bytes(dfs_dict):
    """Exporta todas las secciones a un Excel con múltiples hojas."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True, "bg_color": "#12151c", "font_color": "#94a3b8",
            "border": 1, "border_color": "#252a38",
        })
        fmt_ok   = wb.add_format({"font_color": "#10b981"})
        fmt_err  = wb.add_format({"font_color": "#ef4444"})
        fmt_warn = wb.add_format({"font_color": "#f59e0b"})

        labels = {
            "flete_usa": "Flete USA",
            "flete_mex": "Flete MX",
            "cruce": "Cruce",
            "extra_stop": "Extra Stop",
            "tnu": "TNU",
            "handling": "Handling",
            "cancelados": "Cancelados",
            "ut_up": "Utilidades",
        }
        for key, label in labels.items():
            df = dfs_dict.get(key, pd.DataFrame())
            if df.empty:
                continue
            # Quitar columna interna OK al exportar
            cols = [c for c in df.columns if c != "OK"]
            df[cols].to_excel(writer, sheet_name=label, index=False)
            ws = writer.sheets[label]
            ws.set_row(0, None, fmt_header)
            ws.set_column(0, 0, 18)
            ws.set_column(1, 20, 14)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# COMPONENTES UI
# ─────────────────────────────────────────────────────────────
def mostrar_tabla(df, filtro_estado=None, key_search=None):
    """Muestra dataframe con opción de filtro y búsqueda."""
    if df.empty:
        st.info("Sin registros en esta sección.")
        return

    df_show = df.copy()

    # Filtro de búsqueda
    if key_search:
        q = st.text_input("🔎 Buscar número de viaje", key=key_search)
        if q:
            df_show = df_show[
                df_show["Número Viaje"].astype(str).str.contains(q, case=False)
            ]

    # Filtro de estado
    if filtro_estado and filtro_estado != "Todos" and "OK" in df_show.columns:
        if filtro_estado == "Con anomalía":
            df_show = df_show[df_show["OK"] == False]
        elif filtro_estado == "Sin anomalía":
            df_show = df_show[df_show["OK"] == True]

    # Ocultar columna interna OK
    cols = [c for c in df_show.columns if c != "OK"]
    df_show = df_show[cols]

    st.dataframe(
        df_show,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Diferencia": st.column_config.NumberColumn(format="$%.2f"),
            "Ingreso":    st.column_config.NumberColumn(format="$%.2f"),
            "Costo":      st.column_config.NumberColumn(format="$%.2f"),
            "I Total":    st.column_config.NumberColumn(format="$%.2f"),
            "I Flete":    st.column_config.NumberColumn(format="$%.2f"),
            "I Fuel":     st.column_config.NumberColumn(format="$%.2f"),
            "I Cruce Total": st.column_config.NumberColumn(format="$%.2f"),
            "% Utilidad": st.column_config.NumberColumn(format="%.1%"),
            "Umbral":     st.column_config.NumberColumn(format="%.0%"),
            "Utilidad":   st.column_config.NumberColumn(format="$%.2f"),
        }
    )

    # Contador
    n_err = (df["OK"] == False).sum() if "OK" in df.columns else 0
    st.caption(f"{len(df_show)} registros mostrados | {n_err} anomalías en total")


def seccion_auditoria(titulo, icono, df, umbral_info, key_prefix):
    st.subheader(f"{icono} {titulo}")
    if not df.empty and "OK" in df.columns:
        n_ok  = (df["OK"] == True).sum()
        n_err = (df["OK"] == False).sum()
        col1, col2, col3 = st.columns(3)
        col1.metric("Total registros", len(df))
        col2.metric("✅ Sin anomalía", n_ok)
        col3.metric("❌ Con anomalía", n_err)

    filtro = st.radio(
        "Mostrar",
        ["Todos", "Con anomalía", "Sin anomalía"],
        horizontal=True,
        key=f"{key_prefix}_radio",
    )
    st.caption(umbral_info)
    mostrar_tabla(df, filtro_estado=filtro, key_search=f"{key_prefix}_search")
    st.markdown("---")


# ─────────────────────────────────────────────────────────────
# SIDEBAR Y NAVEGACIÓN
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Lincoln Auditoría")
    st.markdown("---")

    archivo = st.file_uploader(
        "Cargar Automatización_lincoln.xlsx",
        type=["xlsx", "xls"],
        help="Sube el archivo de Automatización Lincoln. Debe contener la hoja 'Companies'.",
    )

    st.markdown("---")

    pagina = st.radio(
        "Sección",
        [
            "📊 Resumen",
            "🚛 Flete USA",
            "🇲🇽 Flete México",
            "🌉 Cruce",
            "📍 Extra Stop",
            "🚫 TNU",
            "📦 Handling",
            "💰 Utilidades",
            "❌ Cancelados",
            "📋 Reglas de auditoría",
        ],
    )

    st.markdown("---")
    st.markdown(
        "<small style='color:#64748b'>Umbrales: Flete USA/MX/Cruce ±$200 · "
        "Extra Stop/TNU/Handling ±$50</small>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────
# CUERPO PRINCIPAL
# ─────────────────────────────────────────────────────────────
if archivo is None:
    st.title("🔍 Lincoln — Auditoría de Viajes")
    st.markdown(
        "Carga tu archivo **Automatización_lincoln.xlsx** desde el panel izquierdo "
        "para comenzar la auditoría automática."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**📥 1. Carga el Excel**\n\nSube el archivo desde la barra lateral. "
                "No necesitas prepararlo previamente.")
    with col2:
        st.info("**🔍 2. Auditoría automática**\n\nEl sistema aplica las 3 reglas de servicio "
                "y detecta anomalías en todos los conceptos.")
    with col3:
        st.info("**📋 3. Revisa y exporta**\n\nFiltra por sección, revisa las observaciones "
                "y exporta el reporte completo a Excel.")

    st.markdown("---")
    st.markdown("#### 📋 Reglas de servicio")
    col1, col2, col3 = st.columns(3)
    col1.success("**Regla 1** — Carretera USA + Unidad capturada\n\nUnidad propia. Sin costo de flete USA.")
    col2.info("**Regla 2** — Broker USA + Unidad capturada\n\nSin costo de flete USA. Flete MX siempre tiene costo.")
    col3.warning("**Regla 3** — Broker USA + Sin unidad\n\nTercero completo. Flete USA39 + Fuel40 ≈ Costo77.")
    st.stop()


# ── Procesar ──────────────────────────────────────────────────
with st.spinner("Procesando archivo y aplicando reglas de auditoría..."):
    file_bytes = archivo.read()
    try:
        dfs, stats, viajes_anomalia = procesar_archivo(file_bytes)
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
        st.stop()

# ── Botón de exportar (siempre visible arriba) ─────────────────
col_exp, _ = st.columns([2, 8])
with col_exp:
    xlsx_bytes = to_excel_bytes(dfs)
    st.download_button(
        label="⬇️ Exportar auditoría",
        data=xlsx_bytes,
        file_name="Auditoria_Lincoln.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─────────────────────────────────────────────────────────────
# PÁGINAS
# ─────────────────────────────────────────────────────────────

if pagina == "📊 Resumen":
    st.title("📊 Resumen de Auditoría")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total viajes",      stats["total"])
    col2.metric("✅ Sin anomalía",   stats["ok"])
    col3.metric("❌ Con anomalía",   stats["anomalias"])
    col4.metric("🚫 Cancelados",     stats["cancelados"])

    st.markdown("---")
    st.subheader("⚠️ Viajes con anomalías detectadas")

    if not viajes_anomalia:
        st.success("¡Sin anomalías! Todos los viajes cumplen las reglas de auditoría.")
    else:
        # Construir reporte consolidado
        rows_reporte = []
        for key, nombre in [
            ("flete_usa", "Flete USA"),
            ("flete_mex", "Flete MX"),
            ("cruce", "Cruce"),
            ("extra_stop", "Extra Stop"),
            ("tnu", "TNU"),
            ("handling", "Handling"),
        ]:
            df_k = dfs[key]
            if df_k.empty or "OK" not in df_k.columns:
                continue
            for _, r in df_k[df_k["OK"] == False].iterrows():
                rows_reporte.append({
                    "Número Viaje": r["Número Viaje"],
                    "Concepto": nombre,
                    "Regla": r.get("Regla", ""),
                    "Observación": r.get("Observación", ""),
                })
        df_rep = pd.DataFrame(rows_reporte)
        st.dataframe(df_rep, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📈 Anomalías por concepto")
    resumen_cnt = []
    for key, nombre in [
        ("flete_usa", "Flete USA"),
        ("flete_mex", "Flete MX"),
        ("cruce", "Cruce"),
        ("extra_stop", "Extra Stop"),
        ("tnu", "TNU"),
        ("handling", "Handling"),
    ]:
        df_k = dfs[key]
        if df_k.empty or "OK" not in df_k.columns:
            total_k = 0; err_k = 0
        else:
            total_k = len(df_k)
            err_k = (df_k["OK"] == False).sum()
        resumen_cnt.append({
            "Concepto": nombre,
            "Registros auditados": total_k,
            "Anomalías": err_k,
            "OK": total_k - err_k,
        })
    st.dataframe(pd.DataFrame(resumen_cnt), use_container_width=True, hide_index=True)


elif pagina == "🚛 Flete USA":
    seccion_auditoria(
        "Flete USA", "🚛", dfs["flete_usa"],
        "R1/R2 (unidad propia): sin costo. R3 (tercero): I_flete39 + I_fuel40 ≈ C77 · variación máxima ±$200",
        "fusa",
    )


elif pagina == "🇲🇽 Flete México":
    seccion_auditoria(
        "Flete México", "🇲🇽", dfs["flete_mex"],
        "Todo ingreso MX debe tener costo (siempre tercero). Variación máxima ±$200",
        "fmex",
    )


elif pagina == "🌉 Cruce":
    seccion_auditoria(
        "Cruce", "🌉", dfs["cruce"],
        "Cruce de mercado: $100–$200. R3 (tercero): ingreso ≈ costo. Variación máxima ±$200",
        "cruce",
    )


elif pagina == "📍 Extra Stop":
    seccion_auditoria(
        "Extra Stop", "📍", dfs["extra_stop"],
        "Variación máxima ±$50. Costo sin ingreso = anomalía importante. Ingreso sin costo = verificar.",
        "exstop",
    )


elif pagina == "🚫 TNU":
    seccion_auditoria(
        "Truck Not Used", "🚫", dfs["tnu"],
        "R3: ingreso51 debe tener costo90. Variación máxima ±$50.",
        "tnu",
    )


elif pagina == "📦 Handling":
    seccion_auditoria(
        "Handling (Maniobras)", "📦", dfs["handling"],
        "R3: ingreso50 debe tener costo89. Ingreso >$1,500 = alerta. Variación máxima ±$50.",
        "handling",
    )


elif pagina == "💰 Utilidades":
    st.subheader("💰 Utilidades por viaje")
    df_ut = dfs["ut_up"]
    if df_ut.empty:
        st.info("Sin datos de utilidad.")
    else:
        st.caption(
            "Unidad propia: alerta si % utilidad < 40% · Sin unidad (tercero): alerta si < 20%"
        )
        filtro_ut = st.radio(
            "Mostrar", ["Todos", "⚠️ Alerta", "✅ OK"], horizontal=True, key="ut_radio"
        )
        df_show = df_ut.copy()
        if filtro_ut == "⚠️ Alerta":
            df_show = df_show[df_show["OK"] == False]
        elif filtro_ut == "✅ OK":
            df_show = df_show[df_show["OK"] == True]

        cols = [c for c in df_show.columns if c != "OK"]
        st.dataframe(
            df_show[cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Ingreso":   st.column_config.NumberColumn(format="$%.2f"),
                "Costo":     st.column_config.NumberColumn(format="$%.2f"),
                "Utilidad":  st.column_config.NumberColumn(format="$%.2f"),
                "% Utilidad": st.column_config.NumberColumn(format="%.1%"),
                "Umbral":    st.column_config.NumberColumn(format="%.0%"),
            }
        )
        n_alerta = (df_ut["OK"] == False).sum()
        st.caption(f"{len(df_show)} registros · {n_alerta} con alerta de utilidad baja")


elif pagina == "❌ Cancelados":
    st.subheader("❌ Viajes Cancelados")
    df_c = dfs["cancelados"]
    if df_c.empty:
        st.info("No hay viajes cancelados.")
    else:
        cols_mostrar = [
            "Número De Viaje", "Estatus", "Cliente", "Servicio",
            "Importe Ingreso", "Importe Costo", "Importe Utilidad",
        ]
        cols_disp = [c for c in cols_mostrar if c in df_c.columns]
        st.dataframe(df_c[cols_disp], use_container_width=True, hide_index=True)
        st.caption(f"{len(df_c)} viajes cancelados")


elif pagina == "📋 Reglas de auditoría":
    st.subheader("📋 Reglas de Auditoría Lincoln")

    st.markdown("#### Reglas de Servicio")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.success(
            "**Regla 1 — Carretera USA + Unidad capturada**\n\n"
            "- Unidad propia hace toda la operación\n"
            "- Ingreso flete USA → col ...USA2\n"
            "- Ingreso fuel → col ...DIESEL3\n"
            "- **Sin costo** de flete USA\n"
            "- Si hay costo → anomalía"
        )
    with col2:
        st.info(
            "**Regla 2 — Broker USA + Unidad capturada**\n\n"
            "- Ingreso flete USA → col ...USA20\n"
            "- Ingreso fuel → col ...DIESEL21\n"
            "- **Sin costo** de flete USA\n"
            "- Flete MX: I19 → C71 (siempre tercero)"
        )
    with col3:
        st.warning(
            "**Regla 3 — Broker USA + Sin unidad**\n\n"
            "- Tercero completo\n"
            "- Ingreso USA → col ...USA39\n"
            "- Fuel → col ...DIESEL40\n"
            "- **I39 + Fuel40 ≈ C77** (±$200)\n"
            "- Flete MX: I38 → C76"
        )

    st.markdown("---")
    st.markdown("#### Equivalencias Ingreso → Costo")
    equiv_data = [
        ("I FREIGHT USATRANSP USA2",              "—",                                  "R1, sin costo"),
        ("I FUEL CHARGES DIESEL3",                "—",                                  "R1, sin costo"),
        ("I FREIGHT USATRANSP USA20",             "—",                                  "R2, sin costo"),
        ("I FUEL CHARGES DIESEL21",               "—",                                  "R2, sin costo"),
        ("I FREIGHT USATRANSP USA39 + DIESEL40",  "C FREIGHT USACT TRANSP USA77",       "R3, suma flete+fuel"),
        ("I FREIGHT MEXTRANSP MEX19",             "C FREIGHT MEXCT TRANSP MEX71",       "R2"),
        ("I FREIGHT MEXTRANSP MEX38",             "C FREIGHT MEXCT TRANSP MEX76",       "R3"),
        ("I FREIGHT MEXTRANSP MEX61",             "C FREIGHT MEXCT TRANSP MEX84",       "Sin identificar"),
        ("I CROSS BORDER EMPTY/LOADED 6 y 7",     "C CROSS BORDER LOADED 66",           "R1"),
        ("I CROSS BORDER EMPTY/LOADED 24 y 25",   "C CROSS BORDER LOADED 68",           "R2"),
        ("I CROSS BORDER EMPTY/LOADED 43 y 44",   "C CROSS BORDER LOADED 73",           "R3"),
        ("I EXTRA STOPPARADA EXTRA5",             "—",                                  "R1, sin costo"),
        ("I EXTRA STOPPARADA EXTRA23",            "C EXTRA STOPCT PARADA EXTRA70",      "R2, ±$50"),
        ("I EXTRA STOPPARADA EXTRA42",            "C EXTRA STOPCT PARADA EXTRA75",      "R3, ±$50"),
        ("I TNU FALSO14",                         "—",                                  "R1, sin costo"),
        ("I TNU FALSO32",                         "—",                                  "R2, sin costo"),
        ("I TNU FALSO51",                         "C TNU MOVIMIENTO EN FALSO90",        "R3, ±$50"),
        ("I HANDLING MANIOBRAS13",                "—",                                  "R1, sin costo"),
        ("I HANDLING MANIOBRAS31",                "—",                                  "R2, sin costo"),
        ("I HANDLING MANIOBRAS50",                "C HANDLING MANIOBRAS89",             "R3, ±$50"),
    ]
    df_eq = pd.DataFrame(equiv_data, columns=["Columna Ingreso", "Columna Costo", "Nota"])
    st.dataframe(df_eq, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### Umbrales de Variación")
    umb_data = [
        ("Flete USA (R3)",  "$200", "I_flete39 + I_fuel40 vs C77"),
        ("Flete México",    "$200", "Ingreso vs Costo por par"),
        ("Cruce",           "$200", "Mercado: $100–$200. >$400 en costo = anomalía"),
        ("Extra Stop",      "$50",  "Costo sin ingreso = prioridad alta"),
        ("TNU",             "$50",  "R3: debe haber costo cuando hay ingreso"),
        ("Handling",        "$50",  "Ingreso >$1,500 = alerta adicional"),
    ]
    st.dataframe(
        pd.DataFrame(umb_data, columns=["Concepto", "Variación máx.", "Nota"]),
        use_container_width=True,
        hide_index=True,
    )
