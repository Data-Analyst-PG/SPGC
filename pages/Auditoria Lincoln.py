from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Lincoln Auditoría", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
def n(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def s(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def money(value: float) -> str:
    return f"${value:,.2f}"


def get_col(row: pd.Series, col: str) -> float:
    return n(row[col]) if col in row.index else 0.0


def get_text(row: pd.Series, col: str) -> str:
    return s(row[col]) if col in row.index else ""


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in df.columns:
        if isinstance(c, str):
            df.rename(columns={c: c.strip()}, inplace=True)
    return df


def get_regla(row: pd.Series) -> int:
    servicio = get_text(row, "Servicio").upper()
    tracto = get_text(row, "Número Tracto")
    if "CARRETERA" in servicio and tracto:
        return 1
    if "BROKER" in servicio and tracto:
        return 2
    if "BROKER" in servicio and not tracto:
        return 3
    return 0


def route_of(row: pd.Series) -> str:
    parts = [
        get_text(row, "Ciudad Origen"),
        get_text(row, "Estado Origen"),
        get_text(row, "Ciudad Destino"),
        get_text(row, "Estado Destino"),
    ]
    parts = [p for p in parts if p]
    return " | ".join(parts)


@dataclass
class AuditRow:
    auditoria: str
    viaje: str
    tracto: str
    tipo_viaje: str
    servicio: str
    regla: int
    estado: str
    observacion: str
    ruta: str
    metricas: Dict[str, float]

    def as_dict(self) -> Dict[str, object]:
        return {
            "Auditoria": self.auditoria,
            "Número Viaje": self.viaje,
            "Tracto": self.tracto,
            "Tipo Viaje": self.tipo_viaje,
            "Servicio": self.servicio,
            "Regla": f"R{self.regla}" if self.regla else "R0",
            "Estado": self.estado,
            "Observación": self.observacion,
            "Ruta": self.ruta,
            **self.metricas,
        }


# -----------------------------
# Auditorías
# -----------------------------
def audit_flete_usa(row: pd.Series) -> AuditRow | None:
    regla = get_regla(row)
    viaje = get_text(row, "Número De Viaje")
    tracto = get_text(row, "Número Tracto")
    tipo_viaje = get_text(row, "Tipo Viaje")
    servicio = get_text(row, "Servicio")
    ruta = route_of(row)

    i_flete = i_fuel = costo = 0.0
    estado = "OK"
    obs = ""

    if regla == 1:
        i_flete = get_col(row, "I FREIGHT USATRANSP USA2")
        i_fuel = get_col(row, "I FUEL CHARGES DIESEL3")
        costo = (
            get_col(row, "C FREIGHT USACT TRANSP USA72")
            + get_col(row, "C FREIGHT USACT TRANSP USA77")
            + get_col(row, "C FREIGHT USACT TRANSP USA78")
        )
        if costo > 0:
            estado = "Anomalía"
            obs = "Regla 1: unidad propia con costo de flete USA."
    elif regla == 2:
        i_flete = get_col(row, "I FREIGHT USATRANSP USA20")
        i_fuel = get_col(row, "I FUEL CHARGES DIESEL21")
        costo = (
            get_col(row, "C FREIGHT USACT TRANSP USA72")
            + get_col(row, "C FREIGHT USACT TRANSP USA77")
            + get_col(row, "C FREIGHT USACT TRANSP USA78")
        )
        if costo > 0:
            estado = "Anomalía"
            obs = "Regla 2: broker USA con tracto no debería traer costo USA en esas columnas."
    elif regla == 3:
        i_flete = get_col(row, "I FREIGHT USATRANSP USA39")
        i_fuel = get_col(row, "I FUEL CHARGES DIESEL40")
        costo = get_col(row, "C FREIGHT USACT TRANSP USA77")
        total_ing = i_flete + i_fuel
        if total_ing > 0 and costo == 0:
            estado = "Anomalía"
            obs = "Regla 3: hay ingreso de flete USA pero no costo."
        elif total_ing == 0 and costo > 0:
            estado = "Anomalía"
            obs = "Regla 3: hay costo de flete USA pero no ingreso."
        elif total_ing > 0 and abs(total_ing - costo) > 200:
            estado = "Anomalía"
            obs = f"Regla 3: diferencia {money(abs(total_ing - costo))} mayor a $200."
    else:
        return None

    if (i_flete + i_fuel + costo) == 0:
        return None

    return AuditRow(
        auditoria="Flete USA",
        viaje=viaje,
        tracto=tracto,
        tipo_viaje=tipo_viaje,
        servicio=servicio,
        regla=regla,
        estado=estado,
        observacion=obs,
        ruta=ruta,
        metricas={
            "I Flete USA": i_flete,
            "I Fuel": i_fuel,
            "C Flete USA": costo,
            "Diferencia": (i_flete + i_fuel) - costo,
        },
    )


def audit_flete_mex(row: pd.Series) -> AuditRow | None:
    regla = get_regla(row)
    viaje = get_text(row, "Número De Viaje")
    tracto = get_text(row, "Número Tracto")
    tipo_viaje = get_text(row, "Tipo Viaje")
    servicio = get_text(row, "Servicio")
    ruta = route_of(row)

    ingreso = costo = 0.0
    estado = "OK"
    obs = ""

    if regla == 1:
        ingreso = get_col(row, "I FREIGHT MEXTRANSP MEX1")
    elif regla == 2:
        ingreso = get_col(row, "I FREIGHT MEXTRANSP MEX19")
        costo = get_col(row, "C FREIGHT MEXCT TRANSP MEX76")
    elif regla == 3:
        ingreso = get_col(row, "I FREIGHT MEXTRANSP MEX38") + get_col(row, "I FREIGHT MEXTRANSP MEX61")
        costo = get_col(row, "C FREIGHT MEXCT TRANSP MEX76") + get_col(row, "C FREIGHT MEXCT TRANSP MEX84")
    else:
        return None

    if ingreso == 0 and costo == 0:
        return None

    diff = ingreso - costo
    if ingreso > 0 and costo == 0 and regla in (2, 3):
        estado = "Anomalía"
        obs = "Flete MX con ingreso pero sin costo de tercero."
    elif ingreso == 0 and costo > 0:
        estado = "Anomalía"
        obs = "Flete MX con costo pero sin ingreso."
    elif abs(diff) > 200:
        estado = "Anomalía"
        obs = f"Flete MX con diferencia {money(abs(diff))} mayor a $200."

    return AuditRow(
        auditoria="Flete MX",
        viaje=viaje,
        tracto=tracto,
        tipo_viaje=tipo_viaje,
        servicio=servicio,
        regla=regla,
        estado=estado,
        observacion=obs,
        ruta=ruta,
        metricas={
            "I Flete MX": ingreso,
            "C Flete MX": costo,
            "Diferencia": diff,
        },
    )


def audit_cruce(row: pd.Series) -> AuditRow | None:
    regla = get_regla(row)
    viaje = get_text(row, "Número De Viaje")
    tracto = get_text(row, "Número Tracto")
    tipo_viaje = get_text(row, "Tipo Viaje")
    servicio = get_text(row, "Servicio")
    ruta = route_of(row)

    i_carg = i_vac = costo = 0.0
    estado = "OK"
    obs = ""

    if regla == 1:
        i_vac = get_col(row, "I CROSS BORDER EMPTYCRUCE VACIO6")
        i_carg = get_col(row, "I CROSS BORDER LOADEDCRUCE CARGADO7")
    elif regla == 2:
        i_vac = get_col(row, "I CROSS BORDER EMPTYCRUCE VACIO24")
        i_carg = get_col(row, "I CROSS BORDER LOADEDCRUCE CARGADO25")
    elif regla == 3:
        i_vac = get_col(row, "I CROSS BORDER EMPTYCRUCE VACIO43")
        i_carg = get_col(row, "I CROSS BORDER LOADEDCRUCE CARGADO44")
        costo = get_col(row, "C CROSS BORDER LOADEDCT CRUCE CARGADO73")
    else:
        return None

    ingreso = i_carg + i_vac
    if ingreso == 0 and costo == 0:
        return None

    diff = ingreso - costo
    if regla == 3:
        if ingreso > 0 and costo == 0:
            estado = "Anomalía"
            obs = "Cruce con ingreso pero sin costo."
        elif ingreso == 0 and costo > 0:
            estado = "Anomalía"
            obs = "Cruce con costo pero sin ingreso."
        elif abs(diff) > 200:
            estado = "Anomalía"
            obs = f"Cruce con diferencia {money(abs(diff))} mayor a $200."
        if costo > 400:
            estado = "Anomalía"
            extra = " Costo de cruce fuera del rango esperado."
            obs = (obs + extra).strip()

    return AuditRow(
        auditoria="Cruce",
        viaje=viaje,
        tracto=tracto,
        tipo_viaje=tipo_viaje,
        servicio=servicio,
        regla=regla,
        estado=estado,
        observacion=obs,
        ruta=ruta,
        metricas={
            "I Cruce Cargado": i_carg,
            "I Cruce Vacío": i_vac,
            "C Cruce": costo,
            "Diferencia": diff,
        },
    )


def audit_small_concept(row: pd.Series, concepto: str) -> AuditRow | None:
    regla = get_regla(row)
    viaje = get_text(row, "Número De Viaje")
    tracto = get_text(row, "Número Tracto")
    tipo_viaje = get_text(row, "Tipo Viaje")
    servicio = get_text(row, "Servicio")
    ruta = route_of(row)

    config = {
        "Extra Stop": {
            1: ("I EXTRA STOPPARADA EXTRA5", None),
            2: ("I EXTRA STOPPARADA EXTRA23", "C EXTRA STOPCT PARADA EXTRA70"),
            3: ("I EXTRA STOPPARADA EXTRA42", "C EXTRA STOPCT PARADA EXTRA75"),
            "tol": 50,
            "max_ing": 300,
        },
        "TNU": {
            1: ("I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO14", None),
            2: ("I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO32", None),
            3: ("I TNU - TRUCK NOT USEDMOVIMIENTO EN FALSO51", "C TNU - TRUCK NOT USEDCT MOVIMIENTO EN FALSO90"),
            "tol": 50,
            "max_ing": None,
        },
        "Handling": {
            1: ("I HANDLING CHARGESMANIOBRAS13", None),
            2: ("I HANDLING CHARGESMANIOBRAS31", None),
            3: ("I HANDLING CHARGESMANIOBRAS50", "C HANDLING CHARGESCT MANIOBRAS89"),
            "tol": 50,
            "max_ing": 1500,
        },
    }
    cfg = config[concepto]
    if regla not in (1, 2, 3):
        return None

    ing_col, cost_col = cfg[regla]
    ingreso = get_col(row, ing_col) if ing_col else 0.0
    costo = get_col(row, cost_col) if cost_col else 0.0
    if ingreso == 0 and costo == 0:
        return None

    diff = ingreso - costo
    estado = "OK"
    obs = ""
    tol = cfg["tol"]
    max_ing = cfg["max_ing"]

    if costo > 0 and ingreso == 0:
        estado = "Anomalía"
        obs = f"{concepto} con costo sin ingreso."
    elif regla == 3 and ingreso > 0 and costo == 0:
        estado = "Anomalía"
        obs = f"{concepto} con ingreso de tercero pero sin costo."
    elif max_ing is not None and ingreso > max_ing:
        estado = "Anomalía"
        obs = f"{concepto} con ingreso alto {money(ingreso)}."
    elif abs(diff) > tol:
        estado = "Anomalía"
        obs = f"{concepto} con diferencia {money(abs(diff))} mayor a ${tol}."

    return AuditRow(
        auditoria=concepto,
        viaje=viaje,
        tracto=tracto,
        tipo_viaje=tipo_viaje,
        servicio=servicio,
        regla=regla,
        estado=estado,
        observacion=obs,
        ruta=ruta,
        metricas={
            f"I {concepto}": ingreso,
            f"C {concepto}": costo,
            "Diferencia": diff,
        },
    )


def audit_utilidad(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in ["Número De Viaje", "Número Tracto", "Servicio", "Tipo Viaje", "% Utilidad", "Importe Ingreso", "Importe Costo", "Importe Utilidad"] if c in df.columns]
    out = df[cols].copy()
    if "% Utilidad" in out.columns:
        out["% Utilidad Num"] = out["% Utilidad"].apply(n)
    else:
        ing = out["Importe Ingreso"].apply(n) if "Importe Ingreso" in out.columns else 0
        uti = out["Importe Utilidad"].apply(n) if "Importe Utilidad" in out.columns else 0
        out["% Utilidad Num"] = (uti / ing.replace(0, pd.NA)) * 100
        out["% Utilidad Num"] = out["% Utilidad Num"].fillna(0)
    out["Estado"] = out["% Utilidad Num"].apply(lambda x: "Anomalía" if x < 30 else "OK")
    out["Observación"] = out["% Utilidad Num"].apply(lambda x: "Utilidad menor a 30%." if x < 30 else "")
    return out


def run_all_audits(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    results: Dict[str, List[Dict[str, object]]] = {
        "Flete USA": [],
        "Flete MX": [],
        "Cruce": [],
        "Extra Stop": [],
        "TNU": [],
        "Handling": [],
    }

    for _, row in df.iterrows():
        for fn, key in [
            (audit_flete_usa, "Flete USA"),
            (audit_flete_mex, "Flete MX"),
            (audit_cruce, "Cruce"),
        ]:
            r = fn(row)
            if r:
                results[key].append(r.as_dict())

        for concepto in ["Extra Stop", "TNU", "Handling"]:
            r = audit_small_concept(row, concepto)
            if r:
                results[concepto].append(r.as_dict())

    out = {k: pd.DataFrame(v) for k, v in results.items()}
    out["Utilidad"] = audit_utilidad(df)

    if "Estatus" in df.columns:
        cancel = df[df["Estatus"].astype(str).str.upper().str.contains("CANCEL", na=False)].copy()
        out["Cancelados"] = cancel
    else:
        out["Cancelados"] = pd.DataFrame()

    anomalies = []
    for name, audit_df in out.items():
        if name == "Cancelados":
            continue
        if "Estado" in audit_df.columns:
            temp = audit_df[audit_df["Estado"] == "Anomalía"].copy()
            if not temp.empty:
                temp.insert(0, "Auditoría", name)
                anomalies.append(temp)
    out["Anomalías"] = pd.concat(anomalies, ignore_index=True) if anomalies else pd.DataFrame()
    return out


def to_excel_bytes(results: Dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in results.items():
            export_df = df.copy()
            export_df.to_excel(writer, sheet_name=name[:31], index=False)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# UI
# -----------------------------
st.title("Lincoln — Auditoría de viajes")
st.caption("Versión en Python/Streamlit convertida desde tu HTML para que la pruebes dentro de tu app.")

with st.expander("Reglas base incluidas", expanded=False):
    st.markdown(
        """
- **Regla 1:** Carretera USA + unidad capturada.
- **Regla 2:** Broker USA + unidad capturada.
- **Regla 3:** Broker USA + sin unidad.
- Tolerancias base: **$200** para Flete USA, Flete MX y Cruce; **$50** para Extra Stop, TNU y Handling.
- Auditoría adicional: **utilidad menor a 30%**.
        """
    )

uploaded = st.file_uploader("Sube el Excel de sistema", type=["xlsx", "xls"])

if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet_name = "Companies" if "Companies" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(uploaded, sheet_name=sheet_name)
        df = normalize_df(df)
        results = run_all_audits(df)

        anom = results["Anomalías"]
        total_viajes = len(df)
        cancelados = len(results["Cancelados"])
        viajes_con_anom = anom["Número De Viaje"].nunique() if not anom.empty and "Número De Viaje" in anom.columns else 0
        viajes_ok = max(total_viajes - viajes_con_anom - cancelados, 0)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total viajes", total_viajes)
        c2.metric("Sin anomalía", viajes_ok)
        c3.metric("Con anomalía", viajes_con_anom)
        c4.metric("Cancelados", cancelados)

        tab_names = ["Anomalías", "Utilidad", "Flete USA", "Flete MX", "Cruce", "Extra Stop", "TNU", "Handling", "Cancelados"]
        tabs = st.tabs(tab_names)

        for tab, name in zip(tabs, tab_names):
            with tab:
                data = results[name]
                if data.empty:
                    st.info("Sin datos en esta sección.")
                    continue

                if "Estado" in data.columns:
                    filtro = st.radio(
                        f"Filtro {name}",
                        ["Todos", "Solo anomalías", "Solo OK"],
                        horizontal=True,
                        key=f"filtro_{name}",
                    )
                    if filtro == "Solo anomalías":
                        data = data[data["Estado"] == "Anomalía"]
                    elif filtro == "Solo OK":
                        data = data[data["Estado"] == "OK"]

                st.dataframe(data, use_container_width=True, hide_index=True)

        st.download_button(
            "Descargar resultado de auditoría",
            data=to_excel_bytes(results),
            file_name="Auditoria_Lincoln_Streamlit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"No pude procesar el archivo: {e}")
else:
    st.info("Sube el archivo Excel para comenzar.")
