from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# Configuracion base
# ============================================================
MIN_UTILITY_PCT = 30.0
USA_TOLERANCE = 200.0

BASE_COLUMNS = [
    "Estado Factura",
    "Numero De Viaje",
    "Servicio",
    "Sucursal",
    "Ciudad Origen",
    "Estado Origen",
    "Ciudad Destino",
    "Estado Destino",
    "Tipo Viaje",
    "Linea Mexicana",
    "Linea Americana",
    "Numero Tracto",
    "Importe Ingreso",
    "Importe Costo",
    "Importe Utilidad",
    "% Utilidad",
]

RENAME_MAP = {
    "Numero De Viaje": "numero_viaje",
    "Servicio": "servicio",
    "Sucursal": "sucursal",
    "Ciudad Origen": "ciudad_origen",
    "Estado Origen": "estado_origen",
    "Ciudad Destino": "ciudad_destino",
    "Estado Destino": "estado_destino",
    "Tipo Viaje": "tipo_viaje",
    "Linea Mexicana": "linea_mexicana",
    "Linea Americana": "linea_americana",
    "Numero Tracto": "numero_tracto",
    "Importe Ingreso": "importe_ingreso",
    "Importe Costo": "importe_costo",
    "Importe Utilidad": "importe_utilidad",
    "% Utilidad": "pct_utilidad",
    "Estado Factura": "estado_factura",
}


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return re.sub(r"\s+", " ", text)


def canonical_col(name: str) -> str:
    text = normalize_text(name)
    text = text.replace("#", "num")
    text = re.sub(r"[^A-Za-z0-9]+", "_", text)
    return text.strip("_")


@dataclass
class AuditIssue:
    numero_viaje: str
    regla: str
    severidad: str
    observacion: str
    detalle: str


def find_sheet_name(excel_file: pd.ExcelFile) -> str:
    preferred = ["Companies", "Company", "Reporte", "Datos"]
    normalized = {normalize_text(s).lower(): s for s in excel_file.sheet_names}
    for pref in preferred:
        key = normalize_text(pref).lower()
        if key in normalized:
            return normalized[key]
    return excel_file.sheet_names[0]


def load_input_dataframe(uploaded_file) -> pd.DataFrame:
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = find_sheet_name(excel)
    df = pd.read_excel(excel, sheet_name=sheet_name)
    df.columns = [canonical_col(c) for c in df.columns]
    return df


def ensure_required_columns(df: pd.DataFrame, required: List[str]) -> Tuple[bool, List[str]]:
    missing = [col for col in required if col not in df.columns]
    return len(missing) == 0, missing


def first_existing(df: pd.DataFrame, options: List[str]) -> Optional[str]:
    for col in options:
        if col in df.columns:
            return col
    return None


def build_working_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()

    # Mapear columnas conocidas a nombres estables internos
    resolved: Dict[str, str] = {}
    for original, internal in RENAME_MAP.items():
        canon = canonical_col(original)
        if canon in df.columns:
            resolved[internal] = canon

    required_internals = [
        "numero_viaje", "servicio", "ciudad_origen", "estado_origen",
        "ciudad_destino", "estado_destino", "numero_tracto", "importe_ingreso",
        "importe_costo", "importe_utilidad", "pct_utilidad"
    ]
    missing = [x for x in required_internals if x not in resolved]
    if missing:
        raise ValueError(
            "Faltan columnas clave para procesar el archivo: " + ", ".join(missing)
        )

    out = pd.DataFrame()
    for internal, canon in resolved.items():
        out[internal] = df[canon]

    for c in ["sucursal", "tipo_viaje", "linea_mexicana", "linea_americana", "estado_factura"]:
        if c in resolved:
            out[c] = df[resolved[c]]
        else:
            out[c] = np.nan

    out["ruta"] = (
        out["ciudad_origen"].fillna("").astype(str).str.strip() + ", " +
        out["estado_origen"].fillna("").astype(str).str.strip() + " - " +
        out["ciudad_destino"].fillna("").astype(str).str.strip() + ", " +
        out["estado_destino"].fillna("").astype(str).str.strip()
    ).str.replace(r"^, | - , ", "", regex=True)

    numeric_candidates = [
        "importe_ingreso", "importe_costo", "importe_utilidad", "pct_utilidad"
    ]
    for col in numeric_candidates:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)

    out["tiene_tracto"] = out["numero_tracto"].notna() & (out["numero_tracto"].astype(str).str.strip() != "")
    out["servicio_norm"] = out["servicio"].fillna("").astype(str).map(normalize_text).str.upper()
    out["es_broker_usa"] = out["servicio_norm"].eq("BROKER USA")
    out["es_carretera_usa"] = out["servicio_norm"].eq("CARRETERA USA")

    # Mantener columnas originales de ingresos/costos/acreedores para reglas futuras
    out = pd.concat([out, df], axis=1)
    return out


def get_col(df: pd.DataFrame, raw_name: str) -> Optional[str]:
    canon = canonical_col(raw_name)
    return canon if canon in df.columns else None


def sum_columns(df: pd.DataFrame, names: List[str]) -> pd.Series:
    cols = [get_col(df, n) for n in names]
    cols = [c for c in cols if c is not None]
    if not cols:
        return pd.Series(0.0, index=df.index)
    temp = df[cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return temp.sum(axis=1)


def add_flete_usa_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    usa_propio_income_cols = [
        "I FREIGHT USATRANSP USA2",
        "I FUEL CHARGES DIESEL3",
    ]
    usa_broker_propio_income_cols = [
        "I FREIGHT USATRANSP USA20",
        "I FUEL CHARGES DIESEL21",
    ]
    usa_broker_tercero_income_cols = [
        "I FREIGHT USATRANSP USA39",
        "I FUEL CHARGES DIESEL40",
        "I FREIGHT USATRANSP USA56",
    ]
    usa_cost_cols = [
        "C FREIGHT USACT TRANSP USA72",
        "C FREIGHT USACT TRANSP USA77",
        "C FREIGHT USACT TRANSP USA78",
    ]

    out["usa_ing_propio"] = sum_columns(out, usa_propio_income_cols)
    out["usa_ing_broker_propio"] = sum_columns(out, usa_broker_propio_income_cols)
    out["usa_ing_broker_tercero"] = sum_columns(out, usa_broker_tercero_income_cols)
    out["usa_cost_total"] = sum_columns(out, usa_cost_cols)

    out["escenario_flete_usa"] = np.select(
        [
            out["es_carretera_usa"] & out["tiene_tracto"],
            out["es_broker_usa"] & out["tiene_tracto"],
            out["es_broker_usa"] & ~out["tiene_tracto"],
        ],
        [
            "CARRETERA_USA_PROPIO",
            "BROKER_USA_PROPIO",
            "BROKER_USA_TERCERO",
        ],
        default="OTRO",
    )
    return out


def build_audit(df: pd.DataFrame, min_utility_pct: float = MIN_UTILITY_PCT, usa_tolerance: float = USA_TOLERANCE) -> pd.DataFrame:
    issues: List[AuditIssue] = []

    for _, row in df.iterrows():
        viaje = str(row.get("numero_viaje", ""))

        # Regla 1: utilidad minima
        pct_util = float(row.get("pct_utilidad", 0) or 0)
        if pct_util < min_utility_pct:
            issues.append(AuditIssue(
                numero_viaje=viaje,
                regla="UTILIDAD_MINIMA",
                severidad="Alta",
                observacion="Utilidad menor al 30%.",
                detalle=f"% utilidad actual: {pct_util:.2f}% (minimo esperado {min_utility_pct:.2f}%)",
            ))

        escenario = row.get("escenario_flete_usa", "OTRO")
        usa_propio = float(row.get("usa_ing_propio", 0) or 0)
        usa_broker_propio = float(row.get("usa_ing_broker_propio", 0) or 0)
        usa_broker_tercero = float(row.get("usa_ing_broker_tercero", 0) or 0)
        usa_cost_total = float(row.get("usa_cost_total", 0) or 0)

        # Regla 2: coherencia de columnas de ingreso por escenario
        if escenario == "CARRETERA_USA_PROPIO":
            if usa_broker_propio > 0 or usa_broker_tercero > 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_INGRESO_CRUZADO",
                    severidad="Media",
                    observacion="Ingreso de Flete USA capturado en columna distinta al escenario esperado.",
                    detalle=(
                        f"Escenario esperado: carretera usa con unidad propia. "
                        f"Detectado broker propio={usa_broker_propio:.2f}, broker tercero={usa_broker_tercero:.2f}."
                    ),
                ))
            if usa_cost_total > 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_COSTO_NO_ESPERADO",
                    severidad="Media",
                    observacion="Hay costo de Flete USA en un viaje que luce como unidad propia.",
                    detalle=f"Costo detectado: {usa_cost_total:.2f}",
                ))

        elif escenario == "BROKER_USA_PROPIO":
            if usa_propio > 0 or usa_broker_tercero > 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_INGRESO_CRUZADO",
                    severidad="Media",
                    observacion="Ingreso de Flete USA capturado en columna distinta al escenario esperado.",
                    detalle=(
                        f"Escenario esperado: broker usa con unidad propia. "
                        f"Detectado carretera propio={usa_propio:.2f}, broker tercero={usa_broker_tercero:.2f}."
                    ),
                ))
            if usa_cost_total > 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_COSTO_NO_ESPERADO",
                    severidad="Media",
                    observacion="Hay costo de Flete USA donde no deberia existir por unidad propia.",
                    detalle=f"Costo detectado: {usa_cost_total:.2f}",
                ))

        elif escenario == "BROKER_USA_TERCERO":
            if usa_broker_tercero <= 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_INGRESO_FALTANTE",
                    severidad="Alta",
                    observacion="Escenario broker usa tercero sin ingreso en la columna esperada.",
                    detalle="No se detecto monto en las columnas de ingreso de broker tercero/filial.",
                ))
            if usa_cost_total <= 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_COSTO_FALTANTE",
                    severidad="Alta",
                    observacion="Broker USA con tercero/filial sin costo de Flete USA.",
                    detalle="Se esperaba costo porque el tractor no viene capturado.",
                ))
            elif abs(usa_broker_tercero - usa_cost_total) > usa_tolerance:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_DIFERENCIA_COSTO",
                    severidad="Alta",
                    observacion="La diferencia entre ingreso y costo de Flete USA excede la tolerancia.",
                    detalle=(
                        f"Ingreso broker tercero={usa_broker_tercero:.2f}, "
                        f"costo={usa_cost_total:.2f}, tolerancia={usa_tolerance:.2f}."
                    ),
                ))
            if usa_propio > 0 or usa_broker_propio > 0:
                issues.append(AuditIssue(
                    numero_viaje=viaje,
                    regla="FLETE_USA_INGRESO_CRUZADO",
                    severidad="Media",
                    observacion="Hay montos en columnas de ingreso de otro escenario.",
                    detalle=(
                        f"Carretera propio={usa_propio:.2f}, broker propio={usa_broker_propio:.2f}."
                    ),
                ))

    audit_df = pd.DataFrame([issue.__dict__ for issue in issues])
    if audit_df.empty:
        return pd.DataFrame(columns=["numero_viaje", "regla", "severidad", "observacion", "detalle"])
    return audit_df.sort_values(["severidad", "numero_viaje", "regla"], ascending=[True, True, True]).reset_index(drop=True)


def compare_with_manual(audit_df: pd.DataFrame, uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    try:
        excel = pd.ExcelFile(uploaded_file)
        if "Auditoria" not in excel.sheet_names:
            return pd.DataFrame(), pd.DataFrame()
        manual = pd.read_excel(excel, sheet_name="Auditoria")
        manual.columns = [canonical_col(c) for c in manual.columns]
        viaje_col = first_existing(manual, ["numero_de_viaje", "numero_viaje"])
        if not viaje_col:
            return pd.DataFrame(), pd.DataFrame()
        manual_viajes = (
            manual[viaje_col]
            .dropna()
            .astype(str)
            .str.strip()
            .drop_duplicates()
            .to_frame(name="numero_viaje")
        )
        auto_viajes = audit_df[["numero_viaje"]].drop_duplicates() if not audit_df.empty else pd.DataFrame(columns=["numero_viaje"])

        detectados = manual_viajes.merge(auto_viajes, on="numero_viaje", how="inner")
        faltantes = manual_viajes.merge(auto_viajes, on="numero_viaje", how="left", indicator=True)
        faltantes = faltantes[faltantes["_merge"] == "left_only"].drop(columns=["_merge"])
        return detectados, faltantes
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def to_excel_bytes(summary_df: pd.DataFrame, audit_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Resumen", index=False)
        audit_df.to_excel(writer, sheet_name="Auditoria_Automatica", index=False)
        detail_df.to_excel(writer, sheet_name="Detalle_Trabajo", index=False)
    output.seek(0)
    return output.getvalue()


def render_summary(df: pd.DataFrame, audit_df: pd.DataFrame, min_utility_pct: float = MIN_UTILITY_PCT) -> pd.DataFrame:
    total_viajes = len(df)
    viajes_con_hallazgo = audit_df["numero_viaje"].nunique() if not audit_df.empty else 0
    total_hallazgos = len(audit_df)
    utilidad_baja = int((df["pct_utilidad"] < min_utility_pct).sum())
    broker_tercero = int((df["escenario_flete_usa"] == "BROKER_USA_TERCERO").sum())
    broker_sin_costo = int(((df["escenario_flete_usa"] == "BROKER_USA_TERCERO") & (df["usa_cost_total"] <= 0)).sum())

    return pd.DataFrame({
        "indicador": [
            "Viajes analizados",
            "Viajes con hallazgo",
            "Hallazgos totales",
            "Viajes con utilidad < 30%",
            "Viajes Broker USA tercero/filial",
            "Broker USA tercero/filial sin costo",
        ],
        "valor": [
            total_viajes,
            viajes_con_hallazgo,
            total_hallazgos,
            utilidad_baja,
            broker_tercero,
            broker_sin_costo,
        ],
    })


def main() -> None:
    st.set_page_config(page_title="Pruebas Auditoria Lincoln", layout="wide")
    st.title("Pruebas de auditoria Lincoln")
    st.caption(
        "Esta pagina sirve para probar reglas basicas de auditoria sobre el archivo descargado del sistema. "
        "Se enfoca primero en utilidad minima y Flete USA segun la logica documentada."
    )

    with st.sidebar:
        st.subheader("Parametros")
        min_utility_pct = st.number_input("Utilidad minima (%)", min_value=0.0, max_value=100.0, value=MIN_UTILITY_PCT, step=1.0)
        usa_tolerance = st.number_input("Tolerancia Flete USA", min_value=0.0, value=USA_TOLERANCE, step=10.0)
        st.markdown("---")
        st.write("Archivo esperado:")
        st.write("- Excel descargado del sistema")
        st.write("- Hoja base tipo Companies")

    uploaded_file = st.file_uploader("Sube el archivo Excel del sistema", type=["xlsx"])
    if uploaded_file is None:
        st.info("Sube un archivo para ejecutar la auditoria.")
        return

    try:
        raw_df = load_input_dataframe(uploaded_file)
        working_df = build_working_df(raw_df)
        working_df = add_flete_usa_columns(working_df)
        audit_df = build_audit(working_df, min_utility_pct=float(min_utility_pct), usa_tolerance=float(usa_tolerance))
        summary_df = render_summary(working_df, audit_df, min_utility_pct=float(min_utility_pct))
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Viajes analizados", len(working_df))
    c2.metric("Viajes con hallazgo", audit_df["numero_viaje"].nunique() if not audit_df.empty else 0)
    c3.metric("Hallazgos totales", len(audit_df))

    st.subheader("Resumen")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.subheader("Hallazgos detectados")
    if audit_df.empty:
        st.success("No se detectaron hallazgos con las reglas actuales.")
    else:
        regla_sel = st.multiselect(
            "Filtrar por regla",
            options=sorted(audit_df["regla"].dropna().unique().tolist()),
            default=sorted(audit_df["regla"].dropna().unique().tolist()),
        )
        sev_sel = st.multiselect(
            "Filtrar por severidad",
            options=sorted(audit_df["severidad"].dropna().unique().tolist()),
            default=sorted(audit_df["severidad"].dropna().unique().tolist()),
        )
        filtered = audit_df[
            audit_df["regla"].isin(regla_sel) &
            audit_df["severidad"].isin(sev_sel)
        ].copy()
        st.dataframe(filtered, use_container_width=True, hide_index=True)

    st.subheader("Detalle de trabajo")
    detail_cols = [
        "numero_viaje", "servicio", "numero_tracto", "ruta", "pct_utilidad",
        "escenario_flete_usa", "usa_ing_propio", "usa_ing_broker_propio",
        "usa_ing_broker_tercero", "usa_cost_total",
    ]
    detail_df = working_df[detail_cols].copy().sort_values(["numero_viaje"])
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

    # Comparacion opcional con una hoja manual llamada Auditoria en el mismo archivo
    detectados, faltantes = compare_with_manual(audit_df, uploaded_file)
    if not detectados.empty or not faltantes.empty:
        st.subheader("Comparacion contra hoja manual 'Auditoria'")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.write("Viajes manuales tambien detectados")
            st.dataframe(detectados, use_container_width=True, hide_index=True)
        with cc2:
            st.write("Viajes presentes en manual pero no detectados aun")
            st.dataframe(faltantes, use_container_width=True, hide_index=True)

    st.download_button(
        label="Descargar resultado en Excel",
        data=to_excel_bytes(summary_df, audit_df, detail_df),
        file_name="resultado_auditoria_lincoln.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Notas de esta version"):
        st.write("- Ya incluye la regla de utilidad minima.")
        st.write("- Ya incluye la logica base de Flete USA segun servicio y numero de tracto.")
        st.write("- Ya valida costo esperado en Broker USA con tercero/filial.")
        st.write("- Ya compara ingreso vs costo con tolerancia configurable.")
        st.write("- Se puede extender despues para Flete Mex, Cruce y otros conceptos.")


if __name__ == "__main__":
    main()
