import streamlit as st
import pandas as pd
import json
from io import BytesIO

st.set_page_config(page_title="Fuel Solutions Parser", layout="wide")

st.title("Fuel Solutions → 3 Tablas + Filtro Año/Mes + Export Excel")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

def safe_json_loads(x):
    try:
        if pd.isna(x):
            return None
        return json.loads(x)
    except Exception:
        return None

def build_tables(df_filtered: pd.DataFrame):
    trips_rows = []
    purchases_rows = []
    onroute_rows = []

    for _, row in df_filtered.iterrows():
        fsid = row["FSID"]
        created_at = row["FSCreatedAt"]
        payload = safe_json_loads(row["FSJSON"])
        if not payload:
            continue

        # 1) Trip (1 fila por FSID)
        origin = payload.get("origin", {}) or {}
        dest = payload.get("destination", {}) or {}

        trip = {
            "FSID": fsid,
            "FSCreatedAt": created_at,
            "customer": payload.get("customer"),
            "unitNumber": payload.get("unitNumber"),
            "origin_location": origin.get("location"),
            "origin_lat": origin.get("lat"),
            "origin_lng": origin.get("lng"),
            "destination_location": dest.get("location"),
            "destination_lat": dest.get("lat"),
            "destination_lng": dest.get("lng"),
            "totalTripDistanceMiles": payload.get("totalTripDistanceMiles"),
            "totalFuelNeededGallons": payload.get("totalFuelNeededGallons"),
            "totalPurchaseNeededGallons": payload.get("totalPurchaseNeededGallons"),
            "savings": payload.get("savings"),
        }
        trips_rows.append(trip)

        # 2) fuelPurchaseLocations (N filas por FSID)
        for p in payload.get("fuelPurchaseLocations", []) or []:
            purchases_rows.append({
                "FSID": fsid,
                "FSCreatedAt": created_at,
                "loc_id": p.get("loc_id"),
                "active": p.get("active"),
                "fuelToPurchase": p.get("fuelToPurchase"),
                "lat": p.get("lat"),
                "lng": p.get("lng"),
                "location": p.get("location"),
                "price": p.get("price"),
                "include": p.get("include"),
                "interstate_exit": p.get("interstate_exit"),
            })

        # 3) fuelStationOnRoute (N filas por FSID) viene como DataFrame serializado
        fsor = payload.get("fuelStationOnRoute")
        if isinstance(fsor, dict) and "columns" in fsor and "data" in fsor:
            cols = fsor["columns"]
            data = fsor["data"]
            tmp = pd.DataFrame(data, columns=cols)
            tmp.insert(0, "FSID", fsid)
            tmp.insert(1, "FSCreatedAt", created_at)
            onroute_rows.append(tmp)

    trips_df = pd.DataFrame(trips_rows)
    purchases_df = pd.DataFrame(purchases_rows)
    onroute_df = pd.concat(onroute_rows, ignore_index=True) if onroute_rows else pd.DataFrame()

    return trips_df, purchases_df, onroute_df

def to_excel_bytes(trips_df, purchases_df, onroute_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        trips_df.to_excel(writer, index=False, sheet_name="Trip")
        purchases_df.to_excel(writer, index=False, sheet_name="Fuel Purchases")
        onroute_df.to_excel(writer, index=False, sheet_name="Stations On Route")
    output.seek(0)
    return output

if uploaded:
    # Lee solo lo que necesitamos para filtrar rápido
    base = pd.read_excel(uploaded, sheet_name="Fuel Solutions", usecols=["FSID", "FSJSON", "FSCreatedAt"])
    base["FSCreatedAt"] = pd.to_datetime(base["FSCreatedAt"], errors="coerce")

    base = base.dropna(subset=["FSCreatedAt"])
    base["Year"] = base["FSCreatedAt"].dt.year
    base["Month"] = base["FSCreatedAt"].dt.month

    years = sorted(base["Year"].dropna().unique().tolist())
    year = st.sidebar.selectbox("Año", years, index=len(years)-1 if years else 0)

    months = sorted(base.loc[base["Year"] == year, "Month"].dropna().unique().tolist())
    month = st.sidebar.selectbox("Mes", months, index=len(months)-1 if months else 0)

    filtered = base[(base["Year"] == year) & (base["Month"] == month)].copy()
    st.caption(f"Registros filtrados: {len(filtered):,} (Año={year}, Mes={month})")

    trips_df, purchases_df, onroute_df = build_tables(filtered)

    st.subheader("Tabla 1: Trip (1 fila por FSID)")
    st.dataframe(trips_df, use_container_width=True)

    st.subheader("Tabla 2: Fuel Purchases (paradas de compra)")
    st.dataframe(purchases_df, use_container_width=True)

    st.subheader("Tabla 3: Stations On Route (todas las estaciones en ruta)")
    st.dataframe(onroute_df, use_container_width=True)

    excel_bytes = to_excel_bytes(trips_df, purchases_df, onroute_df)

    st.download_button(
        label="⬇️ Descargar Excel (3 hojas)",
        data=excel_bytes,
        file_name=f"fuel_solutions_{year}_{month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Sube el Excel para empezar.")
