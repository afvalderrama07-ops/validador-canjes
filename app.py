import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Helpers
# -----------------------------
def to_num(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    s = re.sub(r"[^\d,.\-]", "", s)

    if s.count(",") > 0 and s.count(".") > 0:
        s = s.replace(",", "")
    else:
        if s.count(",") > 0 and s.count(".") == 0:
            s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return np.nan


def build_excel_bytes(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


def safe_has_value(x) -> bool:
    if pd.isna(x):
        return False
    s = str(x).strip()
    return s != "" and s.lower() != "nan"


# ============================================================
# ARCHIVO 1 (tu lógica actual + columnas nuevas + drop columns)
# ============================================================
def procesar_archivo_1(df: pd.DataFrame):
    id_col = "ID de la encuesta"
    fecha_col = "Fecha y hora de la encuesta"
    empleado_col = "Empleado"

    din_col = "Indicar tipo de Dinámica a canjear"
    cat_col = "Especificar en que categoría realizó mas compra."
    polo_col = "¿Cantidad de promocional entregado? - POLO QROMA"
    adicional_col = "¿Realizaste una entrega de promocional adicional?  - POLO QROMA"
    amt_col = "Monto total del comprobante (de productos participantes en canjes regulares)"
    tomo_foto_col = "¿Tomaste foto del comprobante?"
    link_col = "Foto del comprobante (Boleta Factura o Ticket de pago)"

    # 1) Quitar duplicados
    if id_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {id_col}")
    df = df.drop_duplicates(subset=[id_col], keep="first").copy()

    # 2) Eliminar columnas indicadas (Archivo 1)
    cols_to_drop = [
        "¿Se entregó este promocional? - TICKET SORTEO TV ENERO 2026",
        '"RECUERDA: TODA ESTA INFORMACIÓN DEBE LLENARSE TODOS LOS DATOS DEL CLIENTE. NO SE PERMITIRÁ REGISTROS INCOMPLETOS, SI EL CLIENTE NO PERMITE TOMAR FOTO DE SU BOLETA COMUNICARLE QUE NO PODRA PARTICIPAR"',
        "Pregunta adicional - Tu pdv pertenece a la ciudad de CUSCO o AREQUIPA",
        "Cantidad de promocional entregado - TICKET SORTEO TV ENERO 2026",
        "Pregunta adicional - Indicar tu nombre y apellido",
        "Pregunta adicional - Indicar tu numero celular o teléfono (NO AGREGAR ESPACIOS)",
        "Pregunta adicional - Indicar tu numero de DNI O CE (INCLUIR CEROS)",
        "Pregunta adicional - Ingresar el NUMERO DE CORRELATIVO del  ticket 1",
        "Pregunta adicional - Ingresar el NUMERO DE CORRELATIVO del  ticket 2",
        "Pregunta adicional - Especificar en que marca realizó mas compra.",
    ]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors="ignore").copy()

    # 3) AUXILIAR (día del mes) a la izquierda de Empleado
    if fecha_col in df.columns:
        dt = pd.to_datetime(df[fecha_col], errors="coerce", dayfirst=True)
        df["AUXILIAR"] = dt.dt.day
    else:
        df["AUXILIAR"] = np.nan

    if empleado_col in df.columns:
        cols = df.columns.tolist()
        cols.remove("AUXILIAR")
        emp_idx = cols.index(empleado_col)
        cols.insert(emp_idx, "AUXILIAR")
        df = df[cols]

    # 4) Columna adicional fija "NO" entre polos y monto
    df[adicional_col] = "NO"
    if polo_col in df.columns and amt_col in df.columns:
        cols = df.columns.tolist()
        cols.remove(adicional_col)
        monto_idx = cols.index(amt_col)
        cols.insert(monto_idx, adicional_col)
        df = df[cols]

    # 5) ¿Tomaste foto del comprobante? = SI si hay link, sino NO
    if tomo_foto_col in df.columns and link_col in df.columns:
        df[tomo_foto_col] = df[link_col].apply(lambda x: "SI" if safe_has_value(x) else "NO")
    elif tomo_foto_col in df.columns and link_col not in df.columns:
        df[tomo_foto_col] = "NO"

    # 6) Normalización numérica
    df["_monto_num"] = df[amt_col].apply(to_num) if amt_col in df.columns else np.nan
    df["_polos_num"] = pd.to_numeric(df[polo_col], errors="coerce") if polo_col in df.columns else np.nan

    # 7) FOCO: Temple Pato => BASES; resto => ESMALTES
    if din_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {din_col}")

    mask_foco = df[din_col].astype(str).str.strip().eq("Dinámica Foco")
    if cat_col in df.columns:
        mask_temple = (
            df[cat_col].astype(str).str.strip().str.upper().eq("TEMPLE PATO")
        )
        df.loc[mask_foco, cat_col] = "ESMALTES"
        df.loc[mask_foco & mask_temple, cat_col] = "BASES"

    df_foco = df[mask_foco].copy()

    # 8) Validación FOCO
    foco_errors = df_foco[
        ((df_foco["_polos_num"] == 1) & (df_foco["_monto_num"] < 20)) |
        ((df_foco["_polos_num"] == 2) & (df_foco["_monto_num"] < 40))
    ].copy()

    if not foco_errors.empty:
        foco_errors["motivo_error"] = foco_errors.apply(
            lambda r: "FOCO: entregó 1 polo con monto < 20"
            if r["_polos_num"] == 1
            else "FOCO: entregó 2 polos con monto < 40",
            axis=1,
        )

    # 9) Validación MONTO (incluye polos=0)
    mask_monto = df[din_col].astype(str).str.strip().eq("Dinámica Monto")
    df_monto = df[mask_monto].copy()

    monto_errors = df_monto[
        ((df_monto["_polos_num"] == 0) & (df_monto["_monto_num"] < 200)) |
        ((df_monto["_polos_num"] == 1) & (df_monto["_monto_num"] < 200)) |
        ((df_monto["_polos_num"] == 2) & (df_monto["_monto_num"] < 300))
    ].copy()

    if not monto_errors.empty:
        def motivo_monto(r):
            if r["_polos_num"] == 0:
                return "MONTO: polos=0 con monto < 200"
            if r["_polos_num"] == 1:
                return "MONTO: entregó 1 polo con monto < 200"
            return "MONTO: entregó 2 polos con monto < 300"

        monto_errors["motivo_error"] = monto_errors.apply(motivo_monto, axis=1)

    # 10) Resultado final
    df_out = df.copy()
    foco_bad_ids = set(foco_errors[id_col].dropna()) if not foco_errors.empty else set()
    monto_bad_ids = set(monto_errors[id_col].dropna()) if not monto_errors.empty else set()
    bad_ids = foco_bad_ids.union(monto_bad_ids)

    df_out["estado"] = np.where(df_out[id_col].isin(bad_ids), "ERROR", "OK")
    df_out["motivo"] = ""

    if bad_ids:
        motivo_map = {}
        if not foco_errors.empty and "motivo_error" in foco_errors.columns:
            for _, r in foco_errors[[id_col, "motivo_error"]].iterrows():
                motivo_map.setdefault(r[id_col], []).append(r["motivo_error"])
        if not monto_errors.empty and "motivo_error" in monto_errors.columns:
            for _, r in monto_errors[[id_col, "motivo_error"]].iterrows():
                motivo_map.setdefault(r[id_col], []).append(r["motivo_error"])
        df_out["motivo"] = df_out[id_col].map(lambda x: " | ".join(motivo_map.get(x, [])))

    report_cols = [id_col, din_col, polo_col, amt_col, "estado", "motivo"]
    if link_col in df_out.columns:
        report_cols.append(link_col)
    df_errors_report = df_out[df_out["estado"] == "ERROR"][report_cols].copy()

    # Limpiar columnas técnicas
    df_out = df_out.drop(columns=["_monto_num", "_polos_num"], errors="ignore")
    df_foco = df_foco.drop(columns=["_monto_num", "_polos_num"], errors="ignore")
    df_monto = df_monto.drop(columns=["_monto_num", "_polos_num"], errors="ignore")

    resumen = {
        "total_filas": int(len(df_out)),
        "total_ok": int((df_out["estado"] == "OK").sum()),
        "total_error": int((df_out["estado"] == "ERROR").sum()),
        "errores_foco": int(len(foco_errors)),
        "errores_monto": int(len(monto_errors)),
    }

    sheets = {
        "RESULTADO": df_out,
        "ERRORES": df_errors_report,
        "FOCO_FILTRADO": df_foco,
        "MONTO_FILTRADO": df_monto,
    }
    return sheets, resumen


# ============================================================
# ARCHIVO 2 (Ticket Sorteo)
# ============================================================
def procesar_archivo_2(df: pd.DataFrame):
    id_col = "ID de la encuesta"
    ticket_col = "¿Se entregó este promocional? - TICKET SORTEO TV ENERO 2026"
    cant_ticket_col = "Cantidad de promocional entregado - TICKET SORTEO TV ENERO 2026"
    amt_col = "Monto total del comprobante (de productos participantes en canjes regulares)"
    tomo_foto_col = "¿Tomaste foto del comprobante?"

    # 1) Quitar duplicados por ID
    if id_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {id_col}")
    df = df.drop_duplicates(subset=[id_col], keep="first").copy()

    # 2) Eliminar columnas indicadas (Archivo 2)
    cols_to_drop = [
        "Especificar en que categoría realizó mas compra.",
        "Indicar tipo de Dinámica a canjear",
        "Especificar en que marca realizó mas compra.",
        "Especificar la Dinámica Foco",
        "¿Cantidad de promocional entregado? - POLO QROMA",
    ]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors="ignore").copy()

    # 3) Eliminar filas donde Ticket Sorteo = NO
    if ticket_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {ticket_col}")

    ticket_norm = df[ticket_col].astype(str).str.strip().str.upper()
    df_filtrado = df[ticket_norm.ne("NO")].copy()

    # 4) Completar cantidad ticket vacía con 1
    if cant_ticket_col not in df_filtrado.columns:
        raise ValueError(f"No existe la columna requerida: {cant_ticket_col}")

    df_filtrado[cant_ticket_col] = pd.to_numeric(df_filtrado[cant_ticket_col], errors="coerce")
    df_filtrado[cant_ticket_col] = df_filtrado[cant_ticket_col].fillna(1).astype(int)

    # 5) Validación monto >= 200 / >= 300
    if amt_col not in df_filtrado.columns:
        raise ValueError(f"No existe la columna requerida: {amt_col}")

    df_filtrado["_monto_num"] = df_filtrado[amt_col].apply(to_num)

    errores = df_filtrado[
        ((df_filtrado[cant_ticket_col] == 1) & (df_filtrado["_monto_num"] < 200)) |
        ((df_filtrado[cant_ticket_col] == 2) & (df_filtrado["_monto_num"] < 300))
    ].copy()

    df_filtrado["estado"] = "OK"
    df_filtrado["motivo"] = ""

    if not errores.empty:
        errores["motivo_error"] = errores.apply(
            lambda r: "TICKET: cantidad=1 con monto < 200"
            if r[cant_ticket_col] == 1
            else "TICKET: cantidad=2 con monto < 300",
            axis=1,
        )

        bad_ids = set(errores[id_col].dropna())
        df_filtrado.loc[df_filtrado[id_col].isin(bad_ids), "estado"] = "ERROR"

        motivo_map = {}
        for _, r in errores[[id_col, "motivo_error"]].iterrows():
            motivo_map.setdefault(r[id_col], []).append(r["motivo_error"])
        df_filtrado["motivo"] = df_filtrado[id_col].map(lambda x: " | ".join(motivo_map.get(x, [])))

    # 6) ¿Tomaste foto del comprobante? = SI en TODAS las filas (sin condiciones)
    if tomo_foto_col in df_filtrado.columns:
        df_filtrado[tomo_foto_col] = "SI"
    else:
        df_filtrado[tomo_foto_col] = "SI"

    # 7) ID * 1000 (mantener como entero)
    df_filtrado[id_col] = pd.to_numeric(df_filtrado[id_col], errors="coerce")
    df_filtrado[id_col] = (df_filtrado[id_col] * 1000).astype("Int64")

    # Reporte errores (compacto)
    report_cols = [id_col, ticket_col, cant_ticket_col, amt_col, "estado", "motivo"]
    if tomo_foto_col in df_filtrado.columns:
        report_cols.append(tomo_foto_col)

    df_errors_report = df_filtrado[df_filtrado["estado"] == "ERROR"][report_cols].copy()

    # Quitar columna técnica
    df_filtrado = df_filtrado.drop(columns=["_monto_num"], errors="ignore")

    resumen = {
        "total_filas": int(len(df_filtrado)),
        "total_ok": int((df_filtrado["estado"] == "OK").sum()),
        "total_error": int((df_filtrado["estado"] == "ERROR").sum()),
        "filas_eliminadas_por_NO": int(len(df) - len(df_filtrado)),
    }

    sheets = {
        "RESULTADO": df_filtrado,
        "ERRORES": df_errors_report,
    }

    return sheets, resumen


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Validador de Canjes", layout="wide")
st.title("Validador de Canjes (Streamlit)")

tab1, tab2 = st.tabs(["Archivo 1 (Polos)", "Archivo 2 (Ticket Sorteo)"])

with tab1:
    st.subheader("Archivo 1")
    f1 = st.file_uploader("Sube el Excel del Archivo 1", type=["xlsx"], key="file1")
    if f1 is not None:
        df1 = pd.read_excel(f1)
        sheets, resumen = procesar_archivo_1(df1)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total", resumen["total_filas"])
        c2.metric("OK", resumen["total_ok"])
        c3.metric("ERROR", resumen["total_error"])
        c4.metric("Errores FOCO", resumen["errores_foco"])
        c5.metric("Errores MONTO", resumen["errores_monto"])

        st.dataframe(sheets["ERRORES"], use_container_width=True)

        out_bytes = build_excel_bytes(sheets)
        st.download_button(
            label="Descargar resultado (Archivo 1)",
            data=out_bytes,
            file_name="resultado_archivo_1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab2:
    st.subheader("Archivo 2 (Ticket Sorteo)")
    f2 = st.file_uploader("Sube el Excel del Archivo 2", type=["xlsx"], key="file2")
    if f2 is not None:
        df2 = pd.read_excel(f2)
        sheets2, resumen2 = procesar_archivo_2(df2)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total (post filtro)", resumen2["total_filas"])
        c2.metric("OK", resumen2["total_ok"])
        c3.metric("ERROR", resumen2["total_error"])
        c4.metric("Eliminadas por NO", resumen2["filas_eliminadas_por_NO"])

        if resumen2["total_filas"] == 0:
            st.warning("No hay registros con SI (se eliminaron todos por NO). Igual puedes descargar el Excel como evidencia.")

        st.dataframe(sheets2["ERRORES"], use_container_width=True)

        out_bytes2 = build_excel_bytes(sheets2)
        st.download_button(
            label="Descargar resultado (Archivo 2)",
            data=out_bytes2,
            file_name="resultado_archivo_2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
