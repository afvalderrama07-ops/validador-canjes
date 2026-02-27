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


# -----------------------------
# LÓGICA ARCHIVO 1
# -----------------------------
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

    # -----------------------------
    # 1) Quitar duplicados
    # -----------------------------
    if id_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {id_col}")
    df = df.drop_duplicates(subset=[id_col], keep="first").copy()

    # -----------------------------
    # 2) ELIMINAR COLUMNAS (OBLIGATORIO) - inmediatamente después de duplicados
    # -----------------------------
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

    # -----------------------------
    # 3) NUEVO: AUXILIAR (día del mes) a la izquierda de Empleado
    # -----------------------------
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

    # -----------------------------
    # 4) NUEVO: columna adicional fija "NO" entre polos y monto
    # -----------------------------
    df[adicional_col] = "NO"
    if polo_col in df.columns and amt_col in df.columns:
        cols = df.columns.tolist()
        cols.remove(adicional_col)
        monto_idx = cols.index(amt_col)  # insertar justo antes del monto
        cols.insert(monto_idx, adicional_col)
        df = df[cols]

    # -----------------------------
    # 5) NUEVO: ¿Tomaste foto del comprobante? según link
    # -----------------------------
    if tomo_foto_col in df.columns and link_col in df.columns:
        df[tomo_foto_col] = df[link_col].apply(lambda x: "SI" if safe_has_value(x) else "NO")
    elif tomo_foto_col in df.columns and link_col not in df.columns:
        df[tomo_foto_col] = "NO"

    # -----------------------------
    # 6) Normalización numérica
    # -----------------------------
    df["_monto_num"] = df[amt_col].apply(to_num) if amt_col in df.columns else np.nan
    df["_polos_num"] = pd.to_numeric(df[polo_col], errors="coerce") if polo_col in df.columns else np.nan

    # -----------------------------
    # 7) DINÁMICA FOCO - REEMPLAZO CORRECTO (Temple Pato => BASES; resto => ESMALTES)
    # -----------------------------
    if din_col not in df.columns:
        raise ValueError(f"No existe la columna requerida: {din_col}")

    mask_foco = df[din_col].astype(str).str.strip().eq("Dinámica Foco")

    if cat_col in df.columns:
        mask_temple = (
            df[cat_col]
            .astype(str)
            .str.strip()
            .str.upper()
            .eq("TEMPLE PATO")
        )
        df.loc[mask_foco, cat_col] = "ESMALTES"
        df.loc[mask_foco & mask_temple, cat_col] = "BASES"

    df_foco = df[mask_foco].copy()

    # -----------------------------
    # 8) VALIDACIÓN FOCO
    # -----------------------------
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

    # -----------------------------
    # 9) VALIDACIÓN MONTO (incluye polos = 0)
    # -----------------------------
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

    # -----------------------------
    # 10) RESULTADO FINAL (OK/ERROR + motivo)
    # -----------------------------
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

    # limpiar columnas técnicas
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


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Validador de Canjes", layout="wide")
st.title("Validador de Canjes (Streamlit)")

f1 = st.file_uploader("Sube el Excel del Archivo 1", type=["xlsx"])

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
        label="Descargar resultado",
        data=out_bytes,
        file_name="resultado_archivo_1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )