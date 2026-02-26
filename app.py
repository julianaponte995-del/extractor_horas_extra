import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
import io

st.title("Procesador de Horarios 📊")

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

festivos_co = holidays.Colombia()

def limpiar_horarios(texto):
    if not isinstance(texto, str):
        return []

    filas = texto.split("\n")
    resultados = []

    for fila in filas:
        fila = fila.strip().upper()

        match_dia = re.search(r'\b(LU|MA|MI|JU|VI|SA|DO)\b', fila)
        if not match_dia:
            continue

        dia = match_dia.group(1)

        match_horas = re.search(r'(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})', fila)
        if match_horas:
            resultados.append({
                "dia": dia,
                "hora_inicio": match_horas.group(1),
                "hora_fin": match_horas.group(2)
            })

    return resultados

if archivo is not None:

    df = pd.read_excel(archivo)

    df["HORAS"] = df["HORAS"].str.replace("NO TIENE", "0")
    df = df[df["HORAS"] != "0"]
    df = df[df["NPLAN"] != 800]
    df["MATERIA_INI"] = pd.to_datetime(df["MATERIA_INI"], dayfirst=True)
    df["MATERIA_FIN"] = pd.to_datetime(df["MATERIA_FIN"], dayfirst=True)

    # Procesamiento
    df["horarios_lista"] = df["HORAS"].apply(limpiar_horarios)
    df = df.explode("horarios_lista").reset_index(drop=True)

    horarios_df = pd.json_normalize(df["horarios_lista"]).reset_index(drop=True)

    df_final = pd.concat([df.drop(columns=["horarios_lista"]), horarios_df], axis=1)
    df_final = df_final.dropna(subset=["dia"])

    # Minutos
    inicio_recargo_global = 780

    df_final['minutos_inicio'] = (pd.to_timedelta(df_final['hora_inicio'] + ':00') - pd.to_timedelta('06:00:00')).dt.total_seconds() / 60
    df_final['minutos_fin'] = (pd.to_timedelta(df_final['hora_fin'] + ':00') - pd.to_timedelta('06:00:00')).dt.total_seconds() / 60

    df_final['minutos_recargo'] = (
        df_final['minutos_fin'] - np.maximum(df_final['minutos_inicio'], inicio_recargo_global)
    ).clip(lower=0)

    # Filtrar por aquellos que tengan minutos de recargo mayores a 0
    df_final = df_final[df_final["minutos_recargo"] > 0]

    # Eliminar las columnas que no aportan
    df_final = df_final.drop(columns=['MATERIA_ACTIVIDAD', 'TOTAL_HORAS', 'GRUPO'])

    # Calendario
    dias_map = {
        "Monday": "LU",
        "Tuesday": "MA",
        "Wednesday": "MI",
        "Thursday": "JU",
        "Friday": "VI",
        "Saturday": "SA",
        "Sunday": "DO"
    }

    filas = []

    for _, row in df_final.iterrows():
        fechas = pd.date_range(start=row["MATERIA_INI"], end=row["MATERIA_FIN"], freq="D")

        temp = pd.DataFrame({"fecha": fechas})
        temp["dia"] = temp["fecha"].dt.day_name().map(dias_map)
        temp = temp[temp["dia"] == row["dia"]]

        for col in df_final.columns:
            temp[col] = row[col]

        filas.append(temp)

    df_expandido = pd.concat(filas, ignore_index=True)

    df_expandido["es_festivo"] = df_expandido["fecha"].apply(lambda x: x in festivos_co)

    df_expandido["minutos_recargo"] = df_expandido.apply(
        lambda row: 0 if row["es_festivo"] else row["minutos_recargo"],
        axis=1
    )

    df_expandido["horas_recargo"] = df_expandido["minutos_recargo"] / 60

    # FILTRAR POR LOS QUE TIENEN HORAS
    final = df_expandido[df_expandido["horas_recargo"] > 0].copy()

    # QUITAR SEMANA SANTA
    final = final[(final["fecha"] < "2026-03-29") | (final["fecha"] > "2026-04-05")]

    # Poner columna nombre de mes 
    final.insert(1, "mes", final["fecha"].dt.month_name(locale="es_ES"))
    final["mes"] = final["mes"].str.upper() # poner en mayuscula

    # Agrupación
    resultado = final.groupby(["NOMBRE", "MATERIA_ACTIVIDAD"])["horas_recargo"].sum().reset_index()

    pivot = resultado.pivot_table(
        index="NOMBRE",
        values="horas_recargo",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # Mostrar resultados
    st.subheader("Resultados detallados")
    st.dataframe(final)

    # 🔥 EXPORTAR BIEN (AQUÍ ESTABA TU ERROR)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final.to_excel(writer, index=False, sheet_name="Detalle")

    st.download_button(
        label="Descargar Excel 📥",
        data=output.getvalue(),
        file_name="recargos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )













