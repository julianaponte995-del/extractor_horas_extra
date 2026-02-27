import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
import io

st.title("Procesador de Horarios 📊")

archivo = st.file_uploader("Sube tu archivo de Horarios Excel", type=["xlsx"])
archivo_biometrico = st.file_uploader("Sube tu archivo Biométrico Excel", type=["xlsx"])

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

def a_timedelta(valor):
    valor = str(valor).strip()
    if valor in ['0', '0.0', '', 'nan', 'NaT']:
        return None
    for fmt in ['%H:%M:%S', '%H:%M']:
        try:
            t = pd.to_datetime(valor, format=fmt)
            return pd.Timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)
        except:
            continue
    return None

if archivo is not None:
    df = pd.read_excel(archivo)
    df["HORAS"] = df["HORAS"].str.replace("NO TIENE", "0")
    df = df[df["HORAS"] != "0"]
    df = df[df["NPLAN"] != 800]
    df["MATERIA_INI"] = pd.to_datetime(df["MATERIA_INI"], dayfirst=True)
    df["MATERIA_FIN"] = pd.to_datetime(df["MATERIA_FIN"], dayfirst=True)

    df["horarios_lista"] = df["HORAS"].apply(limpiar_horarios)
    df = df.explode("horarios_lista").reset_index(drop=True)
    horarios_df = pd.json_normalize(df["horarios_lista"]).reset_index(drop=True)
    df_final = pd.concat([df.drop(columns=["horarios_lista"]), horarios_df], axis=1)
    df_final = df_final.dropna(subset=["dia"])

    inicio_recargo_global = 780
    df_final['minutos_inicio'] = (pd.to_timedelta(df_final['hora_inicio'] + ':00') - pd.to_timedelta('06:00:00')).dt.total_seconds() / 60
    df_final['minutos_fin'] = (pd.to_timedelta(df_final['hora_fin'] + ':00') - pd.to_timedelta('06:00:00')).dt.total_seconds() / 60
    df_final['minutos_recargo'] = (
        df_final['minutos_fin'] - np.maximum(df_final['minutos_inicio'], inicio_recargo_global)
    ).clip(lower=0)

    df_final = df_final[df_final["minutos_recargo"] > 0]
    df_final = df_final.drop(columns=['MATERIA_ACTIVIDAD', 'TOTAL_HORAS', 'GRUPO'])

    dias_map = {
        "Monday": "LU", "Tuesday": "MA", "Wednesday": "MI",
        "Thursday": "JU", "Friday": "VI", "Saturday": "SA", "Sunday": "DO"
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
        lambda row: 0 if row["es_festivo"] else row["minutos_recargo"], axis=1
    )
    df_expandido["horas_recargo"] = df_expandido["minutos_recargo"] / 60

    final = df_expandido[df_expandido["horas_recargo"] > 0].copy()
    final = final[(final["fecha"] < "2026-03-29") | (final["fecha"] > "2026-04-05")]

    meses_espanol = {
        'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo',
        'April': 'Abril', 'May': 'Mayo', 'June': 'Junio',
        'July': 'Julio', 'August': 'Agosto', 'September': 'Septiembre',
        'October': 'Octubre', 'November': 'Noviembre', 'December': 'Diciembre'
    }
    final.insert(1, "mes", final["fecha"].dt.month_name().map(meses_espanol))
    final["mes"] = final["mes"].str.upper()

    df_agrupado = final.groupby(['DOCUMENTO', 'fecha']).agg(
        Entrada_Real=('hora_inicio', 'min'),
        Salida_Real=('hora_fin', 'max'),
        Suma_Recargos=('horas_recargo', 'sum')
    ).reset_index()

    llave_formateada = (df_agrupado['fecha'].dt.strftime('%d/%m/%Y') + '-' + df_agrupado['DOCUMENTO'].astype(str))
    df_agrupado.insert(0, 'llave', llave_formateada)
    df_agrupado = df_agrupado.sort_values(by=['DOCUMENTO', 'fecha'], ascending=[True, True])

    # ── Mostrar resultados de horarios ──
    st.subheader("Resultados detallados - Horarios")
    st.dataframe(final)

    # ════════════════════════════════════════
    # SECCIÓN BIOMÉTRICO (solo si se subió)
    # ════════════════════════════════════════
    if archivo_biometrico is not None:
        st.subheader("Cruce con Biométrico")

        biometrico = pd.read_excel(
            archivo_biometrico,
            usecols=["fecha", "Documento", "cargo", "hora_entrada", "hora_salida"],
            skiprows=1
        )

        # Llave
        biometrico['fecha'] = pd.to_datetime(biometrico['fecha'], dayfirst=True)
        llave_bio = biometrico['fecha'].dt.strftime('%d/%m/%Y') + '-' + biometrico['Documento'].astype(str)
        biometrico.insert(0, 'llave', llave_bio)

        # Merge con df_agrupado
        df_resultado = pd.merge(
            biometrico,
            df_agrupado[['llave', 'Entrada_Real', 'Salida_Real', 'Suma_Recargos']],
            on='llave',
            how='left'
        )
        df_resultado = df_resultado.rename(columns={
            'Entrada_Real': 'hora_inicio_clase',
            'Salida_Real': 'hora_fin_clase'
        })
        df_resultado = df_resultado.sort_values(by=['fecha', 'Documento'], ascending=[True, True])

        # Diferencia de horas
        # 1. Llenar nulos con 0 de inmediato y espacios con nan para hora salida
        df_resultado['hora_salida'] = df_resultado['hora_salida'].replace(" ",np.nan)
        mascara_sin_salida = df_resultado['hora_salida'].astype(str).str.strip().isin(['0', '0.0', '', 'nan'])
        mascara_sin_clase = df_resultado['hora_fin_clase'].astype(str).str.strip().isin(['0', '0.0', '', 'nan'])
        hora_salida_td    = df_resultado['hora_salida'].apply(a_timedelta)
        hora_fin_clase_td = df_resultado['hora_fin_clase'].apply(a_timedelta)

        # Crear nueva columna con la diferencia
        df_resultado['Diferencia'] = np.where(mascara_sin_clase | mascara_sin_salida,0,
                                              (hora_salida_td - hora_fin_clase_td).dt.total_seconds() / 3600)

        # Total horas reales a pagar
        df_resultado['total_horas'] = np.where(
        mascara_sin_salida,0, np.where(
            df_resultado['Diferencia'] < 0,
            df_resultado['Suma_Recargos'] + df_resultado['Diferencia'],
            df_resultado['Suma_Recargos']))
        df_resultado['total_horas'] = df_resultado['total_horas'].clip(lower=0)

        # Ordenamos el biométrico por llave y por total_horas de forma descendente
        df_resultado = df_resultado.sort_values(['llave', 'total_horas'], ascending=[True, False])

        # Se elimina el duplicado con menor recargo
        df_resultado = df_resultado.drop_duplicates(subset='llave', keep='first'

        st.dataframe(df_resultado)

        # Creación de nuevas columnas en el df agrupado
        
        # meses
        df_agrupado.insert(3, "mes", df_agrupado["fecha"].dt.month_name().map(meses_espanol))
        df_agrupado["mes"] = df_agrupado["mes"].str.upper()

        # Agregar hora de entrada y salida

        # 1. Realizamos el 'BuscarV' (merge)
        df_agrupado = pd.merge(
            df_agrupado, 
            biometrico[['llave', 'hora_entrada', 'hora_salida']], 
            on='llave', 
            how='left'
        )
        
        # 2. Renombramos las columnas para que queden exactamente como las pediste
        df_agrupado = df_agrupado.rename(columns={
            'hora_entrada': 'hora_inicio_labor',
            'hora_salida': 'hora_fin_labor'
        })
        
        # Ordenar por fecha (cronológico) y luego por documento
        df_agrupado = df_agrupado.sort_values(by=['DOCUMENTO', 'fecha'], ascending=[True, True])

        # Aplicar cambio de formato horas y calcular diferencia
        # Detectar filas sin clase y sin horario de salida 
        mascara_sin_clase = df_agrupado['Salida_Real'].astype(str).str.strip().isin(['0', '0.0', '', 'nan'])
        mascara_sin_salida = df_agrupado['hora_fin_labor'].astype(str).str.strip().isin(['0', '0.0', '', 'nan'])
        
        # Convertir a timedelta en columnas temporales
        hora_salida_td    = df_agrupado['hora_fin_labor'].apply(a_timedelta).fillna(pd.Timedelta(0))
        hora_fin_clase_td = df_agrupado['Salida_Real'].apply(a_timedelta).fillna(pd.Timedelta(0))
        
        # 2. Ahora la resta funcionará: (0:00:00 - 20:00:00) = -20 horas
        df_agrupado['Diferencia'] = (hora_salida_td - hora_fin_clase_td).dt.total_seconds() / 3600

        # AGREGAR COLUMNA DE TOTAL HORAS REALES A PAGAR 

        valor_base = np.where(df_agrupado['Diferencia'] < 0, 
                              df_agrupado['Suma_Recargos'] + df_agrupado['Diferencia'], 
                              df_agrupado['Suma_Recargos'])
        
        # Paso 2: Aplicamos el límite de 0 (equivalente al SI externo)
        df_agrupado['total_horas'] = np.where(valor_base < 0, 0, valor_base)

        # ── Descarga archivo biométrico ──
        output_bio = io.BytesIO()
        with pd.ExcelWriter(output_bio, engine='openpyxl') as writer:
            final.to_excel(writer, index=False, sheet_name="Horario")
            df_agrupado.to_excel(writer, index=False, sheet_name="horario_agrupado")
            df_resultado.to_excel(writer, index=False, sheet_name="cruce_biometrico")

        st.download_button(
            label="Descargar Excel con Cruce Biométrico 📥",
            data=output_bio.getvalue(),
            file_name="recargos_con_biometrico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Sube el archivo biométrico para generar el cruce.")




















