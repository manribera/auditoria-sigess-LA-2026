import streamlit as st
import pandas as pd
import io
import re
import unicodedata

st.set_page_config(
    page_title="Auditoría SIGESS 2026",
    layout="wide"
)

col_logo, col_titulo = st.columns([1, 6])

with col_logo:
    st.image("man.png", width=90)

with col_titulo:
    st.title("Auditoría de Libros SIGESS 2026")
st.write(
    "Compara un **Libro Base/Original** contra un **Libro Final 2026** "
    "detectando líneas, indicadores y metas eliminadas, nuevas o modificadas."
)

# =====================================================
# CONFIGURACIÓN
# =====================================================

HOJA_SIGESS = "Informe de avance"

# Esta será la clave principal de comparación.
# Para SIGESS se recomienda usar este ID generado.
COLUMNA_CLAVE = "ID_REGISTRO"


# =====================================================
# FUNCIONES DE LIMPIEZA
# =====================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def quitar_tildes(texto):
    texto = limpiar_texto(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def normalizar_lider(valor):
    valor_original = limpiar_texto(valor)
    valor = quitar_tildes(valor_original)

    if "municipal" in valor or "gobierno local" in valor or valor == "gl":
        return "Gobierno Local"

    if "fuerza" in valor or valor == "fp":
        return "Fuerza Pública"

    return valor_original


def extraer_numero_linea(texto):
    texto = limpiar_texto(texto)
    match = re.search(r"#\s*(\d+)", texto)
    if match:
        return int(match.group(1))
    return ""


def crear_id_registro(delegacion, numero_linea, numero_indicador):
    delegacion = re.sub(r"\s+", "_", limpiar_texto(delegacion))
    delegacion = re.sub(r"[^A-Za-z0-9_ÁÉÍÓÚáéíóúÑñ]", "", delegacion)

    return f"{delegacion}_L{numero_linea}_I{numero_indicador}"


# =====================================================
# EXTRACCIÓN SIGESS
# =====================================================

def extraer_planificacion_sigess(archivo):
    resultados = []

    xls = pd.ExcelFile(archivo, engine="openpyxl")

    if HOJA_SIGESS not in xls.sheet_names:
        raise ValueError(f"El archivo no tiene la hoja '{HOJA_SIGESS}'.")

    df = pd.read_excel(
        xls,
        sheet_name=HOJA_SIGESS,
        header=None,
        engine="openpyxl"
    )

    # Delegación normalmente ubicada en fila 3, columna H
    delegacion = limpiar_texto(df.iloc[2, 7]) if df.shape[0] > 2 and df.shape[1] > 7 else ""

    for i in range(len(df)):
        fila = df.iloc[i]

        texto_linea = limpiar_texto(fila[3]) if len(fila) > 3 else ""
        texto_linea_normalizado = quitar_tildes(texto_linea)

        if texto_linea_normalizado.startswith("linea de accion"):

            numero_linea = extraer_numero_linea(texto_linea)
            problematica = limpiar_texto(fila[5]) if len(fila) > 5 else ""
            lider_bloque = normalizar_lider(fila[7]) if len(fila) > 7 else ""

            fila_inicio_indicadores = i + 4

            # Buscar hasta antes de la siguiente línea de acción
            fila_fin_bloque = len(df)

            for k in range(i + 1, len(df)):
                posible_siguiente = limpiar_texto(df.iloc[k, 3]) if df.shape[1] > 3 else ""

                if quitar_tildes(posible_siguiente).startswith("linea de accion"):
                    fila_fin_bloque = k
                    break

            numero_indicador_real = 1

            for j in range(fila_inicio_indicadores, fila_fin_bloque):
                fila_ind = df.iloc[j]

                sigla_lider = limpiar_texto(fila_ind[2]) if len(fila_ind) > 2 else ""
                responsable = limpiar_texto(fila_ind[3]) if len(fila_ind) > 3 else ""
                indicador_num = limpiar_texto(fila_ind[4]) if len(fila_ind) > 4 else ""
                indicador = limpiar_texto(fila_ind[5]) if len(fila_ind) > 5 else ""
                meta = limpiar_texto(fila_ind[7]) if len(fila_ind) > 7 else ""

                # Evita filas vacías o encabezados
                if not indicador and not meta:
                    continue

                if quitar_tildes(indicador).startswith("indicador"):
                    continue

                if quitar_tildes(meta).startswith("meta"):
                    continue

                lider_estrategico = lider_bloque

                if not lider_estrategico:
                    lider_estrategico = normalizar_lider(responsable or sigla_lider)

                id_registro = crear_id_registro(
                    delegacion,
                    numero_linea,
                    numero_indicador_real
                )

                resultados.append({
                    "ID_REGISTRO": id_registro,
                    "Delegación Policial": delegacion,
                    "Número de Línea": numero_linea,
                    "Problemática": problematica,
                    "Líder Estratégico": lider_estrategico,
                    "Responsable": responsable,
                    "Indicador Número": indicador_num,
                    "Indicador": indicador,
                    "Meta": meta,
                    "Estado Base": "Activa"
                })

                numero_indicador_real += 1

    return pd.DataFrame(resultados)


# =====================================================
# COMPARACIÓN
# =====================================================

def comparar_libros(df_base, df_final):
    claves_base = set(df_base[COLUMNA_CLAVE])
    claves_final = set(df_final[COLUMNA_CLAVE])

    claves_eliminadas = claves_base - claves_final
    claves_nuevas = claves_final - claves_base
    claves_comunes = claves_base.intersection(claves_final)

    eliminados = df_base[df_base[COLUMNA_CLAVE].isin(claves_eliminadas)].copy()
    nuevos = df_final[df_final[COLUMNA_CLAVE].isin(claves_nuevas)].copy()

    modificaciones = []

    columnas_comparar = [
        "Delegación Policial",
        "Número de Línea",
        "Problemática",
        "Líder Estratégico",
        "Responsable",
        "Indicador Número",
        "Indicador",
        "Meta"
    ]

    base_index = df_base.set_index(COLUMNA_CLAVE)
    final_index = df_final.set_index(COLUMNA_CLAVE)

    for clave in claves_comunes:
        fila_base = base_index.loc[clave]
        fila_final = final_index.loc[clave]

        if isinstance(fila_base, pd.DataFrame):
            fila_base = fila_base.iloc[0]

        if isinstance(fila_final, pd.DataFrame):
            fila_final = fila_final.iloc[0]

        for col in columnas_comparar:
            valor_base = limpiar_texto(fila_base.get(col, ""))
            valor_final = limpiar_texto(fila_final.get(col, ""))

            if valor_base != valor_final:
                modificaciones.append({
                    "ID_REGISTRO": clave,
                    "Campo Modificado": col,
                    "Valor Libro Base": valor_base,
                    "Valor Libro Final 2026": valor_final
                })

    df_modificaciones = pd.DataFrame(modificaciones)

    return eliminados, nuevos, df_modificaciones


def generar_excel_reporte(df_base, df_final, eliminados, nuevos, modificaciones, resumen):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
        df_base.to_excel(writer, index=False, sheet_name="BASE_EXTRAIDA")
        df_final.to_excel(writer, index=False, sheet_name="FINAL_EXTRAIDO")
        eliminados.to_excel(writer, index=False, sheet_name="ELIMINADOS")
        nuevos.to_excel(writer, index=False, sheet_name="NUEVOS")
        modificaciones.to_excel(writer, index=False, sheet_name="MODIFICACIONES")

    return output.getvalue()


# =====================================================
# INTERFAZ
# =====================================================

col1, col2 = st.columns(2)

with col1:
    archivo_base = st.file_uploader(
        "📘 Cargar Libro Base / Original",
        type=["xlsx", "xlsm"],
        key="base"
    )

with col2:
    archivo_final = st.file_uploader(
        "📗 Cargar Libro Final 2026",
        type=["xlsx", "xlsm"],
        key="final"
    )

ejecutar = st.button("🚀 Ejecutar Auditoría SIGESS", type="primary")


# =====================================================
# EJECUCIÓN
# =====================================================

if ejecutar:

    if not archivo_base or not archivo_final:
        st.warning("Debes cargar ambos libros para ejecutar la auditoría.")
        st.stop()

    try:
        df_base = extraer_planificacion_sigess(archivo_base)
        df_final = extraer_planificacion_sigess(archivo_final)

        if df_base.empty:
            st.error("No se extrajo información válida del Libro Base.")
            st.stop()

        if df_final.empty:
            st.error("No se extrajo información válida del Libro Final 2026.")
            st.stop()

        eliminados, nuevos, modificaciones = comparar_libros(df_base, df_final)

        total_base = len(df_base)
        total_final = len(df_final)
        total_eliminados = len(eliminados)
        total_nuevos = len(nuevos)
        total_modificados = modificaciones["ID_REGISTRO"].nunique() if not modificaciones.empty else 0

        claves_base = set(df_base[COLUMNA_CLAVE])
        claves_final = set(df_final[COLUMNA_CLAVE])
        total_comunes = len(claves_base.intersection(claves_final))
        total_iguales = total_comunes - total_modificados

        resumen = pd.DataFrame([{
            "Total indicadores Libro Base": total_base,
            "Total indicadores Libro Final 2026": total_final,
            "Indicadores iguales": total_iguales,
            "Indicadores eliminados": total_eliminados,
            "Indicadores nuevos": total_nuevos,
            "Indicadores modificados": total_modificados
        }])

        st.success("Auditoría ejecutada correctamente.")

        m1, m2, m3, m4, m5, m6 = st.columns(6)

        m1.metric("Base", total_base)
        m2.metric("Final 2026", total_final)
        m3.metric("Iguales", total_iguales)
        m4.metric("Eliminados ❌", total_eliminados)
        m5.metric("Nuevos ➕", total_nuevos)
        m6.metric("Modificados ⚠️", total_modificados)

        if total_eliminados > 0:
            st.error(f"Se detectaron {total_eliminados} indicadores/líneas eliminadas del Libro Final 2026.")

        if total_nuevos > 0:
            st.info(f"Se detectaron {total_nuevos} indicadores/líneas nuevas agregadas.")

        if total_modificados > 0:
            st.warning(f"Se detectaron modificaciones en {total_modificados} registros existentes.")

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Eliminados sospechosos ❌",
            "Nuevos agregados ➕",
            "Modificaciones ⚠️",
            "Base extraída 📘",
            "Final extraído 📗"
        ])

        with tab1:
            st.subheader("❌ Indicadores o líneas eliminadas")
            if eliminados.empty:
                st.success("No se detectaron eliminaciones.")
            else:
                st.dataframe(
                    eliminados.style.apply(
                        lambda row: ["background-color: #ffcccc"] * len(row),
                        axis=1
                    ),
                    use_container_width=True
                )

        with tab2:
            st.subheader("➕ Indicadores o líneas nuevas")
            if nuevos.empty:
                st.success("No se detectaron nuevos registros.")
            else:
                st.dataframe(
                    nuevos.style.apply(
                        lambda row: ["background-color: #d9ead3"] * len(row),
                        axis=1
                    ),
                    use_container_width=True
                )

        with tab3:
            st.subheader("⚠️ Modificaciones detectadas")
            if modificaciones.empty:
                st.success("No se detectaron modificaciones.")
            else:
                st.dataframe(
                    modificaciones.style.apply(
                        lambda row: ["background-color: #fff2cc"] * len(row),
                        axis=1
                    ),
                    use_container_width=True
                )

        with tab4:
            st.subheader("📘 Planificación extraída del Libro Base")
            st.dataframe(df_base, use_container_width=True)

        with tab5:
            st.subheader("📗 Planificación extraída del Libro Final 2026")
            st.dataframe(df_final, use_container_width=True)

        reporte = generar_excel_reporte(
            df_base,
            df_final,
            eliminados,
            nuevos,
            modificaciones,
            resumen
        )

        st.download_button(
            label="📥 Descargar Reporte de Auditoría SIGESS",
            data=reporte,
            file_name="REPORTE_AUDITORIA_SIGESS_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error durante la auditoría: {e}")
