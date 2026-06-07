import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from datetime import date
from fpdf import FPDF

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================

st.set_page_config(
    page_title="Seguimiento SIGESS 2026",
    layout="wide"
)

HOJA_SIGESS = "Informe de avance"
COLUMNA_CLAVE = "ID_REGISTRO"


# =====================================================
# ENCABEZADO
# =====================================================

col_logo, col_titulo = st.columns([1, 6])

with col_logo:
    try:
        st.image("man.png", width=90)
    except Exception:
        st.write("")

with col_titulo:
    st.title("Seguimiento Comparativo SIGESS 2026")

st.write(
    "Herramienta para verificar la correspondencia metodológica entre el "
    "**Libro Base 2025** y el **Informe Trimestral de Avance 2026**, "
    "considerando líneas de acción, indicadores, metas y campos clave."
)


# =====================================================
# PANEL EDITABLE DEL INFORME
# =====================================================

with st.sidebar:
    st.header("Datos del informe")

    elaborado_por = st.text_input(
        "Realizado por",
        value="",
        placeholder="Nombre de quien realiza el informe"
    )

    fecha_emision = st.date_input(
        "Fecha de emisión",
        value=date.today()
    )

    st.header("Textos editables del PDF")

    titulo_pdf = st.text_input(
        "Título del informe",
        value="Informe de Seguimiento Comparativo de Líneas de Coordinación Estratégica"
    )

    subtitulo_pdf = st.text_input(
        "Subtítulo",
        value="Verificación de correspondencia metodológica entre el Libro Base 2025 y el Informe Trimestral de Avance 2026"
    )

    objeto_analisis = st.text_area(
        "Objeto del análisis",
        value=(
            "El presente informe tiene como finalidad verificar la correspondencia metodológica "
            "entre el Libro Base 2025 utilizado como referencia y el Informe Trimestral de Avance "
            "evaluado, considerando las líneas de acción, indicadores, metas y demás elementos "
            "asociados a la planificación estratégica."
        ),
        height=150
    )

    alcance_metodologico = st.text_area(
        "Alcance metodológico",
        value=(
            "La revisión se realiza mediante una comparación estructurada entre ambos instrumentos, "
            "identificando coincidencias, registros presentes en el Libro Base no localizados en el "
            "instrumento evaluado, registros incorporados y variaciones en campos clave. "
            "El resultado constituye un insumo técnico de seguimiento y no representa una auditoría "
            "ni una valoración disciplinaria."
        ),
        height=170
    )

    texto_valoracion = st.text_area(
        "Valoración técnica",
        value=(
            "Los resultados obtenidos permiten observar el grado de correspondencia entre la "
            "planificación base y el instrumento trimestral evaluado. Las variaciones identificadas "
            "deben ser revisadas en función del alcance metodológico definido para el seguimiento "
            "de las Líneas de Coordinación Estratégica."
        ),
        height=150
    )

    fuente_pdf = st.text_area(
        "Fuente",
        value=(
            "Fuente: Libro Base 2025 e Informe Trimestral de Avance de Líneas de Coordinación Estratégica 2026."
        ),
        height=80
    )

datos_pdf = {
    "elaborado_por": elaborado_por,
    "fecha_emision": fecha_emision.strftime("%d/%m/%Y"),
    "titulo_pdf": titulo_pdf,
    "subtitulo_pdf": subtitulo_pdf,
    "objeto_analisis": objeto_analisis,
    "alcance_metodologico": alcance_metodologico,
    "texto_valoracion": texto_valoracion,
    "fuente_pdf": fuente_pdf,
}


# =====================================================
# FUNCIONES DE LIMPIEZA Y APOYO
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


def nombre_archivo_seguro(texto):
    texto = limpiar_texto(texto)
    texto = re.sub(r"[^\w\s-]", "", texto)
    texto = re.sub(r"\s+", "_", texto)
    return texto.upper() if texto else "DELEGACION"


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

    delegacion = (
        limpiar_texto(df.iloc[2, 7])
        if df.shape[0] > 2 and df.shape[1] > 7
        else ""
    )

    for i in range(len(df)):
        fila = df.iloc[i]

        texto_linea = limpiar_texto(fila[3]) if len(fila) > 3 else ""
        texto_linea_normalizado = quitar_tildes(texto_linea)

        if texto_linea_normalizado.startswith("linea de accion"):

            numero_linea = extraer_numero_linea(texto_linea)
            problematica = limpiar_texto(fila[5]) if len(fila) > 5 else ""
            lider_bloque = normalizar_lider(fila[7]) if len(fila) > 7 else ""

            fila_inicio_indicadores = i + 4

            fila_fin_bloque = len(df)

            for k in range(i + 1, len(df)):
                posible_siguiente = (
                    limpiar_texto(df.iloc[k, 3])
                    if df.shape[1] > 3
                    else ""
                )

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

    claves_no_localizadas = claves_base - claves_final
    claves_incorporadas = claves_final - claves_base
    claves_comunes = claves_base.intersection(claves_final)

    no_localizados = df_base[
        df_base[COLUMNA_CLAVE].isin(claves_no_localizadas)
    ].copy()

    incorporados = df_final[
        df_final[COLUMNA_CLAVE].isin(claves_incorporadas)
    ].copy()

    variaciones = []

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
                variaciones.append({
                    "ID_REGISTRO": clave,
                    "Campo con variación": col,
                    "Valor Libro Base 2025": valor_base,
                    "Valor Informe Evaluado 2026": valor_final
                })

    return no_localizados, incorporados, pd.DataFrame(variaciones)
# =====================================================
# REPORTE EXCEL
# =====================================================

def generar_excel_reporte(
    df_base,
    df_final,
    no_localizados,
    incorporados,
    variaciones,
    resumen
):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
        df_base.to_excel(writer, index=False, sheet_name="BASE_2025_EXTRAIDA")
        df_final.to_excel(writer, index=False, sheet_name="INFORME_2026_EXTRAIDO")
        no_localizados.to_excel(writer, index=False, sheet_name="NO_LOCALIZADOS")
        incorporados.to_excel(writer, index=False, sheet_name="INCORPORADOS")
        variaciones.to_excel(writer, index=False, sheet_name="VARIACIONES")

        workbook = writer.book
        formato_header = workbook.add_format({
            "bold": True,
            "bg_color": "#1F4E79",
            "font_color": "white",
            "border": 1
        })

        hojas = {
            "RESUMEN": resumen,
            "BASE_2025_EXTRAIDA": df_base,
            "INFORME_2026_EXTRAIDO": df_final,
            "NO_LOCALIZADOS": no_localizados,
            "INCORPORADOS": incorporados,
            "VARIACIONES": variaciones,
        }

        for nombre_hoja, dataframe in hojas.items():
            ws = writer.sheets[nombre_hoja]
            ws.freeze_panes(1, 0)

            if not dataframe.empty:
                ws.autofilter(0, 0, len(dataframe), len(dataframe.columns) - 1)

            for col_num, value in enumerate(dataframe.columns.values):
                ws.write(0, col_num, value, formato_header)
                ws.set_column(col_num, col_num, 28)

    return output.getvalue()


# =====================================================
# REPORTE PDF
# =====================================================

class PDFSeguimiento(FPDF):
    def header(self):
        self.set_fill_color(11, 61, 78)
        self.rect(0, 0, 210, 14, "F")
        self.set_y(16)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")


def limpiar_pdf(texto):
    texto = limpiar_texto(texto)

    reemplazos = {
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "•": "-",
    }

    for viejo, nuevo in reemplazos.items():
        texto = texto.replace(viejo, nuevo)

    return texto


def agregar_titulo_seccion(pdf, titulo):
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(11, 61, 78)
    pdf.multi_cell(0, 8, limpiar_pdf(titulo))
    pdf.ln(3)


def agregar_parrafo(pdf, texto):
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(40, 40, 40)
    pdf.multi_cell(0, 6, limpiar_pdf(texto))
    pdf.ln(4)


def agregar_tabla_simple(pdf, df, columnas, max_filas=12):
    columnas_existentes = [c for c in columnas if c in df.columns]

    if df.empty or not columnas_existentes:
        pdf.set_font("Helvetica", "I", 10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(0, 8, "No se registran datos en esta sección.", ln=True)
        pdf.ln(4)
        return

    df_mostrar = df[columnas_existentes].head(max_filas).copy()

    ancho_total = 190
    ancho_col = ancho_total / len(columnas_existentes)

    pdf.set_font("Helvetica", "B", 7)
    pdf.set_fill_color(11, 61, 78)
    pdf.set_text_color(255, 255, 255)

    for col in columnas_existentes:
        pdf.cell(ancho_col, 7, limpiar_pdf(str(col))[:24], border=1, fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 7)
    pdf.set_text_color(40, 40, 40)

    for _, row in df_mostrar.iterrows():
        for col in columnas_existentes:
            valor = limpiar_pdf(row.get(col, ""))
            pdf.cell(ancho_col, 7, valor[:35], border=1)
        pdf.ln()

    if len(df) > max_filas:
        pdf.ln(2)
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(
            0,
            6,
            f"Se muestran {max_filas} registros de {len(df)}. Ver detalle completo en el Excel anexo.",
            ln=True
        )

    pdf.ln(5)


def agregar_resumen_en_pdf(pdf, resumen):
    columnas = list(resumen.columns)
    agregar_tabla_simple(pdf, resumen, columnas, max_filas=5)


def generar_pdf_seguimiento(
    datos_pdf,
    resumen,
    df_base,
    df_final,
    no_localizados,
    incorporados,
    variaciones
):
    pdf = PDFSeguimiento()
    pdf.set_auto_page_break(auto=True, margin=18)

    total_no_localizados = int(resumen["Registros no localizados"].iloc[0])
    total_incorporados = int(resumen["Registros incorporados"].iloc[0])
    total_variaciones = int(resumen["Registros con variaciones"].iloc[0])

    informe_simplificado = (
        total_no_localizados == 0
        and total_incorporados == 0
        and total_variaciones == 0
    )

    # =====================================================
    # PORTADA
    # =====================================================

    pdf.add_page()
    pdf.ln(18)

    try:
        pdf.image("logo.png", x=80, y=30, w=50)
        pdf.ln(55)
    except Exception:
        pdf.ln(25)

    pdf.set_font("Helvetica", "B", 19)
    pdf.set_text_color(11, 61, 78)
    pdf.multi_cell(0, 10, limpiar_pdf(datos_pdf["titulo_pdf"]), align="C")

    pdf.ln(4)

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(70, 70, 70)
    pdf.multi_cell(0, 7, limpiar_pdf(datos_pdf["subtitulo_pdf"]), align="C")

    pdf.ln(18)

    delegacion = ""
    if not df_final.empty and "Delegación Policial" in df_final.columns:
        delegacion = limpiar_texto(df_final["Delegación Policial"].iloc[0])

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(0, 8, f"Delegación: {limpiar_pdf(delegacion)}", ln=True)
    pdf.cell(0, 8, f"Fecha de emisión: {datos_pdf['fecha_emision']}", ln=True)
    pdf.cell(0, 8, f"Realizado por: {limpiar_pdf(datos_pdf['elaborado_por'])}", ln=True)

    pdf.ln(20)

    pdf.set_font("Helvetica", "I", 10)
    pdf.set_fill_color(235, 242, 245)
    pdf.multi_cell(
        0,
        7,
        limpiar_pdf(
            "Documento técnico generado para apoyar el seguimiento comparativo "
            "de la estructura de líneas de coordinación estratégica."
        ),
        border=1,
        align="C",
        fill=True
    )

    # =====================================================
    # OBJETO Y ALCANCE
    # =====================================================

    pdf.add_page()
    agregar_titulo_seccion(pdf, "1. Objeto del análisis")
    agregar_parrafo(pdf, datos_pdf["objeto_analisis"])

    agregar_titulo_seccion(pdf, "2. Alcance metodológico")
    agregar_parrafo(pdf, datos_pdf["alcance_metodologico"])

    # =====================================================
    # RESUMEN EJECUTIVO
    # =====================================================

    pdf.add_page()
    agregar_titulo_seccion(pdf, "3. Resumen ejecutivo")

    total_base = int(resumen["Total registros Libro Base 2025"].iloc[0])
    total_final = int(resumen["Total registros Informe Evaluado 2026"].iloc[0])
    total_coincidentes = int(resumen["Registros coincidentes"].iloc[0])

    if informe_simplificado:
        texto_resumen = (
            f"Se procesaron {total_base} registros del Libro Base 2025 y "
            f"{total_final} registros del Informe Trimestral de Avance 2026. "
            f"El seguimiento comparativo no identificó registros no localizados, "
            f"registros incorporados ni variaciones en los campos clave evaluados. "
            f"En consecuencia, los {total_coincidentes} registros comparados mantienen "
            f"correspondencia metodológica entre ambos instrumentos."
        )
    else:
        texto_resumen = (
            f"Se procesaron {total_base} registros del Libro Base 2025 y "
            f"{total_final} registros del Informe Trimestral de Avance 2026. "
            f"Como resultado, se identificaron {total_coincidentes} registros coincidentes, "
            f"{total_no_localizados} registros no localizados en el instrumento evaluado, "
            f"{total_incorporados} registros incorporados y "
            f"{total_variaciones} registros con variaciones en campos clave."
        )

    agregar_parrafo(pdf, texto_resumen)
    agregar_resumen_en_pdf(pdf, resumen)

    # =====================================================
    # CUERPO DEL INFORME
    # =====================================================

    if informe_simplificado:
        pdf.add_page()
        agregar_titulo_seccion(pdf, "4. Resultado del seguimiento comparativo")
        agregar_parrafo(
            pdf,
            "Con base en la información extraída de ambos instrumentos, no se observaron "
            "diferencias en las líneas de acción, indicadores, metas y campos clave evaluados. "
            "El Informe Trimestral de Avance 2026 mantiene correspondencia con el Libro Base "
            "2025 utilizado como referencia metodológica."
        )

        agregar_titulo_seccion(pdf, "5. Valoración técnica")
        agregar_parrafo(
            pdf,
            "El resultado obtenido permite considerar que, para los campos evaluados por la "
            "herramienta, ambos instrumentos reúnen las mismas condiciones de estructura y lógica "
            "metodológica. Esta valoración se limita al alcance comparativo definido y constituye "
            "un insumo de seguimiento técnico."
        )

        agregar_titulo_seccion(pdf, "6. Fuente")
        agregar_parrafo(pdf, datos_pdf["fuente_pdf"])

    else:
        pdf.add_page()
        agregar_titulo_seccion(pdf, "4. Registros del Libro Base no localizados")
        agregar_parrafo(
            pdf,
            "Esta sección muestra registros presentes en el Libro Base 2025 que no fueron localizados "
            "en el Informe Trimestral de Avance evaluado."
        )
        agregar_tabla_simple(
            pdf,
            no_localizados,
            ["Número de Línea", "Indicador", "Meta"],
            max_filas=15
        )

        pdf.add_page()
        agregar_titulo_seccion(pdf, "5. Registros incorporados en el instrumento evaluado")
        agregar_parrafo(
            pdf,
            "Esta sección muestra registros presentes en el Informe Trimestral de Avance 2026 "
            "que no se encontraban en el Libro Base 2025 utilizado como referencia."
        )
        agregar_tabla_simple(
            pdf,
            incorporados,
            ["Número de Línea", "Indicador", "Meta"],
            max_filas=15
        )

        pdf.add_page()
        agregar_titulo_seccion(pdf, "6. Variaciones identificadas")
        agregar_parrafo(
            pdf,
            "Esta sección muestra diferencias observadas en campos clave de registros localizados "
            "en ambos instrumentos."
        )
        agregar_tabla_simple(
            pdf,
            variaciones,
            [
                "ID_REGISTRO",
                "Campo con variación",
                "Valor Libro Base 2025",
                "Valor Informe Evaluado 2026"
            ],
            max_filas=15
        )

        pdf.add_page()
        agregar_titulo_seccion(pdf, "7. Valoración técnica")
        agregar_parrafo(pdf, datos_pdf["texto_valoracion"])

        agregar_titulo_seccion(pdf, "8. Fuente")
        agregar_parrafo(pdf, datos_pdf["fuente_pdf"])

    # =====================================================
    # PÁGINA DE CIERRE
    # =====================================================

    pdf.add_page()
    pdf.ln(55)

    try:
        pdf.image("logo.png", x=70, y=80, w=70)
    except Exception:
        pass

    pdf.ln(90)

    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(11, 61, 78)
    pdf.multi_cell(
        0,
        9,
        "Coordinación Nacional\nEstrategia Sembremos Seguridad 2026",
        align="C"
    )

    pdf_bytes = bytes(pdf.output(dest="S"))
    return pdf_bytes


# =====================================================
# INTERFAZ DE CARGA
# =====================================================
if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

if st.button("Limpiar archivos cargados"):
    st.session_state["uploader_version"] += 1
    st.rerun()

col1, col2 = st.columns(2)

with col1:
    archivo_base = st.file_uploader(
        "📘 Cargar Libro Base 2025",
        type=["xlsx", "xlsm"],
        key=f"base_{st.session_state['uploader_version']}"
    )

with col2:
    archivo_final = st.file_uploader(
        "📗 Cargar Informe Trimestral 2026",
        type=["xlsx", "xlsm"],
        key=f"final_{st.session_state['uploader_version']}"
    )
    
ejecutar = st.button("Ejecutar Seguimiento Comparativo", type="primary")

# =====================================================
# EJECUCIÓN
# =====================================================

if ejecutar:

    if not archivo_base or not archivo_final:
        st.warning("Debes cargar ambos libros para ejecutar el seguimiento.")
        st.stop()

    try:
        df_base = extraer_planificacion_sigess(archivo_base)
        df_final = extraer_planificacion_sigess(archivo_final)

        delegacion_reporte = "DELEGACION"

        if not df_final.empty and "Delegación Policial" in df_final.columns:
            delegacion_reporte = nombre_archivo_seguro(
                df_final["Delegación Policial"].iloc[0]
            )

        if df_base.empty:
            st.error("No se extrajo información válida del Libro Base 2025.")
            st.stop()

        if df_final.empty:
            st.error("No se extrajo información válida del Informe Trimestral 2026.")
            st.stop()

        no_localizados, incorporados, variaciones = comparar_libros(df_base, df_final)

        total_base = len(df_base)
        total_final = len(df_final)
        total_no_localizados = len(no_localizados)
        total_incorporados = len(incorporados)
        total_variaciones = (
            variaciones["ID_REGISTRO"].nunique()
            if not variaciones.empty
            else 0
        )

        claves_base = set(df_base[COLUMNA_CLAVE])
        claves_final = set(df_final[COLUMNA_CLAVE])
        total_comunes = len(claves_base.intersection(claves_final))
        total_coincidentes = total_comunes - total_variaciones

        resumen = pd.DataFrame([{
            "Total registros Libro Base 2025": total_base,
            "Total registros Informe Evaluado 2026": total_final,
            "Registros coincidentes": total_coincidentes,
            "Registros no localizados": total_no_localizados,
            "Registros incorporados": total_incorporados,
            "Registros con variaciones": total_variaciones
        }])

        st.success("Seguimiento comparativo ejecutado correctamente.")

        m1, m2, m3, m4, m5, m6 = st.columns(6)

        m1.metric("Base 2025", total_base)
        m2.metric("Informe 2026", total_final)
        m3.metric("Coincidentes", total_coincidentes)
        m4.metric("No localizados", total_no_localizados)
        m5.metric("Incorporados", total_incorporados)
        m6.metric("Con variaciones", total_variaciones)

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "No localizados",
            "Incorporados",
            "Variaciones",
            "Base 2025 extraída",
            "Informe 2026 extraído"
        ])

        with tab1:
            st.subheader("Registros del Libro Base no localizados")
            if no_localizados.empty:
                st.success("No se registran elementos no localizados.")
            else:
                st.dataframe(no_localizados, use_container_width=True)

        with tab2:
            st.subheader("Registros incorporados")
            if incorporados.empty:
                st.success("No se registran elementos incorporados.")
            else:
                st.dataframe(incorporados, use_container_width=True)

        with tab3:
            st.subheader("Variaciones identificadas")
            if variaciones.empty:
                st.success("No se registran variaciones.")
            else:
                st.dataframe(variaciones, use_container_width=True)

        with tab4:
            st.subheader("Planificación extraída del Libro Base 2025")
            st.dataframe(df_base, use_container_width=True)

        with tab5:
            st.subheader("Planificación extraída del Informe Trimestral 2026")
            st.dataframe(df_final, use_container_width=True)

        reporte_excel = generar_excel_reporte(
            df_base,
            df_final,
            no_localizados,
            incorporados,
            variaciones,
            resumen
        )

        reporte_pdf = generar_pdf_seguimiento(
            datos_pdf,
            resumen,
            df_base,
            df_final,
            no_localizados,
            incorporados,
            variaciones
        )

        col_descarga1, col_descarga2 = st.columns(2)

        with col_descarga1:
            st.download_button(
                label="📥 Descargar reporte Excel",
                data=reporte_excel,
                file_name=f"REPORTE_SEGUIMIENTO_{delegacion_reporte}_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_descarga2:
            st.download_button(
                label="📄 Descargar reporte PDF",
                data=reporte_pdf,
                file_name=f"INFORME_SEGUIMIENTO_{delegacion_reporte}_2026.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error durante el seguimiento comparativo: {e}")
