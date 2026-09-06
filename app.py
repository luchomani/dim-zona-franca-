import re
import pdfplumber
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Extractor de Declaraciones de Importación", layout="wide"
)

st.title("📦 Extractor Automatizado de Declaraciones de Importación")
st.markdown(
    "Procesa archivos **sueltos** o **fusionados** asegurando la extracción exacta de cada dato."
)


def extraer_texto_por_paginas(pdf_file):
    """Extrae el texto de cada página manteniendo la referencia de la página."""
    paginas_texto = []
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            texto = page.extract_text() or ""
            paginas_texto.append({"pagina": i + 1, "texto": texto})
    return paginas_texto


def segmentar_declaraciones(paginas_texto):
    """Divide el texto total en bloques individuales si hay múltiples declaraciones

    fusionadas en un solo PDF, o retorna un único bloque si es un archivo suelto.
    """
    texto_completo_unido = "\n".join([p["texto"] for p in paginas_texto])

    # Patrón común que indica el inicio de una Declaración de Importación en Colombia
    # Puedes ajustarlo según la etiqueta exacta que aparezca en tus formatos (ej. "Formulario 500" o "DECLARACIÓN DE IMPORTACIÓN")
    delimitador_patron = r"(?=DECLARACIÓN DE IMPORTACIÓN|FORMULARIO \d+|1\. Núm[ée]ro de Formulario)"

    # Dividir el texto usando el patrón conservando los separadores
    bloques = re.split(delimitador_patron, texto_completo_unido, flags=re.IGNORECASE)

    # Si la división no encuentra múltiples bloques, se toma todo el texto como una sola declaración
    if len(bloques) <= 1:
        return [texto_completo_unido]

    # Limpiar y filtrar bloques vacíos o muy cortos
    declaraciones_separadas = [
        b.strip() for b in bloques if b and len(b.strip()) > 50
    ]

    # Si por alguna razón el split no separó bien pero hay texto, devolver el texto unido
    if not declaraciones_separadas:
        return [texto_completo_unido]

    return declaraciones_separadas


def extraer_datos_formulario(texto_bloque):
    """Aplica expresiones regulares para extraer los campos críticos de cada declaración."""
    datos = {}

    # 1. Número de Formulario
    match_form = re.search(
        r"(?:Formulario|Número de Formulario|No\.)\s*[:\.]?\s*(\d+)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["Numero_Formulario"] = match_form.group(1) if match_form else "No encontrado"

    # 2. NIT del Importador
    match_nit = re.search(
        r"(?:NIT|Identificación)\s*[:\.]?\s*(\d{5,12}-?\d?)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["NIT_Importador"] = match_nit.group(1) if match_nit else "No encontrado"

    # 3. Tasa de Cambio
    match_tasa = re.search(
        r"Tasa de Cambio\s*[:\.]?\s*([\d\.,]+)", texto_bloque, re.IGNORECASE
    )
    datos["Tasa_Cambio"] = match_tasa.group(1) if match_tasa else "No encontrado"

    # 4. Valor FOB (USD)
    match_fob = re.search(
        r"FOB\s*(?:USD)?\s*[:\.]?\s*([\d\.,]+)", texto_bloque, re.IGNORECASE
    )
    datos["FOB_USD"] = match_fob.group(1) if match_fob else "No encontrado"

    # 5. Total Liquidación / Total a Pagar
    match_total = re.search(
        r"(?:Total Liquidación|Total a Pagar)\s*[:\.]?\s*([\d\.,]+)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["Total_Liquidacion"] = (
        match_total.group(1) if match_total else "No encontrado"
    )

    return datos


# Interfaz de Streamlit para la carga de archivos
uploaded_files = st.file_uploader(
    "Sube tus declaraciones de importación (Sueltas o Fusionadas en PDF)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    resultados_totales = []

    with st.spinner("Procesando documentos y extrayendo datos..."):
        for uploaded_file in uploaded_files:
            # Extraer páginas del PDF
            paginas = extraer_texto_por_paginas(uploaded_file)

            # Segmentar si vienen varias declaraciones juntas (fusionadas)
            bloques_declaraciones = segmentar_declaraciones(paginas)

            # Extraer datos de cada bloque detectado
            for idx, bloque in enumerate(bloques_declaraciones):
                info_extraida = extraer_datos_formulario(bloque)
                info_extraida["Archivo_Origen"] = uploaded_file.name
                info_extraida["Bloque_ID"] = idx + 1
                resultados_totales.append(info_extraida)

    # Convertir a DataFrame para visualización profesional
    df_resultados = pd.DataFrame(resultados_totales)

    st.success(
        f"¡Proceso completado! Se procesaron {len(uploaded_files)} archivo(s) generando {len(df_resultados)} registro(s)."
    )

    # Mostrar tabla interactiva
    st.subheader("📊 Datos Extraídos Correctamente")
    st.dataframe(df_resultados, use_container_width=True)

    # Opción de descarga en CSV
    csv = df_resultados.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Descargar Resultados en CSV",
        data=csv,
        file_name="declaraciones_extraidas.csv",
        mime="text/csv",
    )
