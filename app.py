import fitz  # PyMuPDF
import pandas as pd
import re
import streamlit as st

st.set_page_config(
    page_title="Extractor de Declaraciones de Importación", layout="wide"
)

st.title("📦 Extractor Automatizado de Declaraciones de Importación")
st.markdown(
    "Procesa archivos **sueltos** o **fusionados** asegurando la extracción"
    " exacta de cada dato clave."
)


def extraer_texto_por_paginas(pdf_file):
    """Extrae el texto de cada página utilizando PyMuPDF (fitz) desde los bytes."""
    paginas_texto = []
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    for i, page in enumerate(doc):
        texto = page.get_text() or ""
        paginas_texto.append({"pagina": i + 1, "texto": texto})
    return paginas_texto


def segmentar_declaraciones(paginas_texto):
    """Divide inteligentemente el texto en bloques individuales si hay múltiples

    declaraciones fusionadas en un solo PDF, o maneja un archivo suelto.
    """
    texto_completo_unido = "\n".join([p["texto"] for p in paginas_texto])

    # Patrón robusto para detectar el inicio de una declaración de importación
    delimitador_patron = r"(?=Declaraci[oó]n de Importaci[oó]n|Formulario|\b89202\d{10})"

    bloques = re.split(delimitador_patron, texto_completo_unido, flags=re.IGNORECASE)

    if len(bloques) <= 1:
        return [texto_completo_unido]

    declaraciones_separadas = [
        b.strip() for b in bloques if b and len(b.strip()) > 50
    ]

    return (
        declaraciones_separadas
        if declaraciones_separadas
        else [texto_completo_unido]
    )


def extraer_datos_formulario(texto_bloque):
    """Aplica expresiones regulares avanzadas para extraer cada dato con máxima precisión."""
    datos = {}

    # 1. Número de Formulario
    match_form = re.search(
        r"(?:Formulario|Número de formulario|No\.)\s*[:\.]?\s*(\S+)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["Numero_Formulario"] = match_form.group(1) if match_form else ""

    # 2. NIT del Importador
    match_nit = re.search(r"NIT\D*(\d{9,10})", texto_bloque, re.IGNORECASE)
    datos["NIT_Importador"] = match_nit.group(1) if match_nit else ""

    # 3. Factura
    match_factura = re.search(
        r"factura\D*([A-Za-z0-9\-]+)", texto_bloque, re.IGNORECASE
    )
    datos["Factura"] = match_factura.group(1) if match_factura else ""

    # 4. Manifiesto de Carga
    match_manifiesto = re.search(
        r"Manifiesto\s+de\s+carga\D*([A-Za-z0-9\-]+)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["Manifiesto_Carga"] = (
        match_manifiesto.group(1) if match_manifiesto else ""
    )

    # 5. Documento de Transporte
    match_transporte = re.search(
        r"Documento\s+de\s+transporte\D*([A-Za-z0-9\-]+)",
        texto_bloque,
        re.IGNORECASE,
    )
    datos["Documento_Transporte"] = (
        match_transporte.group(1) if match_transporte else ""
    )

    # 6. Subpartida Arancelaria
    match_subpartida = re.search(
        r"Subpartida\s+arancelaria\D*(\d{10})", texto_bloque, re.IGNORECASE
    )
    datos["Subpartida"] = match_subpartida.group(1) if match_subpartida else ""

    # 7. Levante No.
    match_levante = re.search(
        r"Levante\s+No\.?\s*([0-9]{8,15})", texto_bloque, re.IGNORECASE
    )
    datos["Levante_No"] = match_levante.group(1) if match_levante else ""

    return datos


# --- Interfaz Principal en Streamlit ---
uploaded_files = st.file_uploader(
    "Sube tus declaraciones de importación (Sueltas o Fusionadas en PDF)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    resultados_totales = []

    with st.spinner("Procesando archivos y extrayendo información..."):
        for uploaded_file in uploaded_files:
            paginas = extraer_texto_por_paginas(uploaded_file)
            bloques_declaraciones = segmentar_declaraciones(paginas)

            for idx, bloque in enumerate(bloques_declaraciones):
                info_extraida = extraer_datos_formulario(bloque)

                # Validar que el bloque contenga datos reales antes de agregarlo
                if any(info_extraida.values()):
                    info_extraida["Archivo_Origen"] = uploaded_file.name
                    info_extraida["Bloque_ID"] = idx + 1
                    resultados_totales.append(info_extraida)

    if resultados_totales:
        df_resultados = pd.DataFrame(resultados_totales)
        st.success(
            f"¡Éxito! Se procesaron correctamente {len(df_resultados)}"
            " declaraciones."
        )

        st.subheader("📊 Tabla de Datos Extraídos")
        st.dataframe(df_resultados, use_container_width=True)

        # Botón de descarga en CSV
        csv = df_resultados.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar Resultados en CSV",
            data=csv,
            file_name="declaraciones_extraidas.csv",
            mime="text/csv",
        )
    else:
        st.warning(
            "No se encontraron datos válidos en los archivos proporcionados."
        )
