# -*- coding: utf-8 -*-
"""
Procesador Masivo de Declaraciones de Importación (DIM - Formulario 500 DIAN)
================================================================================
App Streamlit para cargar PDFs sueltos o ZIPs con múltiples PDFs de
Declaraciones de Importación, extraer 20 campos estandarizados mediante RegEx
y exportar el resultado a Excel (.xlsx) formateado y CSV.

IMPORTANTE: un mismo PDF puede contener MÁS DE UNA declaración (varias DIM
consecutivas en un solo archivo, cada una con su propia Acta de Inspección
intercalada). La app detecta automáticamente cada declaración dentro del PDF
y genera una fila por cada una.

Ejecutar con:
    streamlit run app.py
"""

import io
import re
import zipfile
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuración general
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="Procesador de DIM",
    page_icon="🧾",
    layout="wide",
)

COLUMNAS = [
    "Número de formulario",
    "NIT Importador",
    "Razón Social Importador",
    "NIT Declarante",
    "Factura",
    "Cod. País Procedencia",
    "Cod. Modo Transporte",
    "Código de Bandera",
    "Tasa de Cambio",
    "Subpartida Arancelaria",
    "Cod. País Origen",
    "Cod. País Compra",
    "Peso Bruto (Kgs)",
    "Peso Neto (Kgs)",
    "Código de Embalaje",
    "No. Bultos",
    "Valor FOB (USD)",
    "Sumatoria Fletes/Seguros/Otros (USD)",
    "Levante No.",
    "Fecha del Levante",
    "Archivo",                # columna auxiliar de trazabilidad
    "Campos_no_encontrados",  # columna auxiliar de QA
]

# Patrones numéricos reutilizables (formato colombiano: "." como separador
# de miles). Se usan con longitud/estructura acotada para no "comerse" el
# número del siguiente campo, que en estos PDF suele quedar pegado sin
# separador al valor anterior.
MONTO = r"\d{1,3}(?:\.\d{3})*\.\d{2}"     # ej: 1.638.00 / 262.50
ENTERO_MILES = r"\d{1,3}(?:\.\d{3})*"      # ej: 1.061

# --------------------------------------------------------------------------
# Utilidades de extracción
# --------------------------------------------------------------------------


def _buscar(patron, texto, flags=re.IGNORECASE, grupo=1):
    """Devuelve el grupo capturado o None si no hay match."""
    m = re.search(patron, texto, flags)
    if m:
        try:
            return m.group(grupo).strip()
        except IndexError:
            return None
    return None


def dividir_dims(texto: str):
    """
    Un mismo PDF puede traer varias Declaraciones de Importación seguidas
    (cada una con su Acta de Inspección intercalada). Esta función separa
    el texto completo en un bloque por cada declaración encontrada, usando
    el título "Declaración de Importación" como marca de inicio.
    """
    partes = re.split(r"(?=Declaraci[oó]n de Importaci[oó]n)", texto, flags=re.IGNORECASE)
    return [p for p in partes if re.search(r"N[uú]mero de formulario", p, re.IGNORECASE)]


def extraer_campos_dim(texto: str, nombre_archivo: str) -> dict:
    """
    Aplica los patrones RegEx sobre el texto de UNA declaración de importación
    y retorna un diccionario con los 20 campos estandarizados + metadatos de QA.
    """
    faltantes = []

    def campo(nombre, patron, grupo=1, default=""):
        valor = _buscar(patron, texto, grupo=grupo)
        if not valor:
            faltantes.append(nombre)
            return default
        return " ".join(valor.split())

    # 1. Número de formulario (campo 4)
    numero_formulario = campo(
        "Número de formulario",
        r"4\s*\.\s*N[uú]mero de formulario\s*\n?\s*(\S+)",
    )

    # 2. NIT Importador (campo 5) -- el NIT colombiano tiene 9 dígitos; se fija
    #    la longitud para no arrastrar el número del campo siguiente (6. DV)
    #    que suele quedar pegado sin separador.
    nit_importador = campo(
        "NIT Importador",
        r"5\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9})",
    )

    # 3. Razón Social Importador (campo 11)
    razon_social = campo(
        "Razón Social Importador",
        r"11\s*\.\s*Apellidos y nombres o Raz[oó]n Social\s*([^\n]+)",
    )

    # 4. NIT Declarante (campo 24)
    nit_declarante = campo(
        "NIT Declarante",
        r"24\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9})",
    )

    # 5. Factura (campo 51)
    factura = campo(
        "Factura",
        r"51\s*\.\s*No\.\s*de\s*factura\s*\n\s*(\S+)",
    )

    # 6. Cod. País Procedencia (campo 53) -- código de 3 dígitos (ver Códigos Países)
    cod_pais_procedencia = campo(
        "Cod. País Procedencia",
        r"53\s*\.\s*Cod\.\s*pa[ií]s procedencia\s*(\d{3})",
    )

    # 7. Cod. Modo Transporte (campo 54) -- código de 1 dígito (ver Modos de transporte)
    cod_modo_transporte = campo(
        "Cod. Modo Transporte",
        r"54\s*\.\s*Cod\.\s*Modo\s*Transporte\s*(\d)",
    )

    # 8. Código de Bandera (campo 55) -- código de país (3 dígitos)
    codigo_bandera = campo(
        "Código de Bandera",
        r"55\s*\.\s*C[oó]digo de Bandera\s*(\d{3})",
    )

    # 9. Tasa de Cambio (campo 58)
    tasa_cambio = campo(
        "Tasa de Cambio",
        r"Tasa de cambio\s*\$?\s*cvs\.?\s*\n?\s*([\d.,]+)",
    )

    # 10. Subpartida Arancelaria (campo 59) -- arancel colombiano: 10 dígitos fijos
    subpartida = campo(
        "Subpartida Arancelaria",
        r"59\s*\.\s*Subpartida arancelaria\s*(\d{10})",
    )

    # 11. Cod. País Origen (campo 66)
    cod_pais_origen = campo(
        "Cod. País Origen",
        r"66\s*\.\s*Cod\.\s*pa[ií]s de origen\s*(\d{3})",
    )

    # 12. Cod. País Compra (campo 70)
    cod_pais_compra = campo(
        "Cod. País Compra",
        r"70\s*\.\s*Cod\.\s*pa[ií]s\s*\n?\s*compra\s*(\d{3})",
    )

    # 13. Peso Bruto Kgs (campo 71)
    peso_bruto = campo(
        "Peso Bruto (Kgs)",
        r"71\s*\.\s*Peso bruto kgs\.\s*dcms\.\s*(" + MONTO + ")",
    )

    # 14. Peso Neto Kgs (campo 72)
    peso_neto = campo(
        "Peso Neto (Kgs)",
        r"72\s*\.\s*Peso neto kgs\.\s*dcms\.\s*(" + MONTO + ")",
    )

    # 15. Código de Embalaje (campo 73)
    codigo_embalaje = campo(
        "Código de Embalaje",
        r"73\s*\.\s*C[oó]digo\s*\n?\s*embalaje\s*([A-Z]{1,4})",
    )

    # 16. No. Bultos (campo 74)
    no_bultos = campo(
        "No. Bultos",
        r"74\s*\.\s*No\.\s*bultos\s*(" + ENTERO_MILES + ")",
    )

    # 17. Valor FOB USD (campo 78)
    valor_fob = campo(
        "Valor FOB (USD)",
        r"78\s*\.\s*Valor FOB USD\s*(" + MONTO + ")",
    )

    # 18. Sumatoria de fletes, seguros y otros gastos USD (campo 82)
    sumatoria_fletes = campo(
        "Sumatoria Fletes/Seguros/Otros (USD)",
        r"82\s*\.\s*Sumatoria de fletes,?\s*seguros\s*\n?\s*y otros gastos USD\s*(" + MONTO + ")",
    )

    # 19. Levante No. (campo 134) -- número de 15 dígitos fijos
    levante_no = campo(
        "Levante No.",
        r"134\s*\.\s*Levante No\.\s*(\d{15})",
    )

    # 20. Fecha del Levante (campo 135)
    fecha_levante = campo(
        "Fecha del Levante",
        r"135\s*\.\s*Fecha\s*\n?\s*(\d{4}\s*-\s*\d{2}\s*-\s*\d{2})",
    )
    if fecha_levante:
        fecha_levante = re.sub(r"\s*-\s*", "-", fecha_levante)

    return {
        "Número de formulario": numero_formulario,
        "NIT Importador": nit_importador,
        "Razón Social Importador": razon_social,
        "NIT Declarante": nit_declarante,
        "Factura": factura,
        "Cod. País Procedencia": cod_pais_procedencia,
        "Cod. Modo Transporte": cod_modo_transporte,
        "Código de Bandera": codigo_bandera,
        "Tasa de Cambio": tasa_cambio,
        "Subpartida Arancelaria": subpartida,
        "Cod. País Origen": cod_pais_origen,
        "Cod. País Compra": cod_pais_compra,
        "Peso Bruto (Kgs)": peso_bruto,
        "Peso Neto (Kgs)": peso_neto,
        "Código de Embalaje": codigo_embalaje,
        "No. Bultos": no_bultos,
        "Valor FOB (USD)": valor_fob,
        "Sumatoria Fletes/Seguros/Otros (USD)": sumatoria_fletes,
        "Levante No.": levante_no,
        "Fecha del Levante": fecha_levante,
        "Archivo": nombre_archivo,
        "Campos_no_encontrados": ", ".join(faltantes) if faltantes else "",
    }


def extraer_texto_pdf(data: bytes) -> str:
    """Extrae el texto plano de un PDF usando PyMuPDF."""
    with fitz.open(stream=data, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)


def obtener_pdfs_desde_upload(uploaded_file):
    """
    Recibe un archivo subido (PDF o ZIP) y retorna una lista de tuplas
    (nombre_archivo, bytes_pdf).
    """
    nombre = uploaded_file.name
    contenido = uploaded_file.read()

    if nombre.lower().endswith(".zip"):
        pdfs = []
        with zipfile.ZipFile(io.BytesIO(contenido)) as zf:
            for info in zf.infolist():
                if info.filename.lower().endswith(".pdf") and not info.is_dir():
                    pdfs.append((info.filename.split("/")[-1], zf.read(info.filename)))
        return pdfs
    elif nombre.lower().endswith(".pdf"):
        return [(nombre, contenido)]
    else:
        return []


def procesar_archivos(uploaded_files, progress_callback=None) -> pd.DataFrame:
    """
    Procesa la lista de archivos subidos y retorna el DataFrame consolidado.
    Cada PDF puede generar MÁS DE UNA fila si contiene varias DIM.
    """
    filas = []
    tareas = []

    for uf in uploaded_files:
        tareas.extend(obtener_pdfs_desde_upload(uf))

    total = max(len(tareas), 1)
    for i, (nombre_pdf, data) in enumerate(tareas, start=1):
        try:
            texto = extraer_texto_pdf(data)
            dim_chunks = dividir_dims(texto)
            if not dim_chunks:
                # No se detectó el marcador de "Declaración de Importación";
                # se intenta igual sobre el texto completo por si acaso.
                dim_chunks = [texto]
            for chunk in dim_chunks:
                fila = extraer_campos_dim(chunk, nombre_pdf)
                filas.append(fila)
        except Exception as exc:  # noqa: BLE001
            fila = {c: "" for c in COLUMNAS}
            fila["Archivo"] = nombre_pdf
            fila["Campos_no_encontrados"] = f"ERROR AL PROCESAR: {exc}"
            filas.append(fila)
        if progress_callback:
            progress_callback(i / total, nombre_pdf)

    if not filas:
        return pd.DataFrame(columns=COLUMNAS)

    return pd.DataFrame(filas, columns=COLUMNAS)


# --------------------------------------------------------------------------
# Exportación a Excel formateado
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def generar_excel(df: pd.DataFrame) -> bytes:
    """Genera un archivo Excel (.xlsx) con formato profesional a partir del DataFrame."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DIM")
        ws = writer.sheets["DIM"]

        n_filas = df.shape[0]
        n_cols = df.shape[1]

        for col_idx in range(1, n_cols + 1):
            celda = ws.cell(row=1, column=col_idx)
            celda.fill = HEADER_FILL
            celda.font = HEADER_FONT
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = THIN_BORDER

        for row_idx in range(2, n_filas + 2):
            for col_idx in range(1, n_cols + 1):
                ws.cell(row=row_idx, column=col_idx).border = THIN_BORDER
                ws.cell(row=row_idx, column=col_idx).alignment = Alignment(vertical="top", wrap_text=True)

        for col_idx, columna in enumerate(df.columns, start=1):
            longitudes = [len(str(columna))] + [
                len(str(v)) for v in df[columna].astype(str).tolist()
            ]
            ancho = min(max(longitudes) + 3, 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = ancho

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

    return buffer.getvalue()


def generar_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")


# --------------------------------------------------------------------------
# Interfaz Streamlit
# --------------------------------------------------------------------------

st.title("🧾 Procesador Masivo de Declaraciones de Importación (DIM)")
st.caption("Formulario 500 DIAN · Extracción automática de 20 campos estandarizados desde PDF")
st.info(
    "ℹ️ Si un mismo PDF contiene varias declaraciones seguidas (por ejemplo, "
    "un archivo combinado con 3 DIM), la app las detecta automáticamente y "
    "genera una fila por cada declaración.",
    icon="ℹ️",
)

if "df_resultado_dim" not in st.session_state:
    st.session_state.df_resultado_dim = pd.DataFrame(columns=COLUMNAS)

with st.container(border=True):
    st.subheader("1. Cargar documentos")
    uploaded_files = st.file_uploader(
        "Arrastra aquí archivos PDF individuales o un ZIP con varios PDFs",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        help="Puedes combinar PDFs sueltos y archivos ZIP en la misma carga.",
        key="dim_uploader",
    )

    procesar = st.button("🚀 Procesar Documentos", type="primary", use_container_width=True)

if procesar:
    if not uploaded_files:
        st.warning("Por favor carga al menos un archivo PDF o ZIP antes de procesar.")
    else:
        progreso = st.progress(0.0, text="Iniciando procesamiento...")

        def _cb(pct, nombre):
            progreso.progress(pct, text=f"Procesando: {nombre}")

        with st.spinner("Extrayendo información de las declaraciones..."):
            df = procesar_archivos(uploaded_files, progress_callback=_cb)

        progreso.empty()
        st.session_state.df_resultado_dim = df
        st.success(f"✅ Procesamiento completado. {len(df)} declaración(es) extraída(s).")

# --------------------------------------------------------------------------
# Resultados
# --------------------------------------------------------------------------

df = st.session_state.df_resultado_dim

if not df.empty:
    st.subheader("2. Resultados")

    busqueda = st.text_input("🔍 Buscar en todos los campos", "")

    df_vista = df.copy()
    if busqueda:
        mask = df_vista.apply(
            lambda fila: fila.astype(str).str.contains(busqueda, case=False, na=False).any(),
            axis=1,
        )
        df_vista = df_vista[mask]

    st.dataframe(df_vista, use_container_width=True, height=420)

    faltantes_totales = df[df["Campos_no_encontrados"] != ""]
    if not faltantes_totales.empty:
        with st.expander(f"⚠️ {len(faltantes_totales)} declaración(es) con campos no detectados (revisión manual)"):
            st.dataframe(
                faltantes_totales[["Número de formulario", "Archivo", "Campos_no_encontrados"]],
                use_container_width=True,
            )

    with st.expander("👁️ Vista previa de una declaración"):
        opciones = (df["Número de formulario"] + " — " + df["Archivo"]).tolist()
        seleccion = st.selectbox("Selecciona una declaración", opciones)
        idx = opciones.index(seleccion)
        st.json(df.iloc[idx].to_dict())

    st.subheader("3. Descargar resultados")
    df_export = df.drop(columns=["Campos_no_encontrados"])

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Descargar Excel (.xlsx)",
            data=generar_excel(df_export),
            file_name=f"dim_procesadas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            "⬇️ Descargar CSV (.csv)",
            data=generar_csv(df_export),
            file_name=f"dim_procesadas_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
else:
    st.info("Carga tus archivos PDF/ZIP y presiona **Procesar Documentos** para ver los resultados aquí.")
