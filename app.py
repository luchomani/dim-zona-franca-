# -*- coding: utf-8 -*-
"""
Procesador Masivo de Declaraciones de Importación (DIM - Formulario 500 DIAN)
================================================================================
App Streamlit para cargar PDFs sueltos o ZIPs con múltiples PDFs de
Declaraciones de Importación, extraer 20 campos estandarizados mediante RegEx
y exportar el resultado a Excel (.xlsx) formateado y CSV.
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
    "Archivo",
    "Campos_no_encontrados",
]

MONTO = r"\d{1,3}(?:\.\d{3})*[\.,]\d{2}|\d+(?:[\.,]\d+)?"
ENTERO_MILES = r"\d{1,3}(?:\.\d{3})*|\d+"

# --------------------------------------------------------------------------
# Funciones Matemáticas y de Limpieza
# --------------------------------------------------------------------------

def limpiar_monto(val_str):
    """Convierte los montos de la DIAN a float nativo de forma limpia y directa."""
    if not val_str:
        return 0.0
    val_str = str(val_str).strip()
    
    # Limpieza de caracteres extraños o letras de relleno típicas de la DIAN
    val_str = re.sub(r'[^\d.,]', '', val_str)
    if not val_str:
        return 0.0

    # Si usa coma como decimal y punto como miles
    if ',' in val_str and '.' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    elif val_str.count('.') > 1:
        partes = val_str.rsplit('.', 1)
        val_str = partes[0].replace('.', '') + '.' + partes[1]

    try:
        return float(val_str)
    except ValueError:
        return 0.0

# --------------------------------------------------------------------------
# Utilidades de extracción
# --------------------------------------------------------------------------

def _buscar(patron, texto, flags=re.IGNORECASE, grupo=1):
    m = re.search(patron, texto, flags)
    if m:
        try:
            return m.group(grupo).strip()
        except IndexError:
            return None
    return None

def dividir_dims(texto: str):
    partes = re.split(r"(?=Declaraci[oó]n de Importaci[oó]n)", texto, flags=re.IGNORECASE)
    return [p for p in partes if re.search(r"N[uú]mero de formulario", p, re.IGNORECASE)]

def extraer_campos_dim(texto: str, nombre_archivo: str) -> dict:
    faltantes = []

    def campo(nombre, patron, grupo=1, default=""):
        valor = _buscar(patron, texto, grupo=grupo)
        if not valor:
            faltantes.append(nombre)
            return default
        return " ".join(valor.split())

    numero_formulario = campo("Número de formulario", r"4\s*\.\s*N[uú]mero de formulario\s*\n?\s*(\S+)")
    nit_importador = campo("NIT Importador", r"5\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9,10})")
    razon_social = campo("Razón Social Importador", r"11\s*\.\s*Apellidos y nombres o Raz[oó]n Social\s*([^\n]+)")
    nit_declarante = campo("NIT Declarante", r"24\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9,10})")
    factura = campo("Factura", r"51\s*\.\s*No\.\s*de\s*factura\s*\n\s*(\S+)")
    
    cod_pais_procedencia = campo("Cod. País Procedencia", r"53\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?procedencia\s*([A-Za-z0-9]{2,3})")
    cod_modo_transporte = campo("Cod. Modo Transporte", r"54\s*\.\s*Cod\.\s*Modo\s*Transporte\s*(\d)")
    codigo_bandera = campo("Código de Bandera", r"55\s*\.\s*C[oó]digo\s+(?:de\s+)?bandera\s*([A-Za-z0-9]{2,3})")
    tasa_cambio = campo("Tasa de Cambio", r"Tasa de cambio\s*\$?\s*cvs\.?\s*\n?\s*([\d.,]+)")
    subpartida = campo("Subpartida Arancelaria", r"59\s*\.\s*Subpartida arancelaria\s*(\d{10})")
    cod_pais_origen = campo("Cod. País Origen", r"66\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?origen\s*([A-Za-z0-9]{2,3})")
    cod_pais_compra = campo("Cod. País Compra", r"70\s*\.\s*Cod\.\s*pa[ií]s\s*\n?\s*compra\s*(\d{2,3})")
    codigo_embalaje = campo("Código de Embalaje", r"73\s*\.\s*C[oó]digo\s*\n?\s*embalaje\s*([A-Za-z0-9]{1,4})")
    
    # -----------------------------------------------------------
    # Extracción y conversión de valores numéricos limpios
    # -----------------------------------------------------------
    peso_bruto = limpiar_monto(campo("Peso Bruto (Kgs)", r"71\s*\.\s*Peso bruto kgs\.\s*dcms\.\s*(" + MONTO + ")"))
    peso_neto = limpiar_monto(campo("Peso Neto (Kgs)", r"72\s*\.\s*Peso neto kgs\.\s*dcms\.\s*(" + MONTO + ")"))
    valor_fob = limpiar_monto(campo("Valor FOB (USD)", r"78\s*\.\s*Valor FOB USD\s*(" + MONTO + ")"))
    sumatoria_fletes = limpiar_monto(campo("Sumatoria Fletes/Seguros/Otros (USD)", r"82\s*\.\s*Sumatoria de fletes,?\s*seguros\s*\n?\s*y otros gastos USD\s*(" + MONTO + ")"))
    
    n_bultos_str = campo("No. Bultos", r"74\s*\.\s*No\.\s*bultos\s*(" + ENTERO_MILES + ")")
    try:
        no_bultos = int(re.sub(r'[^\d]', '', n_bultos_str)) if n_bultos_str else 0
    except ValueError:
        no_bultos = 0
    # -----------------------------------------------------------

    # Patrones mejorados y flexibles para Levante y Fecha de Levante
    levante_no = campo("Levante No.", r"134\.?\s*Levante\s+No\.?\s*([A-Za-z0-9\-_]{5,30})")
    if not levante_no:
        levante_no = campo("Levante No.", r"Levante\s+No\.?\s*([A-Za-z0-9\-_]{5,30})")

    fecha_levante = campo("Fecha del Levante", r"135\.?\s*Fecha[^\d\n]*(\d{4}\s*[-/\.]\s*\d{2}\s*[-/\.]\s*\d{2})")
    if not fecha_levante:
        fecha_levante = campo("Fecha del Levante", r"(?:\b20\d{2}[-/\.]\d{2}[-/\.]\d{2}\b)")
        
    if fecha_levante:
        fecha_levante = re.sub(r"\s+", "", fecha_levante)
        fecha_levante = re.sub(r"[/.]", "-", fecha_levante)

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
    with fitz.open(stream=data, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

def obtener_pdfs_desde_upload(uploaded_file):
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
                dim_chunks = [texto]
            for chunk in dim_chunks:
                fila = extraer_campos_dim(chunk, nombre_pdf)
                filas.append(fila)
        except Exception as exc: 
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
# Exportación a Excel formateado (CON TOTALES Y FORMATO NÚMERICO)
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TOTAL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

def generar_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    df_export = df.copy()

    if not df_export.empty:
        total_row = {c: "" for c in df_export.columns}
        total_row["Número de formulario"] = "TOTALES CONSOLIDADOS"
        total_row["Valor FOB (USD)"] = df_export["Valor FOB (USD)"].sum()
        total_row["Sumatoria Fletes/Seguros/Otros (USD)"] = df_export["Sumatoria Fletes/Seguros/Otros (USD)"].sum()
        total_row["Peso Bruto (Kgs)"] = df_export["Peso Bruto (Kgs)"].sum()
        total_row["Peso Neto (Kgs)"] = df_export["Peso Neto (Kgs)"].sum()
        total_row["No. Bultos"] = df_export["No. Bultos"].sum()
        
        df_export = pd.concat([df_export, pd.DataFrame([total_row])], ignore_index=True)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="DIM")
        ws = writer.sheets["DIM"]

        n_filas = df_export.shape[0]
        n_cols = df_export.shape[1]

        for col_idx in range(1, n_cols + 1):
            celda = ws.cell(row=1, column=col_idx)
            celda.fill = HEADER_FILL
            celda.font = HEADER_FONT
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = THIN_BORDER

        for row_idx in range(2, n_filas + 2):
            for col_idx in range(1, n_cols + 1):
                celda = ws.cell(row=row_idx, column=col_idx)
                celda.border = THIN_BORDER
                celda.alignment = Alignment(vertical="top", wrap_text=True)
                
                if isinstance(celda.value, (int, float)):
                    celda.number_format = '#,##0.00'
                    
                if row_idx == n_filas + 1:
                    celda.fill = TOTAL_FILL
                    celda.font = Font(bold=True)

        for col_idx, columna in enumerate(df_export.columns, start=1):
            longitudes = [len(str(columna))] + [len(str(v)) for v in df_export[columna].astype(str).tolist()]
            ancho = min(max(longitudes) + 3, 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = ancho

        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_filas}"
        ws.freeze_panes = "A2"

    return buffer.getvalue()

def generar_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")

# --------------------------------------------------------------------------
# Interfaz Streamlit
# --------------------------------------------------------------------------

st.title("🧾 Procesador Masivo de Declaraciones de Importación (DIM)")
st.caption("Formulario 500 DIAN · Extracción automática de 20 campos estandarizados desde PDF")

if "df_resultado_dim" not in st.session_state:
    st.session_state.df_resultado_dim = pd.DataFrame(columns=COLUMNAS)

with st.container(border=True):
    st.subheader("1. Cargar documentos")
    uploaded_files = st.file_uploader(
        "Arrastra aquí archivos PDF individuales o un ZIP con varios PDFs",
        type=["pdf", "zip"],
        accept_multiple_files=True,
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
# Resultados y Totales
# --------------------------------------------------------------------------

df = st.session_state.df_resultado_dim

if not df.empty:
    st.subheader("2. Resultados y Consolidado")

    st.markdown("##### 📊 Totales de la extracción actual")
    cols_totales = st.columns(4)
    
    total_fob = df["Valor FOB (USD)"].sum()
    total_fletes = df["Sumatoria Fletes/Seguros/Otros (USD)"].sum()
    total_peso_bruto = df["Peso Bruto (Kgs)"].sum()
    total_peso_neto = df["Peso Neto (Kgs)"].sum()

    cols_totales[0].metric("Total Valor FOB (USD)", f"${total_fob:,.2f}")
    cols_totales[1].metric("Total Fletes/Seguros (USD)", f"${total_fletes:,.2f}")
    cols_totales[2].metric("Total Peso Bruto (Kgs)", f"{total_peso_bruto:,.2f}")
    cols_totales[3].metric("Total Peso Neto (Kgs)", f"{total_peso_neto:,.2f}")
    st.divider()

    busqueda = st.text_input("🔍 Buscar en todos los campos", "")

    df_vista = df.copy()
    if busqueda:
        mask = df_vista.apply(lambda fila: fila.astype(str).str.contains(busqueda, case=False, na=False).any(), axis=1)
        df_vista = df_vista[mask]

    st.dataframe(
        df_vista, 
        use_container_width=True, 
        height=420,
        column_config={
            "Valor FOB (USD)": st.column_config.NumberColumn(format="%.2f"),
            "Sumatoria Fletes/Seguros/Otros (USD)": st.column_config.NumberColumn(format="%.2f"),
            "Peso Bruto (Kgs)": st.column_config.NumberColumn(format="%.2f"),
            "Peso Neto (Kgs)": st.column_config.NumberColumn(format="%.2f"),
            "No. Bultos": st.column_config.NumberColumn(format="%d"),
        }
    )

    faltantes_totales = df[df["Campos_no_encontrados"] != ""]
    if not faltantes_totales.empty:
        with st.expander(f"⚠️ {len(faltantes_totales)} declaración(es) con campos no detectados (revisión manual)"):
            st.dataframe(faltantes_totales[["Número de formulario", "Archivo", "Campos_no_encontrados"]], use_container_width=True)

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
