import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- CONSTANTES PARA REGEX ---
MONTO = r"[\d\.,]+"
ENTERO_MILES = r"[\d\.,]+"

# --- FUNCIONES AUXILIARES ---
def _buscar(patron, texto, grupo=1):
    """Busca un patrón en el texto y devuelve el grupo especificado."""
    match = re.search(patron, texto, re.IGNORECASE)
    return match.group(grupo) if match else None

def limpiar_monto(texto):
    """Limpia los montos quitando signos de dólar y espacios extra."""
    if not texto:
        return "0"
    return re.sub(r"[^\d\.,]", "", texto)

def extraer_texto_pdf(file_data) -> str:
    """Extrae todo el texto de un archivo PDF usando pdfplumber."""
    texto_completo = ""
    with pdfplumber.open(file_data) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                texto_completo += txt + "\n"
    return texto_completo

def dividir_dims(texto: str):
    """
    Divide el texto del PDF fusionado en fragmentos, cada uno correspondiente a una 
    Declaración de Importación (DIM) distinta, usando el campo 4 como separador principal.
    """
    partes = re.split(r"(?=4\.\s*N[uú]mero\s*de\s*formulario)", texto, flags=re.IGNORECASE)
    return [p for p in partes if re.search(r"4\.\s*N[uú]mero\s*de\s*formulario", p, re.IGNORECASE)]

# --- FUNCIÓN PRINCIPAL DE EXTRACCIÓN ---
def extraer_campos_dim(chunk_texto: str, nombre_archivo: str) -> dict:
    faltantes = []

    def campo(nombre, patron, grupo=1, default=""):
        # Buscamos SOLO en el chunk_texto de esta declaración específica
        valor = _buscar(patron, chunk_texto, grupo=grupo)
        if not valor:
            faltantes.append(nombre)
            return default
        return " ".join(valor.split())

    numero_formulario = campo("Número de formulario", r"4\s*\.\s*N[uú]mero de formulario\s*\n?\s*(\S+)")
    nit_importador = campo("NIT Importador", r"5\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9,10})")
    razon_social = campo("Razón Social Importador", r"11\s*\.\s*Apellidos y nombres o Raz[oó]n Social\s*([^\n]+)")
    factura = campo("Factura", r"51\s*\.\s*No\.\s*de\s*factura\s*\n\s*(\S+)")
    
    manifiesto_carga = campo("Manifiesto de carga", r"42\s*\.?\s*Manifiesto\s+de\s+carga\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")
    if not manifiesto_carga:
        manifiesto_carga = campo("Manifiesto de carga", r"42\s*\.?\s*Manifiesto\s+de\s+carga[^\n]*\n\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")

    documento_transporte = campo("Documento de transporte", r"44\s*\.?\s*Documento\s+de\s+transporte\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")
    if not documento_transporte:
        documento_transporte = campo("Documento de transporte", r"44\s*\.?\s*Documento\s+de\s+transporte[^\n]*\n\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")

    cod_pais_procedencia = campo("Cod. País Procedencia", r"53\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?procedencia\s*([A-Za-z0-9]{2,3})")
    cod_modo_transporte = campo("Cod. Modo Transporte", r"54\s*\.\s*Cod\.\s*Modo\s*Transporte\s*(\d)")
    codigo_bandera = campo("Código de Bandera", r"55\s*\.\s*C[oó]digo\s+(?:de\s+)?bandera\s*([A-Za-z0-9]{2,3})")
    tasa_cambio = campo("Tasa de Cambio", r"Tasa de cambio\s*\$?\s*cvs\.?\s*\n?\s*([\d\.,]+)")
    subpartida = campo("Subpartida Arancelaria", r"59\s*\.\s*Subpartida arancelaria\s*(\d{10})")
    cod_pais_origen = campo("Cod. País Origen", r"66\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?origen\s*([A-Za-z0-9]{2,3})")
    cod_pais_compra = campo("Cod. País Compra", r"70\s*\.\s*Cod\s*\.\s*pa[ií]s\s*\n?\s*compra\s*(\d{2,3})")
    codigo_embalaje = campo("Código de Embalaje", r"73\s*\.\s*C[oó]digo\s*\n?\s*embalaje\s*([A-Za-z0-9]{1,4})")
    
    peso_bruto = limpiar_monto(campo("Peso Bruto (Kgs)", r"71\s*\.\s*Peso bruto kgs\.\s*dcms\.\s*(" + MONTO + ")"))
    peso_neto = limpiar_monto(campo("Peso Neto (Kgs)", r"72\s*\.\s*Peso neto kgs\.\s*dcms\.\s*(" + MONTO + ")"))
    valor_fob = limpiar_monto(campo("Valor FOB (USD)", r"78\s*\.\s*Valor FOB USD\s*(" + MONTO + ")"))
    sumatoria_fletes = limpiar_monto(campo("Sumatoria Fletes/Seguros/Otros (USD)", r"82\s*\.\s*Sumatoria de fletes,?\s*seguros\s*\n?\s*y otros gastos USD\s*(" + MONTO + ")"))
    
    cod_unidad_comercial = campo("Cod. Unidad Comercial (76)", r"76\s*\.\s*Cod\.?\s*unidad\s*comercial\s*[\r\n\s]+([A-Za-z]{1,4})\b")
    if not cod_unidad_comercial:
        cod_unidad_comercial = campo("Cod. Unidad Comercial (76)", r"76\s*\.\s*Cod\.?\s*unidad\s*comercial\s+([A-Za-z]{1,4})\b")

    cantidad_str = campo("Cantidad (77)", r"77\s*\.\s*Cantidad\s*(?:dcms\.?)?\s*[\r\n\s]+(" + MONTO + ")")
    if not cantidad_str:
        cantidad_str = campo("Cantidad (77)", r"77\s*\.\s*Cantidad\s*(?:dcms\.?)?\s*(" + MONTO + ")")
    cantidad_comercial = limpiar_monto(cantidad_str)

    n_bultos_str = campo("No. Bultos", r"74\s*\.\s*No\.\s*bultos\s*(" + ENTERO_MILES + ")")
    try:
        no_bultos = int(re.sub(r'[^\d]', '', n_bultos_str)) if n_bultos_str else 0
    except ValueError:
        no_bultos = 0

    # Actas y Levantes
    acta_inspeccion = ""
    m_acta = re.search(r"ACTA\s+DE\s+INSPECCI[OÓ]N\s*(?:No\.?|Número)?\s*[:\.]?\s*([0-9]{8,15})", chunk_texto, re.IGNORECASE)
    if m_acta:
        acta_inspeccion = m_acta.group(1).strip()

    levante_no = ""
    m_lev_box = re.search(r"134\.?\s*Levante\s+No\.?\s*([0-9]{8,15})", chunk_texto, re.IGNORECASE)
    if m_lev_box:
        levante_no = m_lev_box.group(1).strip()
    
    if not levante_no:
        m_lev_gen = re.search(r"(?:Levante|Auto(?:rización)?)\s*(?:No\.?|Número)?\s*[:\.]?\s*([0-9]{8,15})", chunk_texto, re.IGNORECASE)
        if m_lev_gen:
            levante_no = m_lev_gen.group(1).strip()

    if not levante_no:
        faltantes.append("Levante No.")

    fecha_levante = campo("Fecha del Levante", r"135\.?\s*Fecha[^\d\n]*(\d{4}\s*[-/\.]\s*\d{2}\s*[-/\.]\s*\d{2})")
    if not fecha_levante:
        m_fec = re.search(r"\b(20\d{2}[-/\.](?:0[1-9]|1[0-2])[-/\.](?:0[1-9]|[12]\d|3[01]))\b", chunk_texto)
        if m_fec:
            fecha_levante = m_fec.group(1)
            if "Fecha del Levante" in faltantes:
                faltantes.remove("Fecha del Levante")

    if fecha_levante:
        fecha_levante = re.sub(r"\s+", "", fecha_levante)
        fecha_levante = re.sub(r"[/.]", "-", fecha_levante)

    return {
        "Número de formulario": numero_formulario,
        "NIT Importador": nit_importador,
        "Razón Social Importador": razon_social,
        "Factura": factura,
        "Manifiesto de carga": manifiesto_carga,
        "Documento de transporte": documento_transporte,
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
        "Cod. Unidad Comercial (76)": cod_unidad_comercial,
        "Cantidad (77)": cantidad_comercial,
        "No. Bultos": no_bultos,
        "Valor FOB (USD)": valor_fob,
        "Sumatoria Fletes/Seguros/Otros (USD)": sumatoria_fletes,
        "Acta de Inspección No.": acta_inspeccion,
        "Levante No.": levante_no,
        "Fecha del Levante": fecha_levante,
        "Archivo": nombre_archivo,
        "Campos_no_encontrados": ", ".join(faltantes) if faltantes else "",
    }

# --- FUNCIÓN DE PROCESAMIENTO MÚLTIPLE ---
def procesar_archivos(archivos_subidos):
    filas = []
    for archivo in archivos_subidos:
        nombre_pdf = archivo.name
        try:
            # 1. Extraemos todo el texto del archivo
            texto_completo = extraer_texto_pdf(archivo)
            
            # 2. Dividimos el texto en bloques por cada DIM
            dim_chunks = dividir_dims(texto_completo)
            
            # Si por alguna razón no detectó separadores, lo tratamos como una sola DIM
            if not dim_chunks:
                dim_chunks = [texto_completo]
                
            # 3. Extraemos la información de cada fragmento
            for chunk in dim_chunks:
                fila = extraer_campos_dim(chunk, nombre_pdf)
                filas.append(fila)
        except Exception as e:
            st.error(f"Error procesando el archivo {nombre_pdf}: {e}")
            
    return pd.DataFrame(filas)

# --- INTERFAZ DE STREAMLIT ---
def main():
    st.set_page_config(page_title="Extractor DIM", layout="wide")
    st.title("📄 Extractor de Datos - Declaraciones de Importación (DIM)")
    st.write("Sube tus archivos PDF (DIM individuales o fusionados) para extraer los datos a Excel.")
    
    archivos_subidos = st.file_uploader("Sube uno o varios PDFs", type=["pdf"], accept_multiple_files=True)
    
    if archivos_subidos:
        if st.button("Procesar PDFs", type="primary"):
            with st.spinner("Leyendo y procesando archivos..."):
                df_resultado = procesar_archivos(archivos_subidos)
                
                if not df_resultado.empty:
                    st.success("¡Procesamiento exitoso!")
                    st.dataframe(df_resultado)
                    
                    # Generar Excel en memoria para descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_resultado.to_excel(writer, index=False, sheet_name='Datos Extraídos')
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Descargar datos en Excel",
                        data=output,
                        file_name="Datos_DIM_Extraidos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se pudo extraer información de los archivos subidos.")

if __name__ == "__main__":
    main()
