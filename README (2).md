# Procesador Masivo de Declaraciones de Importación (DIM - Formulario 500 DIAN)

Aplicación web interna para extraer automáticamente 20 campos estandarizados
desde Declaraciones de Importación (formulario 500 de la DIAN) en PDF
(individuales o dentro de un ZIP) y exportar el resultado a Excel/CSV.

**Detalle importante:** un mismo PDF puede traer **varias declaraciones
seguidas** (por ejemplo, un archivo combinado con la DIM + su Acta de
Inspección, repetido 3 veces). La app detecta automáticamente cada
declaración dentro del PDF (usando el título "Declaración de Importación"
como marca de inicio) y genera **una fila por cada una** — no por archivo.

## Los 20 campos extraídos

| # | Campo | Casilla DIAN |
|---|-------|---------------|
| 1 | Número de formulario | 4 |
| 2 | NIT Importador | 5 |
| 3 | Razón Social Importador | 11 |
| 4 | NIT Declarante | 24 |
| 5 | Factura | 51 |
| 6 | Cod. País Procedencia | 53 |
| 7 | Cod. Modo Transporte | 54 |
| 8 | Código de Bandera | 55 |
| 9 | Tasa de Cambio | 58 |
| 10 | Subpartida Arancelaria | 59 |
| 11 | Cod. País Origen | 66 |
| 12 | Cod. País Compra | 70 |
| 13 | Peso Bruto (Kgs) | 71 |
| 14 | Peso Neto (Kgs) | 72 |
| 15 | Código de Embalaje | 73 |
| 16 | No. Bultos | 74 |
| 17 | Valor FOB (USD) | 78 |
| 18 | Sumatoria Fletes/Seguros/Otros (USD) | 82 |
| 19 | Levante No. | 134 |
| 20 | Fecha del Levante | 135 |

> ⚠️ **Nota sobre el campo 8 (Código de Bandera):** en tu mensaje original
> lo describiste como "N° 6", pero en el formulario 500 de la DIAN la
> casilla que se llama literalmente "Código de Bandera" es la **55**, y su
> valor es un código de país (viene de la misma tabla que "Códigos Países"
> que adjuntaste, no de la tabla "Modos de transporte"). Implementé la
> extracción sobre la casilla 55. Si en realidad te referías a otra casilla,
> avísame y la ajusto.

## Por qué los patrones RegEx usan longitudes fijas

En este formulario, el texto de una celda suele quedar **pegado sin
separador** al número de la celda siguiente cuando se extrae el PDF (por
ejemplo `"9014799636 . DV."` es en realidad el NIT `901479963` + el número
de campo `6` pegado). Para evitar capturar de más, cada patrón usa una
longitud o formato exacto conocido de antemano:

- Códigos de país / bandera → exactamente 3 dígitos (`\d{3}`)
- Modo de transporte → exactamente 1 dígito (`\d{1}`)
- NIT → exactamente 9 dígitos (`\d{9}`)
- Subpartida arancelaria → exactamente 10 dígitos (`\d{10}`)
- Levante No. → exactamente 15 dígitos (`\d{15}`)
- Montos en USD/pesos con decimales (FOB, fletes, pesos) → patrón
  `#.###.##` que exige un punto antes de cada grupo de 3 dígitos, por lo
  que se detiene naturalmente antes del número de campo siguiente.

Esto se validó extrayendo y comparando contra el texto real de tus 3 DIM de
prueba antes de entregar la app: los 20 campos salieron completos en las 3.

## Arquitectura

Igual que el procesador de actas: Streamlit puro, un solo archivo `app.py`,
para simplicidad de despliegue gratuito.

```
Usuario carga PDFs/ZIP
        │
        ▼
obtener_pdfs_desde_upload()  → descomprime ZIP en memoria
        │
        ▼
extraer_texto_pdf()          → PyMuPDF obtiene el texto plano
        │
        ▼
dividir_dims()                → separa el texto si hay varias DIM en el mismo PDF
        │
        ▼
extraer_campos_dim() (x N)   → RegEx sobre cada DIM → dict de 20 campos + QA
        │
        ▼
pandas.DataFrame             → tabla consolidada (1 fila por DIM), filtrable
        │
        ├──► generar_excel()  → openpyxl (estilos, bordes, autofiltro, freeze)
        └──► generar_csv()    → pandas.to_csv()
```

## 1. Instalación local

```bash
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Ejecutar en local

```bash
streamlit run app.py
```

## 3. Despliegue gratuito en Streamlit Community Cloud

1. Sube `app.py` y `requirements.txt` a un repositorio de GitHub (puede ser
   otro repo distinto al de las actas, o una carpeta separada del mismo).
2. Entra a [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Selecciona el repo, la rama y `app.py`.
4. Deploy.

> Si vas a usar el mismo repositorio que el procesador de actas, ponlos en
> carpetas separadas (por ejemplo `/actas/app.py` y `/dim/app.py`) y crea
> dos apps distintas en Streamlit Cloud, cada una apuntando a su propio
> archivo principal.

## Estructura del proyecto

```
dim_app/
├── app.py              # Aplicación completa (UI + extracción + exportación)
├── requirements.txt    # Dependencias
└── README.md           # Este archivo
```
