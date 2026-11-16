# App Planilla JS + AG Grid + Power Query + Agente IA

Entorno integral todo en uno desarrollado en **Python y Streamlit** que combina la flexibilidad de una hoja de cálculo con motor de fórmulas en JavaScript (**Jspreadsheet CE**), la potencia analítica de **AG Grid**, transformaciones **ETL (estilo Power Query)**, **Tablas Dinámicas** y un **Agente de Inteligencia Artificial (LangChain + OpenAI)** capaz de analizar tus datos en lenguaje natural.

---

## Características Principales

* ** Planilla JS con Motor de Fórmulas Real (Jspreadsheet CE):**
  * Carga instantánea (< 0.5 s) con soporte para fórmulas nativas de Excel (`=SUM()`, `=AVERAGE()`, `=IF()`, referencias de celdas, inserción/eliminación de filas y columnas).
  * **Autoguardado silencioso y Sincronización en 1 clic:** Captura los valores ya calculados de tus fórmulas y los transfiere directamente a Python (`pandas.DataFrame`).
* ** Grilla Analítica Bidireccional (AG Grid):**
  * Edición en tiempo real sincronizada con el estado global de la aplicación.
  * Filtros avanzados estilo Excel, ordenamiento multicriterio, paginación y panel lateral de agrupación.
* ** Power Query (Constructor Visual ETL):**
  * Limpieza y transformación de datos sin código: filtrado por texto/nulos, eliminación de duplicados, conversión a mayúsculas, relleno de valores vacíos y agrupación por métricas (*Group By*) con vista previa en vivo.
* ** Tablas Dinámicas (Pivot Tables):**
  * Generación dinámica de matrices cruzadas configurando filas, columnas, métricas y funciones de agregación (`sum`, `mean`, `count`, `min`, `max`).
* ** Agente IA Analítico (LangChain + OpenAI):**
  * Consultas en lenguaje natural sobre los datos activos utilizando `gpt-4o-mini` o `gpt-4o` con ejecución de código Python en tiempo real.
* ** Importación y Exportación Multiformato:**
  * **Importación:** `.xlsx`, `.xls` (con selector de hojas), `.csv` (detección automática o manual de separador), `.json`, `.tsv` y `.txt`.
  * **Exportación:** Descarga directa en `.xlsx`, `.csv` (delimitador `,` o `;`), `.json` y `.tsv`.

---

## Requisitos e Instalación

1. **Clonar el repositorio o descargar los archivos:**
   ```bash
   git clone [<url-de-tu-repositorio>](https://github.com/wgekko/app_hojas_calulo-power_query_IA.git)
   cd app_hojas_calulo-power_query_IA

   ```

python -m venv venv
# En Windows:
venv\Scripts\activate
# En Linux / macOS:
source venv/bin/activate

pip install streamlit pandas numpy openpyxl streamlit-aggrid streamlit-js-eval langchain-openai langchain-experimental

streamlit run app.py


Flujo de Trabajo Recomendado
Carga de datos: Inicia con la hoja en blanco, presiona Cargar Demo en la barra lateral o sube tu propio archivo (.xlsx, .csv, .json).
Edición y Fórmulas: En la pestaña 
1 Carga & Fórmulas (Planilla JS), ingresa datos o fórmulas tipo Excel. Luego haz clic en el botón 🔄 Sincronizar cambios hacia AG Grid, Power Query e IA.Transformación y Análisis:Refina o filtra registros en 
2 Sincronización & ETL (AG Grid) o aplica pasos de limpieza en 
3 Power Query.Resume indicadores clave en 
4 Tablas Dinámicas.Consultas con IA: Ingresa tu OpenAI API Key en la barra lateral, ve a 
5 Agente IA (LangChain) y realiza preguntas en español sobre tu tabla.Descarga: Ve a 
6 Exportar Datos para guardar el resultado final en el formato que necesites.
Arquitectura Técnica
Frontend / Grillas: Jspreadsheet CE v4 (embebido mediante st.iframe con caché en Base64) + streamlit-aggrid.
Puente JS $\leftrightarrow$ Python: localStorage + streamlit-js-eval para transferencia de estado en un solo clic.
Rendimiento: Uso de @st.cache_data, @st.fragment para aislar pestañas sin recargar el DOM de las grillas y lazy imports para las librerías de LangChain.
Motor de Datos e IA: Pandas, NumPy, OpenPyXL, LangChain Experimental (create_pandas_dataframe_agent) y OpenAI API.
