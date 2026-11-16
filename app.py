import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import base64

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from streamlit_js_eval import streamlit_js_eval

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA (DEBE SER LA PRIMERA LLAMADA)
# ---------------------------------------------------------
st.set_page_config(
    page_title="APP Jspreadsheet + AG Grid + Power Query + IA",
    page_icon=":material/code_blocks:",
    layout="wide"
)

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet" />
    <style>
    .main-header { 
        font-size: 2rem; 
        color: #1f77b4; 
        font-weight: 700; 
        margin-bottom: 4px; 
        display: flex; 
        align-items: center; 
        gap: 10px; 
    }
    .main-header .material-symbols-outlined {
        font-size: 2.2rem;
    }
    .sub-header { font-size: 1.25rem; color: #333333; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

st.markdown(
    '<p class="main-header"><span class="material-symbols-outlined">code_blocks</span> APP Jspreadsheet + AG Grid + Power Query + Agente IA</p>',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# GUÍA GENERAL INTERACTIVA (st.expander)
# ---------------------------------------------------------
with st.expander(":material/menu_book: Guía Rápida de Uso: ¿Cómo funciona cada herramienta de esta Suite?", expanded=False):
    g1, g2, g3 = st.columns(3)
    with g1:
        st.markdown("""
        **:material/folder_open: Barra Lateral (Importar y Configurar)**
        * **Cargar archivo:** Sube archivos `.xlsx`, `.csv`, `.json` o `.tsv`. Se cargarán automáticamente en todas las herramientas.
        * **Hoja en Blanco / Cargar Demo:** Reinicia la grilla o carga datos de ejemplo para probar la suite.
        * **OpenAI API Key:** Necesaria únicamente para activar el **Agente IA** en la pestaña 5.
        """)
        st.markdown("""
        **1️⃣ Carga & Fórmulas (Planilla JS)**
        * Funciona como Excel: escribe datos o fórmulas (`=SUM(A2:A9)`, `=AVERAGE()`, `=IF()`).
        * **Sincronización en 1 clic:** Tus cambios se guardan automáticamente al editar; solo haz clic en **Sincronizar cambios hacia AG Grid, Power Query e IA**.
        """)
    with g2:
        st.markdown("""
        **2️⃣ Sincronización & ETL (AG Grid)**
        * Grilla analítica de alto rendimiento sincronizada en tiempo real con Python.
        * Usa el panel lateral derecho (**Columns / Filters**) para filtrar, ordenar o agrupar columnas.
        """)
        st.markdown("""
        **3️⃣ Power Query (Transformaciones)**
        * Aplica limpieza de datos sin código: filtrar texto, quitar duplicados, pasar a mayúsculas, rellenar vacíos con cero o agrupar totales (*Group By*).
        * Haz clic en **Aplicar Paso ETL** para actualizar toda la suite.
        """)
    with g3:
        st.markdown("""
        **4️⃣ Tablas Dinámicas**
        * Selecciona dimensiones para **Filas**, **Columnas**, una métrica numérica (**Valores**) y la operación (`sum`, `mean`, `count`, `min`, `max`).
        """)
        st.markdown("""
        **5️⃣ Agente IA (LangChain) & 6️⃣ Exportar**
        * **Agente IA:** Haz preguntas en español sobre tus datos. El agente genera y ejecuta código Python sobre tu tabla actual.
        * **Exportar:** Descarga el resultado final en `.xlsx`, `.csv`, `.json` o `.tsv`.
        """)

# ---------------------------------------------------------
# 2. FUNCIONES AUXILIARES Y PUENTE PYTHON <-> JSPREADSHEET
# ---------------------------------------------------------
def crear_df_vacio(filas=20, columnas=10):
    cols = [chr(65 + i) if i < 26 else f"Col_{i+1}" for i in range(columnas)]
    return pd.DataFrame("", index=range(filas), columns=cols)

def df_a_matriz_jspreadsheet(df: pd.DataFrame):
    if df is None or df.empty:
        return [["" for _ in range(10)] for _ in range(20)], [chr(65 + i) for i in range(10)]

    tiene_encabezados = not all(
        (len(str(c)) == 1 and 'A' <= str(c) <= 'Z') or str(c).startswith("Unnamed")
        for c in df.columns
    )

    matriz = []
    if tiene_encabezados:
        matriz.append([str(c) for c in df.columns])

    for row in df.itertuples(index=False):
        fila_limpia = []
        for val in row:
            if pd.isna(val):
                fila_limpia.append("")
            elif isinstance(val, (int, float, np.number)) and not isinstance(val, bool):
                fila_limpia.append(float(val) if isinstance(val, float) else int(val))
            else:
                fila_limpia.append(str(val))
        matriz.append(fila_limpia)

    num_cols = max(10, len(matriz[0]) if matriz else 10)
    for fila in matriz:
        while len(fila) < num_cols:
            fila.append("")
    while len(matriz) < 20:
        matriz.append([""] * num_cols)

    headers = [chr(65 + i) if i < 26 else f"Col_{i+1}" for i in range(num_cols)]
    return matriz, headers

@st.cache_data(show_spinner=False)
def construir_iframe_jspreadsheet(matriz_json: str, headers_json: str) -> str:
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <script src="https://bossanova.uk/jspreadsheet/v4/jexcel.js"></script>
        <script src="https://jsuites.net/v4/jsuites.js"></script>
        <link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v4/jexcel.css" type="text/css" />
        <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />
        <style>
            body {{ margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #fff; }}
            #estado-barra {{
                background: #f8fafc; color: #334155; padding: 6px 14px;
                border-bottom: 1px solid #e2e8f0; font-size: 12px;
                display: flex; justify-content: space-between; align-items: center;
            }}
            #badge-auto {{ color: #059669; font-weight: 600; }}
            #contenedor-tabla {{ width: 100%; height: calc(100vh - 34px); overflow: auto; }}
        </style>
    </head>
    <body>
        <div id="estado-barra">
            <span>⚡ <b>Planilla JS</b> (Soporta fórmulas Excel: <code>=SUM(A2:A9)</code>, <code>=AVERAGE()</code>, <code>=IF()</code>)</span>
            <span id="badge-auto">🟢 Autoguardado en memoria listo</span>
        </div>
        <div id="contenedor-tabla">
            <div id="spreadsheet"></div>
        </div>

        <script>
            var dataInicial = {matriz_json};
            var columnasHeaders = {headers_json}.map(function(h) {{
                return {{ title: h, width: 120 }};
            }});

            function guardarSnapshotSilencioso(instancia) {{
                try {{
                    var datosCalculados = instancia.getData(false);
                    window.parent.localStorage.setItem("jspreadsheet_sync_payload", JSON.stringify(datosCalculados));
                    var badge = document.getElementById("badge-auto");
                    badge.innerText = "🟢 Cambios listos para sincronizar (" + new Date().toLocaleTimeString() + ")";
                }} catch (e) {{
                    console.error("Error guardando snapshot:", e);
                }}
            }}

            var tabla = jspreadsheet(document.getElementById('spreadsheet'), {{
                data: dataInicial,
                columns: columnasHeaders,
                minDimensions: [10, 20],
                tableOverflow: true,
                tableWidth: "100%",
                tableHeight: "525px",
                allowInsertRow: true,
                allowInsertColumn: true,
                allowDeleteRow: true,
                allowDeleteColumn: true,
                columnSorting: true,
                search: true,
                defaultColAlign: 'left',
                onload: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                onchange: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                oninsertrow: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                ondeleterow: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                oninsertcolumn: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                ondeletecolumn: function(instance) {{ guardarSnapshotSilencioso(instance); }},
                onpaste: function(instance) {{ guardarSnapshotSilencioso(instance); }}
            }});
        </script>
    </body>
    </html>
    """
    b64_html = base64.b64encode(html_code.encode("utf-8")).decode("utf-8")
    return f"data:text/html;base64,{b64_html}"

def matriz_js_a_df(matriz_raw: list) -> pd.DataFrame:
    if not matriz_raw:
        return crear_df_vacio()

    filas_limpias = [
        fila for fila in matriz_raw
        if any(str(celda).strip() != "" and celda is not None for celda in fila)
    ]

    if not filas_limpias:
        return crear_df_vacio()

    max_col = 0
    for f in filas_limpias:
        for idx, val in enumerate(f):
            if str(val).strip() != "" and val is not None:
                max_col = max(max_col, idx + 1)

    filas_recortadas = [f[:max_col] for f in filas_limpias]

    # Verificar si la primera fila parece encabezado (contiene texto no numérico)
    primera_fila_es_header = any(
        isinstance(c, str) and c.strip() != "" and not c.strip().replace(".", "", 1).replace("-", "", 1).isdigit()
        for c in filas_recortadas[0]
    )

    if len(filas_recortadas) > 1 and primera_fila_es_header:
        posibles_cols = [str(c).strip() if str(c).strip() != "" else f"Col_{i+1}" for i, c in enumerate(filas_recortadas[0])]
        cols_unicas = []
        for i, c in enumerate(posibles_cols):
            cols_unicas.append(f"{c}_{i}" if c in cols_unicas else c)
        df = pd.DataFrame(filas_recortadas[1:], columns=cols_unicas)
    else:
        cols = [chr(65 + i) if i < 26 else f"Col_{i+1}" for i in range(max_col)]
        df = pd.DataFrame(filas_recortadas, columns=cols)

    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass

    return df

# ---------------------------------------------------------
# 3. ESTADO GLOBAL DE LA APLICACIÓN
# ---------------------------------------------------------
if "df_activo" not in st.session_state:
    st.session_state["df_activo"] = crear_df_vacio(20, 10)

if "sincronizado_aggrid" not in st.session_state:
    st.session_state["sincronizado_aggrid"] = False

if "ultimo_archivo" not in st.session_state:
    st.session_state["ultimo_archivo"] = None

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "sync_counter" not in st.session_state:
    st.session_state["sync_counter"] = 0

if "pendiente_lectura_js" not in st.session_state:
    st.session_state["pendiente_lectura_js"] = False

# ---------------------------------------------------------
# 4. BARRA LATERAL: IMPORTAR ARCHIVOS Y CONFIGURACIÓN IA
# ---------------------------------------------------------
st.sidebar.header(":material/folder_open: 1. Importar Datos")
archivo = st.sidebar.file_uploader(
    "Cargar archivo", 
    type=["csv", "xlsx", "xls", "json", "tsv", "txt"],
    key=f"uploader_{st.session_state['uploader_key']}",
    help="Carga datos iniciales en la Planilla JS y AG Grid."
)

if archivo is not None:
    nombre = archivo.name.lower()
    sep_csv = "Auto"
    if nombre.endswith(".csv"):
        sep_csv = st.sidebar.selectbox("Separador CSV", options=["Auto", ",", ";", "|"], index=0)

    id_arch = f"{archivo.name}_{archivo.size}_{sep_csv}"
    if st.session_state["ultimo_archivo"] != id_arch:
        try:
            if nombre.endswith((".xlsx", ".xls")):
                excel_file = pd.ExcelFile(archivo)
                hoja_sel = st.sidebar.selectbox("Seleccionar hoja", excel_file.sheet_names) if len(excel_file.sheet_names) > 1 else excel_file.sheet_names[0]
                st.session_state["df_activo"] = pd.read_excel(excel_file, sheet_name=hoja_sel)
            elif nombre.endswith(".csv"):
                st.session_state["df_activo"] = pd.read_csv(archivo, sep=None, engine="python") if sep_csv == "Auto" else pd.read_csv(archivo, sep=sep_csv)
            elif nombre.endswith((".tsv", ".txt")):
                st.session_state["df_activo"] = pd.read_csv(archivo, sep="\t")
            elif nombre.endswith(".json"):
                st.session_state["df_activo"] = pd.read_json(archivo)
                
            st.session_state["ultimo_archivo"] = id_arch
            st.session_state["sincronizado_aggrid"] = True
            st.sidebar.success(f"¡'{archivo.name}' cargado!")
        except Exception as e:
            st.sidebar.error(f"Error al leer archivo: {e}")

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    if st.button(":material/cleaning_services: Hoja en Blanco", use_container_width=True):
        st.session_state["df_activo"] = crear_df_vacio(20, 10)
        st.session_state["ultimo_archivo"] = None
        st.session_state["sincronizado_aggrid"] = False
        st.session_state["uploader_key"] += 1  # Limpia también el file_uploader
        st.rerun()
with col_sb2:
    if st.button(":material/change_circle: Cargar Demo", use_container_width=True):
        st.session_state["df_activo"] = pd.DataFrame({
            "Legajo": [101, 102, 103, 104, 105, 106, 107, 108],
            "Empleado": ["Ana Gómez", "Carlos Pérez", "Lucía Fernández", "Martín López", "Sofía Ruiz", "Diego Torres", "Valeria Sosa", "Jorge Benítez"],
            "Sector": ["Administración", "Ventas", "Administración", "Producción", "Ventas", "Producción", "Administración", "Ventas"],
            "Categoria": ["Senior", "Junior", "Senior", "Semi-Senior", "Junior", "Senior", "Semi-Senior", "Junior"],
            "Sueldo_Basico": [850000.0, 520000.0, 850000.0, 680000.0, 520000.0, 850000.0, 680000.0, 520000.0],
            "Horas_Extra": [5, 12, 0, 8, 15, 2, 4, 10]
        })
        st.session_state["sincronizado_aggrid"] = True
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.header(":material/support_agent: 2. Configuración Agente IA")
openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password", placeholder="sk-...")
modelo_ia = st.sidebar.selectbox("Modelo", ["gpt-4o-mini", "gpt-4o"], index=0)

# ---------------------------------------------------------
# 5. FRAGMENTOS AISLADOS CON AYUDA CONTEXTUAL (st.popover)
# ---------------------------------------------------------
@st.fragment
def render_tab_power_query():
    c_tit, c_ayuda = st.columns([5, 1])
    with c_tit:
        st.markdown('<p class="sub-header">Constructor de Pasos ETL (Estilo Power Query)</p>', unsafe_allow_html=True)
    with c_ayuda:
        with st.popover(":material/live_help: Ayuda ETL"):
            st.markdown("""
            **¿Cómo usar Power Query?**
            1. Elige una operación en el menú desplegable izquierdo.
            2. Revisa el resultado en la **Vista Previa** de la derecha.
            3. Haz clic en **Aplicar Paso ETL** para actualizar toda la aplicación.
            """)

    df_actual = st.session_state["df_activo"]
    col_q1, col_q2 = st.columns([1, 2])
    cols_actuales = df_actual.columns.tolist()

    with col_q1:
        st.markdown("#### :material/settings: Operaciones de Transformación")
        tipo_etl = st.selectbox(
            "Seleccionar Transformación",
            ["Filtrar por Columna", "Eliminar Duplicados", "Convertir Texto a Mayúsculas", "Rellenar Vacíos / Nulos con Cero", "Agrupar por (Group By)"]
        )

        df_trans = df_actual.copy()

        if tipo_etl == "Filtrar por Columna":
            c_fil = st.selectbox("Columna a filtrar", options=cols_actuales)
            val_fil = st.text_input("Texto o valor requerido (dejar vacío para descartar nulos)")
            if val_fil:
                df_trans = df_trans[df_trans[c_fil].astype(str).str.contains(val_fil, na=False, case=False)]
            else:
                df_trans = df_trans[df_trans[c_fil].astype(str).str.strip() != ""]

        elif tipo_etl == "Eliminar Duplicados":
            dup_cols = st.multiselect("Considerar columnas", options=cols_actuales)
            df_trans = df_trans.drop_duplicates(subset=dup_cols) if dup_cols else df_trans.drop_duplicates()

        elif tipo_etl == "Convertir Texto a Mayúsculas":
            c_up = st.selectbox("Columna de texto", options=cols_actuales)
            df_trans[c_up] = df_trans[c_up].astype(str).str.upper()

        elif tipo_etl == "Rellenar Vacíos / Nulos con Cero":
            df_trans = df_trans.fillna(0).replace("", 0)

        elif tipo_etl == "Agrupar por (Group By)":
            c_grp = st.selectbox("Agrupar por dimensión", options=cols_actuales)
            df_num_check = df_actual.copy()
            for c in df_num_check.columns:
                try:
                    df_num_check[c] = pd.to_numeric(df_num_check[c])
                except (ValueError, TypeError):
                    pass
            num_cols_etl = df_num_check.select_dtypes(include=['number']).columns.tolist()
            if num_cols_etl:
                c_met = st.selectbox("Métrica numérica a sumar", options=num_cols_etl)
                df_trans = df_num_check.groupby(c_grp)[c_met].sum().reset_index()
            else:
                st.warning("Se requiere al menos una columna numérica para agrupar.")

        if st.button(":material/check_box: Aplicar Paso ETL a AG Grid y Planilla JS", use_container_width=True):
            st.session_state["df_activo"] = df_trans
            st.session_state["sincronizado_aggrid"] = True
            st.rerun(scope="app")  # Reinicia toda la app para sincronizar Jspreadsheet y AG Grid

    with col_q2:
        st.markdown("#### :material/visibility: Vista Previa del Resultado Transformado")
        st.dataframe(df_trans, use_container_width=True)
        st.info(f"Filas resultantes: {len(df_trans)}")

@st.fragment
def render_tab_pivote():
    c_tit, c_ayuda = st.columns([5, 1])
    with c_tit:
        st.markdown('<p class="sub-header">Generador de Tablas Dinámicas</p>', unsafe_allow_html=True)
    with c_ayuda:
        with st.popover(":material/live_help: Ayuda Pivote"):
            st.markdown("""
            **¿Cómo armar una Tabla Dinámica?**
            * **Filas:** Categorías por las que quieres agrupar verticalmente (ej. `Sector`).
            * **Columnas:** Subcategorías opcionales en horizontal (ej. `Categoria`).
            * **Valores:** Columna numérica a calcular (ej. `Sueldo_Basico`).
            * **Agregación:** Función matemática (`sum`, `mean`, `count`, `min`, `max`).
            """)

    df_actual = st.session_state["df_activo"]
    cols_actuales = df_actual.columns.tolist()
    
    df_temp_piv = df_actual.copy()
    for c in df_temp_piv.columns:
        conv = pd.to_numeric(df_temp_piv[c], errors='coerce')
        if conv.notna().sum() > 0:
            df_temp_piv[c] = conv

    cols_num_piv = df_temp_piv.select_dtypes(include=['number']).columns.tolist()

    if cols_actuales and cols_num_piv:
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            filas_p = st.multiselect("Filas (Index)", options=cols_actuales, default=[cols_actuales[2]] if len(cols_actuales)>2 else [cols_actuales[0]])
        with p2:
            cols_p = st.multiselect("Columnas (Columns)", options=cols_actuales, default=[])
        with p3:
            val_p = st.selectbox("Valores (Métrica)", options=cols_num_piv, index=0)
        with p4:
            agg_p = st.selectbox("Agregación", options=["sum", "mean", "count", "min", "max"])

        if filas_p:
            try:
                pivot_res = pd.pivot_table(
                    df_temp_piv,
                    index=filas_p,
                    columns=cols_p if cols_p else None,
                    values=val_p,
                    aggfunc=agg_p,
                    fill_value=0
                )
                st.markdown("#### :material/table_convert: Matriz Dinámica Generada")
                st.dataframe(pivot_res, use_container_width=True)
            except Exception as ex:
                st.warning(f"No se pudo generar la tabla dinámica: {ex}")
        else:
            st.warning("Selecciona al menos una dimensión de fila.")
    else:
        st.info(":material/sync_arrow_up: Sincroniza datos numéricos desde la Planilla JS o AG Grid para activar las tablas dinámicas.")

@st.fragment
def render_tab_agente(api_key, modelo):
    c_tit, c_ayuda = st.columns([5, 1])
    with c_tit:
        st.markdown('<p class="sub-header">Laboratorio de Agentes IA (LangChain)</p>', unsafe_allow_html=True)
    with c_ayuda:
        with st.popover(":material/live_help: Ayuda Agente IA"):
            st.markdown("""
            **¿Qué puedes pedirle al Agente?**
            * **Cálculos complejos:** *"Calcula el costo total de horas extra si cada hora vale $5.000."*
            * **Filtros y rankings:** *"Muestra los 3 empleados con mayor sueldo básico."*
            * **Auditoría:** *"¿Hay filas duplicadas o valores faltantes en la tabla?"*
            """)

    df_actual = st.session_state["df_activo"]
    col_ag1, col_ag2 = st.columns([1.2, 1])

    with col_ag1:
        st.markdown("#### :material/question_exchange: Pregunta al Agente en Lenguaje Natural")
        prompt_ia = st.text_area(
            "Escribe tu consulta sobre los datos sincronizados:",
            height=140,
            placeholder="Ejemplos:\n- ¿Cuál es el sector con mayor sueldo básico total?\n- Calcula el promedio de horas extra de los empleados sénior.\n- Detecta valores atípicos o resume la tabla actual."
        )

        if st.button(":material/touch_double_2: Ejecutar Consulta con Agente", use_container_width=True):
            if not api_key:
                st.error(":material/warning: Por favor, ingresa tu OpenAI API Key en la barra lateral.")
            elif not prompt_ia.strip():
                st.warning("Escribe una instrucción para el agente.")
            else:
                with st.spinner("Cargando motor IA y analizando los datos..."):
                    try:
                        from langchain_openai import ChatOpenAI
                        from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent

                        df_para_ia = df_actual.copy()
                        for col in df_para_ia.columns:
                            try:
                                df_para_ia[col] = pd.to_numeric(df_para_ia[col])
                            except (ValueError, TypeError):
                                pass

                        llm = ChatOpenAI(temperature=0, model=modelo, api_key=api_key)
                        agent = create_pandas_dataframe_agent(
                            llm,
                            df_para_ia,
                            verbose=True,
                            agent_type="openai-tools",
                            allow_dangerous_code=True
                        )
                        respuesta_ia = agent.invoke({"input": prompt_ia})
                        
                        st.markdown("#### :material/reply_all: Respuesta del Agente:")
                        st.success(respuesta_ia["output"])
                    except Exception as e:
                        st.error(f"Error al ejecutar el agente: {e}")

    with col_ag2:
        st.markdown("#### :material/data_table: Datos Actuales Visibles para el Agente")
        st.dataframe(df_actual, use_container_width=True)

@st.fragment
def render_tab_export():
    st.markdown('<p class="sub-header">Exportar Planilla en Múltiples Formatos</p>', unsafe_allow_html=True)
    df_actual = st.session_state["df_activo"]
    nombre_archivo = st.text_input("Nombre base del archivo", value="suite_datos_procesados")
    
    ex1, ex2, ex3, ex4 = st.columns(4)
    with ex1:
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
            df_actual.to_excel(writer, index=False, sheet_name="Datos")
        st.download_button(":material/data_table: Descargar .XLSX", buf_xlsx.getvalue(), f"{nombre_archivo}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with ex2:
        delim = st.radio("Delimitador CSV", [",", ";"], horizontal=True, key="delim_exp")
        csv_bytes = df_actual.to_csv(index=False, sep=delim).encode("utf-8-sig")
        st.download_button(":material/data_table: Descargar .CSV", csv_bytes, f"{nombre_archivo}.csv", "text/csv", use_container_width=True)
    with ex3:
        json_str = df_actual.to_json(orient="records", indent=2, force_ascii=False)
        st.download_button(":material/data_table: Descargar .JSON", json_str.encode("utf-8"), f"{nombre_archivo}.json", "application/json", use_container_width=True)
    with ex4:
        tsv_bytes = df_actual.to_csv(index=False, sep="\t").encode("utf-8")
        st.download_button(":material/data_table: Descargar .TSV", tsv_bytes, f"{nombre_archivo}.tsv", "text/tab-separated-values", use_container_width=True)

# ---------------------------------------------------------
# 6. PESTAÑAS PRINCIPALES
# ---------------------------------------------------------
tab_js, tab_aggrid, tab_query, tab_pivote, tab_agente, tab_export = st.tabs([
    "1️⃣ Carga & Fórmulas (Planilla JS)",
    "2️⃣ Sincronización & ETL (AG Grid)",
    "3️⃣ Power Query (Transformaciones)",
    "4️⃣ Tablas Dinámicas",
    "5️⃣ Agente IA (LangChain)",
    "6️⃣ Exportar Datos"
])

df_actual = st.session_state["df_activo"]

# TAB 1: PLANILLA JS RÁPIDA CON SINCRONIZACIÓN EN 1 SOLO CLIC
with tab_js:
    col_btn_sync, col_estado, col_pop = st.columns([1.8, 2.2, 0.5])
    
    with col_btn_sync:
        if st.button(":material/cached: Sincronizar cambios hacia AG Grid, Power Query e IA", type="primary", use_container_width=True):
            st.session_state["sync_counter"] += 1
            st.session_state["pendiente_lectura_js"] = True

    with col_estado:
        if st.session_state["sincronizado_aggrid"]:
            st.success(":material/library_add_check: Datos sincronizados y activos en AG Grid, ETL e IA.")
        else:
            st.info(":material/edit_square: Edita celdas o fórmulas abajo y pulsa el botón azul para sincronizar en 1 clic.")

    with col_pop:
        with st.popover(":material/live_help: Ayuda"):
            st.markdown("""
            **Sincronización en 1 solo clic:**
            * Cada vez que editas una celda o fórmula (`=SUM(...)`), la grilla guarda tus cambios automáticamente.
            * Pulsa el botón azul **Sincronizar cambios...** para pasar los valores ya calculados a **AG Grid, Power Query y el Agente IA**.
            """)

    if st.session_state["pendiente_lectura_js"]:
        raw_json = streamlit_js_eval(
            js_expressions="window.localStorage.getItem('jspreadsheet_sync_payload')",
            key=f"leer_js_1clic_{st.session_state['sync_counter']}"
        )
        if raw_json is not None:
            st.session_state["pendiente_lectura_js"] = False
            try:
                matriz = json.loads(raw_json)
                st.session_state["df_activo"] = matriz_js_a_df(matriz)
                st.session_state["sincronizado_aggrid"] = True
                st.toast(":material/cached: ¡Sincronizado en 1 clic con AG Grid, Power Query y Agente IA!")
                st.rerun()
            except Exception as e:
                st.error(f"Error al sincronizar datos: {e}")

    matriz_datos, headers_cols = df_a_matriz_jspreadsheet(df_actual)
    data_uri_iframe = construir_iframe_jspreadsheet(
        json.dumps(matriz_datos, ensure_ascii=False),
        json.dumps(headers_cols)
    )
    st.iframe(data_uri_iframe, height=565)

# TAB 2: AG GRID
with tab_aggrid:
    c_tit, c_ayuda = st.columns([5, 1])
    with c_tit:
        st.markdown('<p class="sub-header">Grilla Analítica AG Grid (Sincronización en Tiempo Real con Python)</p>', unsafe_allow_html=True)
    with c_ayuda:
        with st.popover(":material/live_help: Ayuda AG Grid"):
            st.markdown("""
            **Funciones de AG Grid:**
            * **Edición directa:** Haz doble clic en cualquier celda para modificar su valor; se guarda al instante.
            * **Panel lateral derecho:** Haz clic en `Columns` o `Filters` en el borde derecho de la tabla para filtrar valores o activar agrupamientos.
            """)

    gb = GridOptionsBuilder.from_dataframe(df_actual)
    gb.configure_default_column(editable=True, groupable=True, filterable=True, sortable=True, resizable=True)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
    gb.configure_side_bar()
    grid_options = gb.build()

    grid_response = AgGrid(
        df_actual,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        theme="streamlit",
        height=480,
        key="aggrid_principal"
    )

    df_aggrid_actualizado = pd.DataFrame(grid_response["data"])
    if not df_aggrid_actualizado.empty and not df_aggrid_actualizado.equals(df_actual):
        st.session_state["df_activo"] = df_aggrid_actualizado
        st.toast(":material/cached: Cambios en AG Grid sincronizados con Python y el Agente IA")

# TABS 3 A 6 (FRAGMENTOS AISLADOS)
with tab_query:
    render_tab_power_query()

with tab_pivote:
    render_tab_pivote()

with tab_agente:
    render_tab_agente(openai_api_key, modelo_ia)

with tab_export:
    render_tab_export()