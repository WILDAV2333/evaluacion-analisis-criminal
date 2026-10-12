import streamlit as st
import pandas as pd
import sqlite3
import json
import io
from datetime import datetime

# Configuración visual para pantallas móviles y de escritorio
st.set_page_config(
    page_title="Evaluación - Proyecto Integrador",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------------
# 1. BASE DE DATOS LOCAL Y MIGRACIÓN AUTOMÁTICA
# -------------------------------------------------------------
def get_db():
    conn = sqlite3.connect("calificaciones_defensa.db", check_same_thread=False)
    cursor = conn.cursor()
    
    # Tabla de evaluaciones
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS evaluaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            paralelo TEXT,
            evaluador TEXT,
            grupo TEXT,
            alumno TEXT,
            nota_escrito REAL,
            oral_dominio_nivel TEXT,
            oral_dominio_pts REAL,
            oral_interrogatorio_nivel TEXT,
            oral_interrogatorio_pts REAL,
            nota_oral REAL,
            nota_total REAL
        )
    """)
    
    cursor.execute("PRAGMA table_info(evaluaciones)")
    columnas = [col[1] for col in cursor.fetchall()]
    if "paralelo" not in columnas:
        cursor.execute("ALTER TABLE evaluaciones ADD COLUMN paralelo TEXT DEFAULT 'Paralelo A'")
        
    # Tabla para docentes/evaluadores
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS config_docentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paralelo TEXT,
            nombre TEXT
        )
    """)
    # Tabla para grupos y alumnos en formato JSON
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS config_grupos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paralelo TEXT,
            grupo TEXT,
            alumnos_json TEXT
        )
    """)
    conn.commit()
    return conn

conn = get_db()

# Cargar configuración activa desde la Base de Datos
def cargar_configuracion(paralelo):
    cursor = conn.cursor()
    cursor.execute("SELECT nombre FROM config_docentes WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    docentes = [row[0] for row in cursor.fetchall()]
    
    cursor.execute("SELECT grupo, alumnos_json FROM config_grupos WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    grupos_db = cursor.fetchall()
    
    grupos_dict = {}
    for g_nombre, alums_raw in grupos_db:
        try:
            grupos_dict[g_nombre] = json.loads(alums_raw)
        except Exception:
            grupos_dict[g_nombre] = []
            
    return docentes, grupos_dict

# Función para separar Grado y Nombres/Apellidos
def extraer_grado_y_nombres(nombre_completo):
    partes = nombre_completo.strip().split(maxsplit=1)
    grados_conocidos = {
        "sgto", "sgto.", "sgtos", "sgtos.", "cbos", "cbos.", "poli", "poli.",
        "policia", "policía", "tnte", "tnte.", "cptn", "cptn.", "myor", "myor.",
        "subt", "subt.", "subp", "subp.", "subo", "subo.", "crnl", "crnl."
    }
    if len(partes) > 1 and (partes[0].lower() in grados_conocidos or partes[0].endswith(".")):
        return partes[0], partes[1]
    return "-", nombre_completo.strip()

# -------------------------------------------------------------
# 2. BAREMO EXACTO SEGÚN RÚBRICA
# -------------------------------------------------------------
PUNTOS_CRITERIO = {
    "1.00": {"Deficiente": 0.25, "Básico": 0.60, "Bueno": 0.80, "Excelente": 1.00},
    "0.75": {"Deficiente": 0.19, "Básico": 0.45, "Bueno": 0.60, "Excelente": 0.75},
    "0.50": {"Deficiente": 0.13, "Básico": 0.30, "Bueno": 0.40, "Excelente": 0.50},
    "6.00": {"Deficiente": 1.50, "Básico": 3.60, "Bueno": 4.80, "Excelente": 6.00}
}

CRITERIOS_ESCRITO = [
    ("Enfoque legal, doctrinario y ético", "0.75"),
    ("Comprensión criminológica e hipótesis", "1.00"),
    ("Gestión y análisis de datos", "1.00"),
    ("Análisis geoespacial y táctico", "1.00"),
    ("Análisis estratégico y prospectivo", "1.00"),
    ("Ciberinteligencia y OSINT", "1.00"),
    ("Producción de inteligencia y propuesta", "0.75"),
    ("Presentación del informe", "0.50"),
    ("Sustentación", "0.50"),
    ("Elemento Diferenciador", "0.50")
]

OPCIONES_NIVEL = ["Excelente", "Bueno", "Básico", "Deficiente"]
PARALELOS_DISPONIBLES = ["Paralelo A", "Paralelo B"]

# -------------------------------------------------------------
# 3. INTERFAZ: PESTAÑAS
# -------------------------------------------------------------
pestana_evaluar, pestana_admin = st.tabs(["📝 Formulario de Calificación", "🔒 Administrador"])

# =============================================================
# --- PESTAÑA EVALUADOR (SOLO LECTURA DE LISTAS) ---
# =============================================================
with pestana_evaluar:
    st.title("Defensa de Proyecto Integrador")
    st.subheader("Curso de Especialización en Análisis Criminal")
    
    paralelo_sel = st.selectbox(
        "Seleccione el Paralelo:",
        PARALELOS_DISPONIBLES,
        index=None,
        placeholder="-- Seleccione Paralelo A o B --"
    )
    
    col_eval, col_grp = st.columns(2)
    evaluador_activo = None
    grupo_activo = None
    alumnos_activos = []
    
    if paralelo_sel:
        docentes_paralelo, grupos_paralelo = cargar_configuracion(paralelo_sel)
        
        with col_eval:
            if docentes_paralelo:
                evaluador_activo = st.selectbox(
                    f"Miembro del Tribunal ({paralelo_sel}):",
                    docentes_paralelo,
                    index=None,
                    placeholder="Seleccione su nombre..."
                )
            else:
                st.warning(f"No hay evaluadores asignados al {paralelo_sel} por el Administrador.")
                
        with col_grp:
            if grupos_paralelo:
                grupo_activo = st.selectbox(
                    f"Grupo a Evaluar ({paralelo_sel}):",
                    list(grupos_paralelo.keys()),
                    index=None,
                    placeholder="Seleccione el grupo..."
                )
                if grupo_activo:
                    alumnos_activos = grupos_paralelo[grupo_activo]
            else:
                st.warning(f"No hay grupos registrados en el {paralelo_sel} por el Administrador.")
    else:
        st.info("👆 Seleccione primero el Paralelo para cargar el Tribunal y los Grupos correspondientes.")

    with st.form("form_calificacion"):
        # PARTE 1: TRABAJO ESCRITO (GRUPAL - 8 PUNTOS)
        st.markdown("### 📄 PARTE I: TRABAJO ESCRITO (8 PUNTOS)")
        st.caption("Esta calificación aplica a los integrantes del grupo investigativo. Todos los campos son obligatorios.")
        
        cols_escrito = st.columns(2)
        respuestas_escrito = {}
        campos_faltantes = []
        
        for idx, (crit_nombre, peso) in enumerate(CRITERIOS_ESCRITO):
            col_target = cols_escrito[idx % 2]
            with col_target:
                nivel = st.selectbox(
                    f"{crit_nombre} (Máx: {peso} pts)",
                    OPCIONES_NIVEL,
                    index=None,
                    placeholder="-- Seleccione calificación --",
                    key=f"esc_{idx}"
                )
                if nivel is not None:
                    respuestas_escrito[crit_nombre] = PUNTOS_CRITERIO[peso][nivel]
                else:
                    campos_faltantes.append(f"Trabajo Escrito: {crit_nombre}")
                    
        st.divider()
        
        # PARTE 2: SUSTENTACIÓN ORAL (INDIVIDUAL - 12 PUNTOS)
        num_alumnos = len(alumnos_activos)
        st.markdown(f"### 🗣️ PARTE II: SUSTENTACIÓN ORAL INDIVIDUAL (12 PUNTOS) — {num_alumnos} Integrantes")
        st.caption("Obligatorio calificar individualmente a cada uno de los integrantes del grupoPara generar directamente el archivo en formato **Excel (`.xlsx`)** con el formato exacto requerido, se necesitan dos ajustes:

1. **Agregar `openpyxl` a `requirements.txt`**: Esta es la librería estándar de Python para crear y descargar archivos `.xlsx` reales.
2. **Actualizar la función de descarga en `app.py`**: El sistema procesa el texto ingresado en el alumno (separando automáticamente el grado de los nombres/apellidos si están escritos juntos, o ubicándolos en sus columnas), numera con un contador ordinal (`Ord.`), calcula los promedios finales sobre los 20 puntos y genera el Excel listo para firma y archivo.

---

### Paso 1: Actualizar `requirements.txt` en GitHub

Entra a tu archivo **`requirements.txt`** en GitHub, dale al lápiz (editar) y déjalo exactamente así:

```text
streamlit
pandas
openpyxl
