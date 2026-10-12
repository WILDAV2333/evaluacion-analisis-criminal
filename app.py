import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime

# Configuración visual para móviles y escritorio
st.set_page_config(
    page_title="Evaluación - Proyecto Integrador",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------------
# 1. BASE DE DATOS LOCAL
# -------------------------------------------------------------
def get_db():
    conn = sqlite3.connect("calificaciones_defensa.db", check_same_thread=False)
    cursor = conn.cursor()
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
    conn.commit()
    return conn

conn = get_db()

# -------------------------------------------------------------
# 2. BAREMO EXACTO SEGÚN RÚBRICA (Rubrica Proyecto Final.xlsx)
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

# -------------------------------------------------------------
# 3. ESTRUCTURA DE PARALELOS, TRIBUNALES Y GRUPOS
# (Personaliza los nombres cuando tengas la nómina final)
# -------------------------------------------------------------
ESTRUCTURA_CURSO = {
    "Paralelo A": {
        "tribunal": [
            "Evaluador A1",
            "Evaluador A2",
            "Evaluador A3"
        ],
        "grupos": {
            "Grupo 1 - Caso A": [
                "Alumno Policial A1", "Alumno Policial A2", "Alumno Policial A3", "Alumno Policial A4", "Alumno Policial A5"
            ],
            "Grupo 2 - Caso B": [
                "Alumno Policial A6", "Alumno Policial A7", "Alumno Policial A8", "Alumno Policial A9", "Alumno Policial A10"
            ]
        }
    },
    "Paralelo B": {
        "tribunal": [
            "Evaluador B1",
            "Evaluador B2",
            "Evaluador B3"
        ],
        "grupos": {
            "Grupo 1 - Caso C": [
                "Alumno Policial B1", "Alumno Policial B2", "Alumno Policial B3", "Alumno Policial B4", "Alumno Policial B5"
            ],
            "Grupo 2 - Caso D": [
                "Alumno Policial B6", "Alumno Policial B7", "Alumno Policial B8", "Alumno Policial B9", "Alumno Policial B10"
            ]
        }
    }
}

# -------------------------------------------------------------
# 4. INTERFAZ: PESTAÑAS
# -------------------------------------------------------------
pestana_evaluar, pestana_admin = st.tabs(["📝 Formulario de Calificación", "🔒 Administrador"])

# --- PESTAÑA EVALUADOR ---
with pestana_evaluar:
    st.title("Defensa de Proyecto Integrador")
    st.subheader("Curso de Especialización en Análisis Criminal")
    
    # Selector inicial de Paralelo
    paralelo_sel = st.selectbox(
        "Seleccione el Paralelo:",
        list(ESTRUCTURA_CURSO.keys()),
        index=None,
        placeholder="-- Seleccione Paralelo A o B --"
    )
    
    col_eval, col_grp = st.columns(2)
    evaluador_activo = None
    grupo_activo = None
    alumnos_activos = []
    
    if paralelo_sel:
        datos_paralelo = ESTRUCTURA_CURSO[paralelo_sel]
        with col_eval:
            evaluador_activo = st.selectbox(
                f"Miembro del Tribunal ({paralelo_sel}):",
                datos_paralelo["tribunal"],
                index=None,
                placeholder="Seleccione su nombre..."
            )
        with col_grp:
            grupo_activo = st.selectbox(
                f"Grupo a Evaluar ({paralelo_sel}):",
                list(datos_paralelo["grupos"].keys()),
                index=None,
                placeholder="Seleccione el grupo..."
            )
        
        if grupo_activo:
            alumnos_activos = datos_paralelo["grupos"][grupo_activo]
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
        st.markdown("### 🗣️ PARTE II: SUSTENTACIÓN ORAL INDIVIDUAL (12 PUNTOS)")
        st.caption("Obligatorio calificar individualmente a cada uno de los 5 integrantes.")
        
        evaluaciones_orales = []
        
        if alumnos_activos:
            for i, alumno in enumerate(alumnos_activos):
                with st.expander(f"👤 Integrante {i+1}: {alumno}", expanded=True):
                    c_dom, c_preg = st.columns(2)
                    with c_dom:
                        dom_nivel = st.radio(
                            "Dominio del tema (Máx 6 pts):",
                            OPCIONES_NIVEL,
                            index=None,
                            horizontal=True,
                            key=f"dom_{i}"
                        )
                    with c_preg:
                        preg_nivel = st.radio(
                            "Interrogatorio (Máx 6 pts):",
                            OPCIONES_NIVEL,
                            index=None,
                            horizontal=True,
                            key=f"preg_{i}"
                        )
                    
                    if dom_nivel is None:
                        campos_faltantes.append(f"{alumno} ➔ Dominio del tema")
                    if preg_nivel is None:
                        campos_faltantes.append(f"{alumno} ➔ Interrogatorio")
                        
                    if dom_nivel and preg_nivel:
                        pts_dom = PUNTOS_CRITERIO["6.00"][dom_nivel]
                        pts_preg = PUNTOS_CRITERIO["6.00"][preg_nivel]
                        subtotal_oral = pts_dom + pts_preg
                        
                        evaluaciones_orales.append({
                            "alumno": alumno,
                            "dom_nivel": dom_nivel,
                            "pts_dom": pts_dom,
                            "preg_nivel": preg_nivel,
                            "pts_preg": pts_preg,
                            "subtotal_oral": subtotal_oral
                        })
        else:
            st.warning("⚠️ Debe seleccionar Paralelo y Grupo arriba para habilitar la nómina de alumnos.")
        
        enviar = st.form_submit_button("💾 Guardar y Enviar Evaluación del Grupo", use_container_width=True)
        
        if enviar:
            if not paralelo_sel:
                st.error("⚠️ Debe seleccionar el Paralelo en la parte superior.")
            elif not evaluador_activo:
                st.error("⚠️ Debe seleccionar su nombre como Miembro del Tribunal.")
            elif not grupo_activo:
                st.error("⚠️ Debe seleccionar el Grupo a Evaluar.")
            elif len(campos_faltantes) > 0:
                st.error("🚨 **EVALUACIÓN INCOMPLETA: Hay campos sin calificar.**")
                st.warning("Por favor complete los siguientes ítems antes de guardar:")
                for item in campos_faltantes:
                    st.write(f"❌ **{item}**")
            else:
                subtotal_escrito = sum(respuestas_escrito.values())
                cursor = conn.cursor()
                fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for ev in evaluaciones_orales:
                    nota_total_calc = round(subtotal_escrito + ev["subtotal_oral"], 2)
                    cursor.execute("""
                        INSERT INTO evaluaciones 
                        (fecha, paralelo, evaluador, grupo, alumno, nota_escrito, oral_dominio_nivel, 
                         oral_dominio_pts, oral_interrogatorio_nivel, oral_interrogatorio_pts, 
                         nota_oral, nota_total)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        fecha_reg, paralelo_sel, evaluador_activo, grupo_activo, ev["alumno"],
                        subtotal_escrito, ev["dom_nivel"], ev["pts_dom"],
                        ev["preg_nivel"], ev["pts_preg"], ev["subtotal_oral"], nota_total_calc
                    ))
                conn.commit()
                st.success(f"✅ ¡Evaluación del {paralelo_sel} registrada con éxito!")
                st.balloons()

# --- PESTAÑA ADMINISTRADOR ---
with pestana_admin:
    st.subheader("Consolidado de Calificaciones (Acceso Exclusivo)")
    password = st.text_input("Contraseña de Administrador:", type="password")
    
    if password == "AdminPolicial2026":
        df = pd.read_sql_query("SELECT * FROM evaluaciones", conn)
        if df.empty:
            st.info("No hay evaluaciones registradas hasta el momento.")
        else:
            filtro_admin = st.selectbox("Filtrar visualización por Paralelo:", ["Todos", "Paralelo A", "Paralelo B"])
            df_mostrar = df if filtro_admin == "Todos" else df[df["paralelo"] == filtro_admin]
            
            st.markdown(f"#### Promedios Finales por Alumno ({filtro_admin})")
            if not df_mostrar.empty:
                resumen = df_mostrar.groupby(["paralelo", "grupo", "alumno"]).agg(
                    Promedio_Escrito=("nota_escrito", "mean"),
                    Promedio_Oral=("nota_oral", "mean"),
                    Nota_Final=("nota_total", "mean"),
                    Tribunales_Evaluaron=("evaluador", "nunique")
                ).round(2).reset_index()
                st.dataframe(resumen, use_container_width=True)
                
                st.markdown("#### Detalle Completo de Calificaciones por Tribunal")
                st.dataframe(df_mostrar, use_container_width=True)
                
                csv = df_mostrar.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Descargar Acta en CSV ({filtro_admin})",
                    data=csv,
                    file_name=f"calificaciones_{filtro_admin.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info(f"No hay registros aún para {filtro_admin}.")
    elif password != "":
        st.error("Contraseña incorrecta.")
