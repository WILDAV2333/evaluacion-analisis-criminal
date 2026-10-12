import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime

# Configuración visual para pantallas móviles y de escritorio
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

GRUPOS = {
    "Grupo 1 - Caso A": [
        "Alumno Policial 1", "Alumno Policial 2", "Alumno Policial 3", "Alumno Policial 4", "Alumno Policial 5"
    ],
    "Grupo 2 - Caso B": [
        "Alumno Policial 6", "Alumno Policial 7", "Alumno Policial 8", "Alumno Policial 9", "Alumno Policial 10"
    ]
}

EVALUADORES = [
    "Miembro del Tribunal 1",
    "Miembro del Tribunal 2",
    "Miembro del Tribunal 3"
]

# -------------------------------------------------------------
# 3. INTERFAZ: PESTAÑAS
# -------------------------------------------------------------
pestana_evaluar, pestana_admin = st.tabs(["📝 Formulario de Calificación", "🔒 Administrador"])

# --- PESTAÑA EVALUADOR ---
with pestana_evaluar:
    st.title("Defensa de Proyecto Integrador")
    st.subheader("Curso de Especialización en Análisis Criminal")
    
    col_eval, col_grp = st.columns(2)
    with col_eval:
        evaluador_activo = st.selectbox(
            "Miembro del Tribunal:",
            EVALUADORES,
            index=None,
            placeholder="Seleccione su nombre..."
        )
    with col_grp:
        grupo_activo = st.selectbox(
            "Grupo a Evaluar:",
            list(GRUPOS.keys()),
            index=None,
            placeholder="Seleccione el grupo..."
        )
        
    alumnos_activos = GRUPOS[grupo_activo] if grupo_activo else []
    
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
        st.caption("Obligatorio calificar a cada integrante según su intervención y respuesta a preguntas.")
        
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
                        campos_faltantes.append(f"{alumno} -> Dominio del tema")
                    if preg_nivel is None:
                        campos_faltantes.append(f"{alumno} -> Interrogatorio")
                        
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
            st.info("ℹ️ Primero debe seleccionar un Grupo arriba para cargar a los integrantes.")
        
        enviar = st.form_submit_button("💾 Guardar y Enviar Evaluación del Grupo", use_container_width=True)
        
        if enviar:
            # Validaciones previas
            if not evaluador_activo:
                st.error("⚠️ Debe seleccionar su nombre como Miembro del Tribunal en la parte superior.")
            elif not grupo_activo:
                st.error("⚠️ Debe seleccionar el Grupo a Evaluar en la parte superior.")
            elif len(campos_faltantes) > 0:
                st.error("🚨 **EVALUACIÓN INCOMPLETA: Faltan campos por calificar.**")
                st.warning("Por favor complete los siguientes ítems pendientes antes de enviar:")
                for item in campos_faltantes:
                    st.write(f"❌ **{item}**")
            else:
                # Todos los datos completos: proceder a guardar
                subtotal_escrito = sum(respuestas_escrito.values())
                cursor = conn.cursor()
                fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for ev in evaluaciones_orales:
                    nota_total_calc = round(subtotal_escrito + ev["subtotal_oral"], 2)
                    cursor.execute("""
                        INSERT INTO evaluaciones 
                        (fecha, evaluador, grupo, alumno, nota_escrito, oral_dominio_nivel, 
                         oral_dominio_pts, oral_interrogatorio_nivel, oral_interrogatorio_pts, 
                         nota_oral, nota_total)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        fecha_reg, evaluador_activo, grupo_activo, ev["alumno"],
                        subtotal_escrito, ev["dom_nivel"], ev["pts_dom"],
                        ev["preg_nivel"], ev["pts_preg"], ev["subtotal_oral"], nota_total_calc
                    ))
                conn.commit()
                st.success("✅ ¡Evaluación registrada con éxito! Todas las notas han sido guardadas.")
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
            st.markdown("#### Promedios Finales por Alumno")
            resumen = df.groupby(["grupo", "alumno"]).agg(
                Promedio_Escrito=("nota_escrito", "mean"),
                Promedio_Oral=("nota_oral", "mean"),
                Nota_Final=("nota_total", "mean"),
                Tribunales_Evaluaron=("evaluador", "nunique")
            ).round(2).reset_index()
            st.dataframe(resumen, use_container_width=True)
            
            st.markdown("#### Detalle Completo de Calificaciones por Tribunal")
            st.dataframe(df, use_container_width=True)
            
            # Descarga de Excel/CSV
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar Acta Completa (CSV)",
                data=csv,
                file_name=f"calificaciones_analisis_criminal_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    elif password != "":
        st.error("Contraseña incorrecta.")
