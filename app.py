import streamlit as st
import pandas as pd
import sqlite3
import json
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
    
    # Migración automática: verificar si la columna 'paralelo' existe en la tabla previa
    cursor.execute("PRAGMA table_info(evaluaciones)")
    columnas = [col[1] for col in cursor.fetchall()]
    if "paralelo" not in columnas:
        cursor.execute("ALTER TABLE evaluaciones ADD COLUMN paralelo TEXT DEFAULT 'Paralelo A'")
        
    # Tabla para docentes/evaluadores configurados por el administrador
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS config_docentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paralelo TEXT,
            nombre TEXT
        )
    """)
    # Tabla para grupos y sus alumnos configurados por el administrador
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
    # Docentes
    cursor.execute("SELECT nombre FROM config_docentes WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    docentes = [row[0] for row in cursor.fetchall()]
    
    # Grupos y alumnos
    cursor.execute("SELECT grupo, alumnos_json FROM config_grupos WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    grupos_db = cursor.fetchall()
    
    grupos_dict = {}
    for g_nombre, alums_raw in grupos_db:
        try:
            grupos_dict[g_nombre] = json.loads(alums_raw)
        except Exception:
            grupos_dict[g_nombre] = []
            
    return docentes, grupos_dict

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
        st.markdown("### 🗣️ PARTE II: SUSTENTACIÓN ORAL INDIVIDUAL (12 PUNTOS)")
        st.caption("Obligatorio calificar individualmente a cada uno de los integrantes.")
        
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

# =============================================================
# --- PESTAÑA ADMINISTRADOR (GESTIÓN Y CONSOLIDADO) ---
# =============================================================
with pestana_admin:
    st.subheader("Panel de Control del Administrador")
    password = st.text_input("Ingrese la Contraseña:", type="password")
    
    if password == "AdminPolicial2026":
        tabs_admin = st.tabs(["📊 Consolidado de Notas", "⚙️ Gestión de Docentes y Grupos"])
        
        # SUB-PESTAÑA 1: CONSOLIDADO Y DESCARGA
        with tabs_admin[0]:
            df = pd.read_sql_query("SELECT * FROM evaluaciones", conn)
            if df.empty:
                st.info("No hay evaluaciones registradas hasta el momento.")
            else:
                if "paralelo" not in df.columns:
                    df["paralelo"] = "Paralelo A"
                    
                filtro_admin = st.selectbox("Filtrar por Paralelo:", ["Todos", "Paralelo A", "Paralelo B"])
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
                    
                    st.markdown("#### Detalle Completo de Calificaciones")
                    st.dataframe(df_mostrar, use_container_width=True)
                    
                    csv = df_mostrar.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"📥 Descargar Acta Completa ({filtro_admin}) en CSV",
                        data=csv,
                        file_name=f"calificaciones_{filtro_admin.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                else:
                    st.info(f"No hay registros aún para {filtro_admin}.")
                    
        # SUB-PESTAÑA 2: GESTIÓN DE DOCENTES, GRUPOS Y CURSANTES
        with tabs_admin[1]:
            st.markdown("### 👥 Configuración de Nóminas y Tribunales")
            paralelo_gestionar = st.selectbox("Seleccione el Paralelo a configurar:", PARALELOS_DISPONIBLES)
            
            c_docentes, c_grupos = st.columns(2)
            
            # --- Gestión de Docentes / Evaluadores ---
            with c_docentes:
                st.markdown(f"#### 1. Miembros del Tribunal ({paralelo_gestionar})")
                with st.form("form_agregar_docente", clear_on_submit=True):
                    nuevo_docente = st.text_input("Grado y Nombre del Evaluador (ej. Mayor Juan Pérez):")
                    btn_docente = st.form_submit_button("➕ Agregar Evaluador")
                    if btn_docente and nuevo_docente.strip():
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO config_docentes (paralelo, nombre) VALUES (?, ?)", (paralelo_gestionar, nuevo_docente.strip()))
                        conn.commit()
                        st.success(f"Evaluador agregado: {nuevo_docente.strip()}")
                        st.rerun()
                
                # Lista de docentes actuales con opción de eliminar
                cursor = conn.cursor()
                cursor.execute("SELECT id, nombre FROM config_docentes WHERE paralelo = ?", (paralelo_gestionar,))
                docs_actuales = cursor.fetchall()
                if docs_actuales:
                    st.markdown("**Evaluadores registrados:**")
                    for d_id, d_nom in docs_actuales:
                        col_d1, col_d2 = st.columns([4, 1])
                        col_d1.write(f"• {d_nom}")
                        if col_d2.button("🗑️", key=f"del_doc_{d_id}"):
                            cursor.execute("DELETE FROM config_docentes WHERE id = ?", (d_id,))
                            conn.commit()
                            st.rerun()
                else:
                    st.caption("No hay evaluadores agregados para este paralelo.")
                    
            # --- Gestión de Grupos y Alumnos ---
            with c_grupos:
                st.markdown(f"#### 2. Grupos y Alumnos ({paralelo_gestionar})")
                with st.form("form_agregar_grupo", clear_on_submit=True):
                    nom_grupo = st.text_input("Nombre del Grupo y Caso (ej. Grupo 1 - Caso Homicidios):")
                    al1 = st.text_input("Integrante 1 (Grado y Nombre):")
                    al2 = st.text_input("Integrante 2 (Grado y Nombre):")
                    al3 = st.text_input("Integrante 3 (Grado y Nombre):")
                    al4 = st.text_input("Integrante 4 (Grado y Nombre):")
                    al5 = st.text_input("Integrante 5 (Grado y Nombre):")
                    btn_grupo = st.form_submit_button("➕ Registrar Grupo y Alumnos")
                    
                    if btn_grupo:
                        alumnos_lista = [a.strip() for a in [al1, al2, al3, al4, al5] if a.strip()]
                        if not nom_grupo.strip():
                            st.error("Debe ingresar un nombre para el grupo.")
                        elif len(alumnos_lista) == 0:
                            st.error("Debe ingresar al menos un alumno para este grupo.")
                        else:
                            cursor = conn.cursor()
                            cursor.execute(
                                "INSERT INTO config_grupos (paralelo, grupo, alumnos_json) VALUES (?, ?, ?)",
                                (paralelo_gestionar, nom_grupo.strip(), json.dumps(alumnos_lista))
                            )
                            conn.commit()
                            st.success(f"Grupo registrado con {len(alumnos_lista)} integrantes.")
                            st.rerun()
                
                # Lista de grupos actuales con opción de eliminar
                cursor.execute("SELECT id, grupo, alumnos_json FROM config_grupos WHERE paralelo = ?", (paralelo_gestionar,))
                grps_actuales = cursor.fetchall()
                if grps_actuales:
                    st.markdown("**Grupos registrados:**")
                    for g_id, g_nom, g_alums_raw in grps_actuales:
                        alums = json.loads(g_alums_raw)
                        with st.expander(f"📁 {g_nom} ({len(alums)} integrantes)"):
                            for a in alums:
                                st.write(f"- {a}")
                            if st.button("Eliminar Grupo", key=f"del_grp_{g_id}"):
                                cursor.execute("DELETE FROM config_grupos WHERE id = ?", (g_id,))
                                conn.commit()
                                st.rerun()
                else:
                    st.caption("No hay grupos agregados para este paralelo.")
                    
    elif password != "":
        st.error("Contraseña incorrecta.")
