import streamlit as st
import pandas as pd
import sqlite3
import json
import io
import os
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
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS evaluaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            paralelo TEXT,
            evaluador TEXT,
            grupo TEXT,
            alumno TEXT,
            nota_escrito REAL,
            oral_sustentacion_nivel TEXT,
            oral_sustentacion_pts REAL,
            oral_interrogatorio_nivel TEXT,
            oral_interrogatorio_pts REAL,
            oral_estrategico_nivel TEXT,
            oral_estrategico_pts REAL,
            oral_diferenciador_nivel TEXT,
            oral_diferenciador_pts REAL,
            nota_oral REAL,
            nota_total REAL
        )
    """)
    
    cursor.execute("PRAGMA table_info(evaluaciones)")
    columnas = [col[1] for col in cursor.fetchall()]
    columnas_requeridas = [
        ("paralelo", "TEXT DEFAULT 'Paralelo A'"),
        ("oral_sustentacion_nivel", "TEXT"),
        ("oral_sustentacion_pts", "REAL DEFAULT 0"),
        ("oral_interrogatorio_nivel", "TEXT"),
        ("oral_interrogatorio_pts", "REAL DEFAULT 0"),
        ("oral_estrategico_nivel", "TEXT"),
        ("oral_estrategico_pts", "REAL DEFAULT 0"),
        ("oral_diferenciador_nivel", "TEXT"),
        ("oral_diferenciador_pts", "REAL DEFAULT 0")
    ]
    for col_nom, col_tipo in columnas_requeridas:
        if col_nom not in columnas:
            cursor.execute(f"ALTER TABLE evaluaciones ADD COLUMN {col_nom} {col_tipo}")
        
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS config_docentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paralelo TEXT,
            nombre TEXT
        )
    """)
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

def separar_grado_nombre(cadena_alumno):
    partes = cadena_alumno.strip().split()
    grados_comunes = [
        "sgto.", "sgto", "cbos.", "cbos", "poli.", "poli", "policía", "policia",
        "subt.", "subt", "tnte.", "tnte", "teniente", "capt.", "capt", "capitán",
        "capitan", "mayr.", "mayr", "mayor", "sgtos.", "sgtos", "sbop.", "sbop",
        "sbsp.", "sbsp", "sgtop.", "sgtop", "crnl.", "crnl"
    ]
    if len(partes) >= 2 and partes[0].lower() in grados_comunes:
        if len(partes) >= 3 and partes[1].lower() in ["segundo", "primero", "2do", "1ro"]:
            return f"{partes[0]} {partes[1]}", " ".join(partes[2:])
        return partes[0], " ".join(partes[1:])
    return "S/G", cadena_alumno.strip()

# -------------------------------------------------------------
# 2. DEFINICIÓN DETALLADA DE LA RÚBRICA (RUBRICA IICEAC.xlsx)
# -------------------------------------------------------------
# PARTE I: 7 Criterios x 3 Indicadores específicos cada uno (Total: 8 Pts)
CRITERIOS_DETALLADOS_ESCRITO = [
    {
        "criterio": "ENFOQUE LEGAL Y TIPIFICACIÓN",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Identificación clara del tipo penal aplicable al caso; sustento en normativa legal vigente (COIP u otra normativa pertinente).",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Relación coherente entre los hechos analizados y su tipificación jurídica.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Argumentación legal estructurada (elementos del delito: conducta, tipicidad, antijuridicidad, culpabilidad).",
                "max": 0.34,
                "valores": {"Excelente (100%)": 0.34, "Bueno (80%)": 0.27, "Básico (60%)": 0.20, "Deficiente (25%)": 0.09}
            }
        ]
    },
    {
        "criterio": "COMPRENSIÓN CRIMINOLÓGICA E HIPÓTESIS",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Habilidades para la formulación de hipótesis investigativas claras y verificables.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Identificación de patrones delictivos (modus operandi, temporalidad, recurrencia).",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Análisis del fenómeno desde una perspectiva criminológica (causas, motivaciones, comportamiento) y relación entre variables.",
                "max": 0.34,
                "valores": {"Excelente (100%)": 0.34, "Bueno (80%)": 0.27, "Básico (60%)": 0.20, "Deficiente (25%)": 0.09}
            }
        ]
    },
    {
        "criterio": "GESTIÓN Y ANÁLISIS DE DATOS",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Capacidad de recopilar, procesar, depurar y analizar información proveniente de diversas fuentes.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Uso de herramientas tecnológicas y metodologías de análisis de datos.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Orientadas a la generación de insumos analíticos que aporten valor al proceso investigativo.",
                "max": 0.34,
                "valores": {"Excelente (100%)": 0.34, "Bueno (80%)": 0.27, "Básico (60%)": 0.20, "Deficiente (25%)": 0.09}
            }
        ]
    },
    {
        "criterio": "ANÁLISIS GEOESPACIAL Y TÁCTICO",
        "total_pts": 1.50,
        "indicadores": [
            {
                "desc": "Capacidad de recopilar, procesar, depurar y analizar información geoespacial proveniente de diversas fuentes.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            },
            {
                "desc": "Uso de herramientas tecnológicas y metodologías de análisis táctico y mapas delictivos.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            },
            {
                "desc": "Orientadas a la generación de insumos analíticos tácticos y espaciales que aporten al proceso.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            }
        ]
    },
    {
        "criterio": "CIBERINTELIGENCIA Y OSINT",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Uso de fuentes abiertas (redes sociales, web, registros públicos).",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Aplicación de técnicas básicas de búsqueda, filtrado, verificación y validación de la información obtenida.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Integración de hallazgos OSINT dentro del análisis general plasmado en el informe.",
                "max": 0.34,
                "valores": {"Excelente (100%)": 0.34, "Bueno (80%)": 0.27, "Básico (60%)": 0.20, "Deficiente (25%)": 0.09}
            }
        ]
    },
    {
        "criterio": "PRODUCCIÓN DE INTELIGENCIA Y PROPUESTA",
        "total_pts": 1.50,
        "indicadores": [
            {
                "desc": "Elaboración de un producto de inteligencia estructurado, que contenga una síntesis clara de hallazgos relevantes.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            },
            {
                "desc": "Generación de recomendaciones operativas y/o estratégicas.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            },
            {
                "desc": "Mantenimiento de un enfoque práctico orientado a la toma de decisiones institucionales.",
                "max": 0.50,
                "valores": {"Excelente (100%)": 0.50, "Bueno (80%)": 0.40, "Básico (60%)": 0.30, "Deficiente (25%)": 0.13}
            }
        ]
    },
    {
        "criterio": "PRESENTACIÓN FORMAL DEL DOCUMENTO",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Estructura lógica del documento entregable físico (inicio, teoría de caso, desarrollo, conclusiones).",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Redacción clara, coherente y sin errores ortográficos.",
                "max": 0.33,
                "valores": {"Excelente (100%)": 0.33, "Bueno (80%)": 0.26, "Básico (60%)": 0.20, "Deficiente (25%)": 0.08}
            },
            {
                "desc": "Uso adecuado de tablas, gráficos, mapas e imágenes integradas al informe.",
                "max": 0.34,
                "valores": {"Excelente (100%)": 0.34, "Bueno (80%)": 0.27, "Básico (60%)": 0.20, "Deficiente (25%)": 0.09}
            }
        ]
    }
]

# PARTE II: 4 Criterios Oficiales Orales (Total: 12 Pts)
CRITERIOS_ORAL = [
    ("Sustentación (Exposición y Defensa Oral)", 3.00, "Dominio técnico, claridad, seguridad, argumentación sólida y uso de medios tecnológicos."),
    ("Interrogatorio y Resolución ante el Tribunal", 3.00, "Respuestas fundamentadas, justificación metodológica, aplomo y control profesional."),
    ("Análisis Estratégico y Prospectivo", 3.00, "Tendencias delictivas, proyección de escenarios futuros (desplazamiento, mutación) y análisis de riesgos."),
    ("Elemento Diferenciador y Pensamiento Crítico", 3.00, "Innovación, perspicacia analítica, capacidad de evitar sesgos y criterio propio.")
]

PUNTOS_ORAL = {
    "Excelente (100%)": 3.00,
    "Bueno (80%)": 2.40,
    "Básico (60%)": 1.80,
    "Deficiente (25%)": 0.75
}

OPCIONES_NIVEL = ["Excelente (100%)", "Bueno (80%)", "Básico (60%)", "Deficiente (25%)"]
PARALELOS_DISPONIBLES = ["Paralelo A", "Paralelo B"]

# -------------------------------------------------------------
# 3. INTERFAZ: PESTAÑAS
# -------------------------------------------------------------
pestana_evaluar, pestana_admin = st.tabs(["📝 Formulario de Calificación", "🔒 Administrador"])

# =============================================================
# --- PESTAÑA EVALUADOR ---
# =============================================================
with pestana_evaluar:
    banner_candidatos = [
        "banner.jpeg",
        "banner.jpg",
        "banner.png",
        "Banner - II Curso de Esp Ana Criminal.jpeg"
    ]
    banner_encontrado = None
    for b in banner_candidatos:
        if os.path.exists(b):
            banner_encontrado = b
            break
            
    if banner_encontrado:
        st.image(banner_encontrado, use_container_width=True)
    else:
        st.title("Defensa de Proyecto Integrador")
        st.subheader("II Curso de Especialización en Análisis Criminal")
        
    st.markdown("<h3 style='text-align: center; margin-top: 10px; margin-bottom: 25px;'>Defensa de Proyecto Integrador</h3>", unsafe_allow_html=True)
    
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
        st.markdown("### 📄 PARTE I: DESARROLLO DEL CASO — PRODUCTO ENTREGABLE FÍSICO (8 PUNTOS)")
        st.caption("Esta calificación aplica a los integrantes del grupo investigativo. Evalúe individualmente cada indicador.")
        
        respuestas_escrito = []
        campos_faltantes = []
        
        # Despliegue de los 7 criterios con sus 3 indicadores cada uno
        for c_idx, crit_obj in enumerate(CRITERIOS_DETALLADOS_ESCRITO):
            c_nombre = crit_obj["criterio"]
            c_tot = crit_obj["total_pts"]
            
            with st.container():
                st.markdown(f"#### 📌 {c_nombre} *(Puntaje Máximo: {c_tot:.2f} pts)*")
                
                for ind_idx, ind in enumerate(crit_obj["indicadores"]):
                    c_desc, c_sel = st.columns([3, 2])
                    with c_desc:
                        st.markdown(f"**Indicador {ind_idx + 1}:** {ind['desc']}")
                        st.caption(f"Valor Máximo: {ind['max']:.2f} pts")
                    with c_sel:
                        opc_sel = st.selectbox(
                            f"Nivel para Ind {ind_idx+1} de {c_nombre}",
                            OPCIONES_NIVEL,
                            index=None,
                            placeholder="-- Seleccione calificación --",
                            key=f"esc_{c_idx}_{ind_idx}",
                            label_visibility="collapsed"
                        )
                        if opc_sel is not None:
                            pts_obtenidos = ind["valores"][opc_sel]
                            respuestas_escrito.append(pts_obtenidos)
                        else:
                            campos_faltantes.append(f"{c_nombre} ➔ Indicador {ind_idx + 1}")
                    st.divider()

        # PARTE 2: SUSTENTACIÓN ORAL (INDIVIDUAL - 12 PUNTOS)
        num_alumnos = len(alumnos_activos)
        st.markdown(f"### 🗣️ PARTE II: SUSTENTACIÓN ORAL, DEFENSA E INTERROGATORIO (12 PUNTOS) — {num_alumnos} Integrantes")
        st.caption("Obligatorio calificar individualmente los 4 criterios a cada integrante del grupo.")
        
        evaluaciones_orales = []
        
        if alumnos_activos:
            for i, alumno in enumerate(alumnos_activos):
                with st.expander(f"👤 Integrante {i+1}: {alumno}", expanded=True):
                    al_notas = {}
                    al_completo = True
                    
                    cols_oral = st.columns(2)
                    for o_idx, (o_crit, o_peso, o_desc) in enumerate(CRITERIOS_ORAL):
                        with cols_oral[o_idx % 2]:
                            st.markdown(f"**{o_crit}** *(Máx: {o_peso:.2f} pts)*")
                            st.caption(o_desc)
                            calif_oral = st.radio(
                                f"Calif_{alumno}_{o_idx}",
                                OPCIONES_NIVEL,
                                index=None,
                                horizontal=True,
                                key=f"oral_{i}_{o_idx}",
                                label_visibility="collapsed"
                            )
                            if calif_oral is None:
                                campos_faltantes.append(f"{alumno} ➔ {o_crit}")
                                al_completo = False
                            else:
                                al_notas[o_idx] = {
                                    "nivel": calif_oral,
                                    "pts": PUNTOS_ORAL[calif_oral]
                                }
                            st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
                    
                    if al_completo:
                        subtotal_indiv = sum(n["pts"] for n in al_notas.values())
                        evaluaciones_orales.append({
                            "alumno": alumno,
                            "sust_nivel": al_notas[0]["nivel"],
                            "sust_pts": al_notas[0]["pts"],
                            "interrog_nivel": al_notas[1]["nivel"],
                            "interrog_pts": al_notas[1]["pts"],
                            "estrateg_nivel": al_notas[2]["nivel"],
                            "estrateg_pts": al_notas[2]["pts"],
                            "diferenc_nivel": al_notas[3]["nivel"],
                            "diferenc_pts": al_notas[3]["pts"],
                            "subtotal_oral": subtotal_indiv
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
                st.error("🚨 **EVALUACIÓN INCOMPLETA: Hay campos o indicadores sin calificar.**")
                st.warning("Por favor complete los siguientes ítems antes de guardar:")
                for item in campos_faltantes:
                    st.write(f"❌ **{item}**")
            else:
                subtotal_escrito = round(sum(respuestas_escrito), 2)
                cursor = conn.cursor()
                fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for ev in evaluaciones_orales:
                    nota_total_calc = round(subtotal_escrito + ev["subtotal_oral"], 2)
                    cursor.execute("""
                        INSERT INTO evaluaciones 
                        (fecha, paralelo, evaluador, grupo, alumno, nota_escrito, 
                         oral_sustentacion_nivel, oral_sustentacion_pts,
                         oral_interrogatorio_nivel, oral_interrogatorio_pts,
                         oral_estrategico_nivel, oral_estrategico_pts,
                         oral_diferenciador_nivel, oral_diferenciador_pts,
                         nota_oral, nota_total)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        fecha_reg, paralelo_sel, evaluador_activo, grupo_activo, ev["alumno"],
                        subtotal_escrito,
                        ev["sust_nivel"], ev["sust_pts"],
                        ev["interrog_nivel"], ev["interrog_pts"],
                        ev["estrateg_nivel"], ev["estrateg_pts"],
                        ev["diferenc_nivel"], ev["diferenc_pts"],
                        ev["subtotal_oral"], nota_total_calc
                    ))
                conn.commit()
                st.success(f"✅ ¡Evaluación del {paralelo_sel} registrada con éxito!")
                st.balloons()

# =============================================================
# --- PESTAÑA ADMINISTRADOR ---
# =============================================================
with pestana_admin:
    st.subheader("Panel de Control del Administrador")
    password = st.text_input("Ingrese la Contraseña:", type="password")
    
    if password == "AdminPolicial2026":
        tabs_admin = st.tabs(["📊 Consolidado de Notas", "⚙️ Gestión de Docentes y Grupos"])
        
        # SUB-PESTAÑA 1: CONSOLIDADO Y DESCARGA EXCEL (.XLSX)
        with tabs_admin[0]:
            df = pd.read_sql_query("SELECT * FROM evaluaciones", conn)
            if df.empty:
                st.info("No hay evaluaciones registradas hasta el momento.")
            else:
                if "paralelo" not in df.columns:
                    df["paralelo"] = "Paralelo A"
                    
                filtro_admin = st.selectbox("Filtrar por Paralelo:", ["Todos", "Paralelo A", "Paralelo B"])
                df_filtrado = df if filtro_admin == "Todos" else df[df["paralelo"] == filtro_admin]
                
                st.markdown(f"#### Promedios Finales por Alumno ({filtro_admin})")
                if not df_filtrado.empty:
                    resumen = df_filtrado.groupby(["paralelo", "grupo", "alumno"]).agg(
                        nota_escrito_prom=("nota_escrito", "mean"),
                        nota_oral_prom=("nota_oral", "mean"),
                        nota_total_prom=("nota_total", "mean")
                    ).round(2).reset_index()
                    
                    lista_excel = []
                    for idx, row in resumen.iterrows():
                        grado, nombres_apellidos = separar_grado_nombre(row["alumno"])
                        lista_excel.append({
                            "Ord.": idx + 1,
                            "Paralelo": row["paralelo"],
                            "Grupo": row["grupo"],
                            "Grado": grado,
                            "Nombres y Apellidos": nombres_apellidos,
                            "Calificación del Trabajo Escrito": row["nota_escrito_prom"],
                            "Calificación SUSTENTACIÓN ORAL INDIVIDUAL": row["nota_oral_prom"],
                            "Total Calificación Final": row["nota_total_prom"]
                        })
                    
                    df_oficial = pd.DataFrame(lista_excel)
                    st.dataframe(df_oficial, use_container_width=True)
                    
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                        df_oficial.to_excel(writer, index=False, sheet_name="Acta de Calificaciones")
                    excel_data = output_excel.getvalue()
                    
                    st.download_button(
                        label=f"📊 Descargar Acta Oficial en Excel (.xlsx) - {filtro_admin}",
                        data=excel_data,
                        file_name=f"Acta_Calificaciones_{filtro_admin.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.markdown("#### Detalle Completo de Registros Individuales")
                    st.dataframe(df_filtrado, use_container_width=True)
                else:
                    st.info(f"No hay registros aún para {filtro_admin}.")
            
            # --- ZONA DE MANTENIMIENTO: BORRAR PRUEBAS ---
            st.divider()
            with st.expander("⚠️ Zona de Mantenimiento: Vaciar Evaluaciones de Prueba"):
                st.warning("Esta acción eliminará únicamente las calificaciones registradas. No afectará a los evaluadores ni a los grupos configurados.")
                confirmar = st.checkbox("Confirmo que deseo borrar todas las calificaciones de prueba registradas.")
                if st.button("🗑️ Borrar Todas las Calificaciones", disabled=not confirmar):
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM evaluaciones")
                    conn.commit()
                    st.success("✅ Base de datos de calificaciones reiniciada a cero.")
                    st.rerun()
                    
        # SUB-PESTAÑA 2: GESTIÓN DE DOCENTES, GRUPOS Y CURSANTES DINÁMICOS
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
                    
            # --- Gestión Dinámica de Grupos y Alumnos ---
            with c_grupos:
                st.markdown(f"#### 2. Grupos y Alumnos ({paralelo_gestionar})")
                with st.form("form_agregar_grupo", clear_on_submit=True):
                    nom_grupo = st.text_input("Nombre del Grupo y Caso (ej. Grupo 1 - Caso Homicidios):")
                    alumnos_texto = st.text_area(
                        "Integrantes (escriba o pegue un nombre con grado por línea):",
                        help="Ejemplo:\\nSgto. Juan Pérez\\nCbos. Carlos Ruiz\\nPoli. Ana Gómez",
                        placeholder="Sgto. Juan Pérez\\nCbos. Carlos Ruiz\\nPoli. Ana Gómez\\nPoli. Luis Castro"
                    )
                    btn_grupo = st.form_submit_button("➕ Registrar Grupo y Alumnos")
                    
                    if btn_grupo:
                        alumnos_lista = [linea.strip() for linea in alumnos_texto.split("\n") if linea.strip()]
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
                            st.success(f"Grupo registrado exitosamente con {len(alumnos_lista)} integrantes.")
                            st.rerun()
                
                cursor.execute("SELECT id, grupo, alumnos_json FROM config_grupos WHERE paralelo = ?", (paralelo_gestionar,))
                grps_actuales = cursor.fetchall()
                if grps_actuales:
                    st.markdown("**Grupos registrados:**")
                    for g_id, g_nom, g_alums_raw in grps_actuales:
                        alums = json.loads(g_alums_raw)
                        with st.expander(f"📁 {g_nom} ({len(alums)} integrantes)"):
                            texto_actual = "\n".join(alums)
                            nuevo_texto_alums = st.text_area(
                                "Modificar integrantes (un nombre por línea):",
                                value=texto_actual,
                                key=f"edit_txt_{g_id}"
                            )
                            col_guardar, col_borrar = st.columns(2)
                            if col_guardar.button("💾 Guardar Cambios", key=f"btn_save_{g_id}"):
                                nueva_lista = [l.strip() for l in nuevo_texto_alums.split("\n") if l.strip()]
                                cursor.execute(
                                    "UPDATE config_grupos SET alumnos_json = ? WHERE id = ?",
                                    (json.dumps(nueva_lista), g_id)
                                )
                                conn.commit()
                                st.success("Integrantes actualizados.")
                                st.rerun()
                            if col_borrar.button("🗑️ Eliminar Grupo", key=f"del_grp_{g_id}"):
                                cursor.execute("DELETE FROM config_grupos WHERE id = ?", (g_id,))
                                conn.commit()
                                st.rerun()
                else:
                    st.caption("No hay grupos agregados para este paralelo.")
                    
    elif password != "":
        st.error("Contraseña incorrecta.")
