import streamlit as st
import pandas as pd
import sqlite3
import json
import io
import os
from datetime import datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

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
            detalles_escrito_json TEXT,
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
        ("detalles_escrito_json", "TEXT"),
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
            tema_caso TEXT,
            alumnos_json TEXT
        )
    """)
    cursor.execute("PRAGMA table_info(config_grupos)")
    cols_grp = [col[1] for col in cursor.fetchall()]
    if "tema_caso" not in cols_grp:
        cursor.execute("ALTER TABLE config_grupos ADD COLUMN tema_caso TEXT DEFAULT ''")

    conn.commit()
    return conn

conn = get_db()

def cargar_configuracion(paralelo):
    cursor = conn.cursor()
    cursor.execute("SELECT nombre FROM config_docentes WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    docentes = [row[0] for row in cursor.fetchall()]
    
    cursor.execute("SELECT grupo, alumnos_json, tema_caso FROM config_grupos WHERE paralelo = ? ORDER BY id ASC", (paralelo,))
    grupos_db = cursor.fetchall()
    
    grupos_dict = {}
    temas_dict = {}
    for g_nombre, alums_raw, tema in grupos_db:
        try:
            grupos_dict[g_nombre] = json.loads(alums_raw)
        except Exception:
            grupos_dict[g_nombre] = []
        temas_dict[g_nombre] = tema if tema else "Caso Práctico Integrador"
            
    return docentes, grupos_dict, temas_dict

# -------------------------------------------------------------
# FUNCIÓN DE SEPARACIÓN DE GRADOS POLICIALES EXACTOS
# -------------------------------------------------------------
def separar_grado_nombre(cadena_alumno):
    partes = cadena_alumno.strip().split(maxsplit=1)
    if not partes:
        return "S/G", ""
    
    primer_token = partes[0].strip()
    token_lower = primer_token.lower()
    
    mapa_grados = {
        "cptn.": "Cptn.", "cptn": "Cptn.",
        "tnte.": "Tnte.", "tnte": "Tnte.",
        "sbte.": "Sbte.", "sbte": "Sbte.", "subt.": "Sbte.", "subt": "Sbte.",
        "sgop.": "Sgop.", "sgop": "Sgop.", "sgto.": "Sgop.", "sgto": "Sgop.",
        "sgos.": "Sgos.", "sgos": "Sgos.",
        "cbop.": "Cbop.", "cbop": "Cbop.",
        "cbos.": "Cbos.", "cbos": "Cbos.",
        "poli.": "Poli.", "poli": "Poli.", "policia": "Poli.", "policía": "Poli.",
        "myor.": "Myor.", "myor": "Myor.", "mayor": "Myor.",
        "crnl.": "Crnl.", "crnl": "Crnl.", "coronel": "Crnl.",
        "gral.": "Gral.", "gral": "Gral."
    }
    
    if token_lower in mapa_grados:
        grado_oficial = mapa_grados[token_lower]
        resto_nombres = partes[1].strip() if len(partes) > 1 else ""
        return grado_oficial, resto_nombres
    elif (token_lower.rstrip(".") + ".") in mapa_grados:
        grado_oficial = mapa_grados[token_lower.rstrip(".") + "."]
        resto_nombres = partes[1].strip() if len(partes) > 1 else ""
        return grado_oficial, resto_nombres
        
    return "S/G", cadena_alumno.strip()

# -------------------------------------------------------------
# 2. DEFINICIÓN DETALLADA DE LA RÚBRICA (5%, 50%, 75%, 100%)
# -------------------------------------------------------------
OPCIONES_NIVEL = ["Excelente", "Bueno", "Básico", "Deficiente"]
PARALELOS_DISPONIBLES = ["Paralelo A", "Paralelo B"]

CRITERIOS_DETALLADOS_ESCRITO = [
    {
        "criterio": "ENFOQUE LEGAL Y TIPIFICACIÓN",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Identificación clara del tipo penal aplicable al caso; sustento en normativa legal vigente (COIP u otra normativa pertinente).",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Relación coherente entre los hechos analizados y su tipificación jurídica.",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Argumentación legal estructurada (elementos del delito: conducta, tipicidad, antijuridicidad, culpabilidad).",
                "max": 0.34,
                "valores": {"Excelente": 0.34, "Bueno": 0.25, "Básico": 0.16, "Deficiente": 0.01}
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
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Identificación de patrones delictivos (modus operandi, temporalidad, recurrencia).",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Análisis del fenómeno desde una perspectiva criminológica (causas, motivaciones, comportamiento) y la relación entre variables (víctima, victimario, entorno).",
                "max": 0.34,
                "valores": {"Excelente": 0.34, "Bueno": 0.25, "Básico": 0.16, "Deficiente": 0.01}
            }
        ]
    },
    {
        "criterio": "ANÁLISIS DE INVESTIGACIÓN CRIMINAL",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Análisis del plan de investigación o presentan un plan de investigación.",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Realizan el mapeo de jerarquías (tipo de estructura criminal, nodos y enlaces (facilitadores)).",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Evaluación procedimental (robustez de las pruebas o evidencias presentadas).",
                "max": 0.34,
                "valores": {"Excelente": 0.34, "Bueno": 0.25, "Básico": 0.16, "Deficiente": 0.01}
            }
        ]
    },
    {
        "criterio": "ANÁLISIS GEOESPACIAL Y TÁCTICO",
        "total_pts": 1.50,
        "indicadores": [
            {
                "desc": "Capacidad de recopilar, procesar, depurar y analizar información proveniente de diversas fuentes.",
                "max": 0.50,
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.03}
            },
            {
                "desc": "Uso de herramientas tecnológicas y metodologías de análisis de datos.",
                "max": 0.50,
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.03}
            },
            {
                "desc": "Orientadas a la generación de insumos analíticos que aporten valor al proceso investigativo.",
                "max": 0.50,
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.02}
            }
        ]
    },
    {
        "criterio": "OSINT Y CIBERINTELIGENCIA",
        "total_pts": 1.00,
        "indicadores": [
            {
                "desc": "Uso de fuentes abiertas (redes sociales, web, registros públicos).",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Aplicación de técnicas básicas de búsqueda, filtrado, verificación y validación de la información obtenida.",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Integración de hallazgos OSINT dentro del análisis general plasmado en el informe.",
                "max": 0.34,
                "valores": {"Excelente": 0.34, "Bueno": 0.25, "Básico": 0.16, "Deficiente": 0.01}
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
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.03}
            },
            {
                "desc": "Generación de recomendaciones operativas y/o estratégicas.",
                "max": 0.50,
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.03}
            },
            {
                "desc": "Mantenimiento de un enfoque práctico orientado a la toma de decisiones institucionales.",
                "max": 0.50,
                "valores": {"Excelente": 0.50, "Bueno": 0.38, "Básico": 0.25, "Deficiente": 0.02}
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
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Redacción clara, coherente y sin errores ortográficos.",
                "max": 0.33,
                "valores": {"Excelente": 0.33, "Bueno": 0.25, "Básico": 0.17, "Deficiente": 0.02}
            },
            {
                "desc": "Uso adecuado de tablas, gráficos, mapas e imágenes integradas al informe.",
                "max": 0.34,
                "valores": {"Excelente": 0.34, "Bueno": 0.25, "Básico": 0.16, "Deficiente": 0.01}
            }
        ]
    }
]

CRITERIOS_ORAL = [
    ("Sustentación (Exposición y Defensa Oral)", 3.00, "Dominio técnico del tema expuesto, claridad, seguridad, argumentación sólida y material audiovisual de apoyo."),
    ("Interrogatorio y Resolución ante el Tribunal", 3.00, "Respuestas precisas, fundamentadas, justificación metodológica, aplomo y seguridad profesional."),
    ("Análisis Estratégico y Prospectivo (Defensa de Escenarios)", 3.00, "Identificación de tendencias delictivas, proyección de escenarios futuros y formulación de alertas."),
    ("Elemento Diferenciador (Valor Agregado y Pensamiento Crítico)", 3.00, "Innovación, perspicacia analítica, perspectiva crítica, evitar sesgos y criterio propio.")
]

PUNTOS_ORAL = {
    "Excelente": 3.00,
    "Bueno": 2.25,
    "Básico": 1.50,
    "Deficiente": 0.15
}

# -------------------------------------------------------------
# FUNCIÓN GENERADORA DEL ACTA OFICIAL DOCX COMPLETA
# -------------------------------------------------------------
def generar_acta_docx_oficial(plantilla_path, paralelo, grupo_nombre, tema_caso, fecha_str, integrantes, tribunal_nombres, notas_escrito_21, desglose_alumnos):
    doc = docx.Document(plantilla_path)
    
    # 1. Encabezado de la primera página
    doc.paragraphs[5].text = f"Integrantes del Grupo Nro: {grupo_nombre}                           Paralelo: {paralelo}"
    
    for i in range(6):
        p_idx = 6 + i
        if i < len(integrantes):
            doc.paragraphs[p_idx].text = f"{i+1}. {integrantes[i]}"
        else:
            doc.paragraphs[p_idx].text = f"{i+1}. ----------------------------------------"
            
    doc.paragraphs[12].text = f"Tema del Caso Analizado: {tema_caso}"
    doc.paragraphs[13].text = f"Fecha: {fecha_str}"
    
    for i in range(3):
        p_idx = 15 + i
        if i < len(tribunal_nombres):
            doc.paragraphs[p_idx].text = f"{i+1}. {tribunal_nombres[i]}"
        else:
            doc.paragraphs[p_idx].text = f"{i+1}. ----------------------------------------"
            
    # 2. Tabla 0: Parte I (Entregable Físico - 21 Indicadores)
    t0 = doc.tables[0]
    t0.rows[11].cells[1].text = "ANÁLISIS DE INVESTIGACIÓN CRIMINAL"
    t0.rows[12].cells[1].text = "Análisis del plan de investigación o presentan un plan de investigación"
    t0.rows[13].cells[1].text = "Realizan el mapeo de jerarquías (tipo de estructura criminal, nodos y enlaces (facilitadores))"
    t0.rows[14].cells[1].text = "Evaluación procedimental (robustez de las pruebas o evidencias presentadas)"
    t0.rows[21].cells[1].text = "OSINT Y CIBERINTELIGENCIA"
    
    mapa_crit_t0 = [
        ([2, 3, 4], 5),
        ([7, 8, 9], 10),
        ([12, 13, 14], 15),
        ([17, 18, 19], 20),
        ([22, 23, 24], 25),
        ([27, 28, 29], 30),
        ([32, 33, 34], 35)
    ]
    
    tot_escrito_acum = 0.0
    ind_counter = 0
    for ind_rows, tot_row in mapa_crit_t0:
        subtot_crit = 0.0
        for r in ind_rows:
            val = float(notas_escrito_21[ind_counter]) if ind_counter < len(notas_escrito_21) else 0.0
            t0.rows[r].cells[2].text = f"{val:.2f}"
            subtot_crit += val
            ind_counter += 1
        t0.rows[tot_row].cells[2].text = f"{subtot_crit:.2f}"
        tot_escrito_acum += subtot_crit
        
    t0.rows[36].cells[2].text = f"{tot_escrito_acum:.2f} / 8 puntos"
    
    # 3. Tabla 1: Sustentación Oral - Promedio General de Grupo
    t1 = doc.tables[1]
    mapa_crit_t1 = [
        ([2, 3, 4], 5),
        ([7, 8, 9], 10),
        ([12, 13, 14], 15),
        ([17, 18, 19], 20)
    ]
    
    # Promedios orales del grupo para la tabla base
    if desglose_alumnos:
        avg_sust = sum(al["sust"] for al in desglose_alumnos) / len(desglose_alumnos)
        avg_interrog = sum(al["interrog"] for al in desglose_alumnos) / len(desglose_alumnos)
        avg_estrateg = sum(al["estrateg"] for al in desglose_alumnos) / len(desglose_alumnos)
        avg_diferenc = sum(al["diferenc"] for al in desglose_alumnos) / len(desglose_alumnos)
    else:
        avg_sust = avg_interrog = avg_estrateg = avg_diferenc = 0.0
        
    notas_oral_grupo = [avg_sust, avg_interrog, avg_estrateg, avg_diferenc]
    tot_oral_acum = 0.0
    for c_idx, (ind_rows, tot_row) in enumerate(mapa_crit_t1):
        crit_val = float(notas_oral_grupo[c_idx])
        val_ind = crit_val / 3.0
        for r in ind_rows:
            t1.rows[r].cells[2].text = f"{val_ind:.2f}"
        t1.rows[tot_row].cells[2].text = f"{crit_val:.2f}"
        tot_oral_acum += crit_val
        
    t1.rows[21].cells[2].text = f"{tot_oral_acum:.2f} / 12 puntos"
    
    # 4. Tabla 2: Total General del Grupo
    t2 = doc.tables[2]
    tot_general = tot_escrito_acum + tot_oral_acum
    t2.rows[0].cells[1].text = f"{tot_general:.2f} / 20 puntos"
    
    # 5. INSERCIÓN DE LA TABLA INDIVIDUALIZADA DE SUSTENTACIÓN ORAL
    # Insertar párrafo de título antes de la tabla de firmas
    p_individual = doc.add_paragraph()
    p_individual.paragraph_format.space_before = Pt(14)
    p_individual.paragraph_format.space_after = Pt(6)
    r_ind = p_individual.add_run("CUADRO DE CALIFICACIÓN INDIVIDUALIZADA DE LA SUSTENTACIÓN ORAL Y NOTA FINAL")
    r_ind.bold = True
    r_ind.font.size = Pt(10)
    r_ind.font.name = "Arial"
    p_individual.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Crear tabla con 10 columnas oficiales
    columnas_indiv = [
        "Ord.", "Grado", "Nombres y Apellidos", 
        "Sustentación\n(3.00)", "Interrogatorio\n(3.00)", 
        "Análisis Estrat.\n(3.00)", "Elem. Diferenc.\n(3.00)", 
        "TOTAL ORAL\n(12.00)", "TRABAJO ESCRITO\n(8.00)", "CALIFICACIÓN FINAL\n(20.00)"
    ]
    
    num_filas = len(desglose_alumnos) + 1
    t_indiv = doc.add_table(rows=num_filas, cols=10)
    t_indiv.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_indiv.style = "Table Grid"
    
    # Cabecera de la tabla individual
    hdr_cells = t_indiv.rows[0].cells
    for c_i, col_name in enumerate(columnas_indiv):
        hdr_cells[c_i].text = col_name
        shading = parse_xml(r'<w:shd {} w:fill="E2EFDA"/>'.format(nsdecls('w')))
        hdr_cells[c_i]._tc.get_or_add_tcPr().append(shading)
        p = hdr_cells[c_i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(7.5)
            run.font.name = "Arial"
            
    # Llenado de filas de cada integrante
    for row_i, al in enumerate(desglose_alumnos, start=1):
        row_cells = t_indiv.rows[row_i].cells
        grado, nombres = separar_grado_nombre(al["nombre"])
        
        datos_fila = [
            str(row_i),
            grado,
            nombres,
            f"{al['sust']:.2f}",
            f"{al['interrog']:.2f}",
            f"{al['estrateg']:.2f}",
            f"{al['diferenc']:.2f}",
            f"{al['total_oral']:.2f}",
            f"{tot_escrito_acum:.2f}",
            f"{(al['total_oral'] + tot_escrito_acum):.2f}"
        ]
        
        for c_i, valor in enumerate(datos_fila):
            row_cells[c_i].text = valor
            p = row_cells[c_i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_i == 2 else WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(8)
                run.font.name = "Arial"
                if c_i in [7, 8, 9]:
                    run.font.bold = True

    # 6. Tabla 3: Cuadro de Firmas Oficial (3 Miembros del Tribunal)
    t3 = doc.tables[3]
    t3.rows[0].cells[0].text = "\n\n\n\n___________________________\nFirma: Evaluador 1\nNombres: " + (tribunal_nombres[0] if len(tribunal_nombres) > 0 else "………………………")
    t3.rows[0].cells[1].text = "\n\n\n\n___________________________\nFirma: Evaluador 2\nNombres: " + (tribunal_nombres[1] if len(tribunal_nombres) > 1 else "………………………")
    
    # Fila 1 es combinada única: NO tocar cells[1] para no borrar cells[0]
    t3.rows[1].cells[0].text = "\n\n\n\n___________________________\nFirma: Evaluador 3\nNombres: " + (tribunal_nombres[2] if len(tribunal_nombres) > 2 else "………………………")
    
    # Centrar texto de las firmas
    for row in t3.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.size = Pt(8.5)
                    run.font.name = "Arial"

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

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
    tema_activo = ""
    
    if paralelo_sel:
        docentes_paralelo, grupos_paralelo, temas_paralelo = cargar_configuracion(paralelo_sel)
        
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
                    tema_activo = temas_paralelo.get(grupo_activo, "Caso Práctico Integrador")
            else:
                st.warning(f"No hay grupos registrados en el {paralelo_sel} por el Administrador.")
    else:
        st.info("👆 Seleccione primero el Paralelo para cargar el Tribunal y los Grupos correspondientes.")

    with st.form("form_calificacion"):
        st.markdown("### 📄 PARTE I: DESARROLLO DEL CASO — PRODUCTO ENTREGABLE FÍSICO (8 PUNTOS)")
        st.caption("Esta calificación aplica a los integrantes del grupo investigativo. Evalúe individualmente cada indicador.")
        
        respuestas_escrito = []
        campos_faltantes = []
        
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
                detalles_escrito_json_str = json.dumps(respuestas_escrito)
                cursor = conn.cursor()
                fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for ev in evaluaciones_orales:
                    nota_total_calc = round(subtotal_escrito + ev["subtotal_oral"], 2)
                    cursor.execute("""
                        INSERT INTO evaluaciones 
                        (fecha, paralelo, evaluador, grupo, alumno, nota_escrito, detalles_escrito_json,
                         oral_sustentacion_nivel, oral_sustentacion_pts,
                         oral_interrogatorio_nivel, oral_interrogatorio_pts,
                         oral_estrategico_nivel, oral_estrategico_pts,
                         oral_diferenciador_nivel, oral_diferenciador_pts,
                         nota_oral, nota_total)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        fecha_reg, paralelo_sel, evaluador_activo, grupo_activo, ev["alumno"],
                        subtotal_escrito, detalles_escrito_json_str,
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
        tabs_admin = st.tabs(["📊 Consolidado de Notas", "📄 Generar Actas Word (.docx)", "⚙️ Gestión de Docentes y Grupos"])
        
        # SUB-PESTAÑA 1: CONSOLIDADO Y EXCEL (.XLSX)
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
                        label=f"📊 Descargar Resumen en Excel (.xlsx) - {filtro_admin}",
                        data=excel_data,
                        file_name=f"Acta_Calificaciones_{filtro_admin.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.info(f"No hay registros aún para {filtro_admin}.")
            
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

        # SUB-PESTAÑA 2: GENERACIÓN DE ACTAS OFICIALES EN WORD (.DOCX)
        with tabs_admin[1]:
            st.markdown("### 🖨️ Generación de Actas Físicas Oficiales (Formato Institucional Word)")
            st.caption("Consolida automáticamente las notas emitidas por los miembros del tribunal calificador.")
            
            plantilla_oficial = "INDICADORES PARA EVALUACIÓN DEL EJERCICIO PRACTICO.docx"
            if not os.path.exists(plantilla_oficial):
                st.error(f"⚠️ No se encontró el archivo de plantilla `{plantilla_oficial}` en el repositorio.")
                st.info("Asegúrese de subir la plantilla Word con ese nombre exacto a su repositorio en GitHub.")
            else:
                par_acta = st.selectbox("Seleccione el Paralelo:", PARALELOS_DISPONIBLES, key="par_acta_sel")
                docentes_p, grupos_p, temas_p = cargar_configuracion(par_acta)
                
                if grupos_p:
                    grp_acta = st.selectbox("Seleccione el Grupo:", list(grupos_p.keys()), key="grp_acta_sel")
                    
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT evaluador, alumno, detalles_escrito_json, 
                               oral_sustentacion_pts, oral_interrogatorio_pts, 
                               oral_estrategico_pts, oral_diferenciador_pts, 
                               nota_oral, fecha
                        FROM evaluaciones
                        WHERE paralelo = ? AND grupo = ?
                    """, (par_acta, grp_acta))
                    registros = cursor.fetchall()
                    
                    if registros:
                        evaluadores_registrados = list(dict.fromkeys(r[0] for r in registros))
                        st.success(f"📋 Evaluaciones registradas de {len(evaluadores_registrados)} evaluador(es): {', '.join(evaluadores_registrados)}")
                        
                        # Priorizar la lista de los 3 docentes configurados en el tribunal
                        tribunal_para_acta = docentes_p[:3] if len(docentes_p) >= 3 else evaluadores_registrados
                        
                        # Promediar los 21 indicadores del trabajo escrito
                        listas_indicadores = []
                        for r in registros:
                            if r[2]:
                                try:
                                    listas_indicadores.append(json.loads(r[2]))
                                except Exception:
                                    pass
                                    
                        notas_21_prom = [0.0] * 21
                        if listas_indicadores:
                            num_listas = len(listas_indicadores)
                            for i in range(21):
                                notas_21_prom[i] = round(sum(l[i] for l in listas_indicadores if i < len(l)) / num_listas, 2)
                        
                        # Calcular promedio oral individual por cada alumno del grupo
                        integrantes_acta = grupos_p.get(grp_acta, [])
                        desglose_alumnos = []
                        for al_nombre in integrantes_acta:
                            evals_al = [r for r in registros if r[1] == al_nombre]
                            if evals_al:
                                p_sust = round(sum(r[3] for r in evals_al) / len(evals_al), 2)
                                p_interrog = round(sum(r[4] for r in evals_al) / len(evals_al), 2)
                                p_estrateg = round(sum(r[5] for r in evals_al) / len(evals_al), 2)
                                p_diferenc = round(sum(r[6] for r in evals_al) / len(evals_al), 2)
                                p_tot_oral = round(sum(r[7] for r in evals_al) / len(evals_al), 2)
                            else:
                                p_sust = p_interrog = p_estrateg = p_diferenc = p_tot_oral = 0.0
                                
                            desglose_alumnos.append({
                                "nombre": al_nombre,
                                "sust": p_sust,
                                "interrog": p_interrog,
                                "estrateg": p_estrateg,
                                "diferenc": p_diferenc,
                                "total_oral": p_tot_oral
                            })
                        
                        fecha_acta = registros[0][8] if registros[0][8] else datetime.now().strftime("%d de %B de %Y")
                        tema_acta = temas_p.get(grp_acta, "Proyecto Integrador de Análisis Criminal")
                        
                        acta_bytes = generar_acta_docx_oficial(
                            plantilla_path=plantilla_oficial,
                            paralelo=par_acta,
                            grupo_nombre=grp_acta,
                            tema_caso=tema_acta,
                            fecha_str=fecha_acta,
                            integrantes=integrantes_acta,
                            tribunal_nombres=tribunal_para_acta,
                            notas_escrito_21=notas_21_prom,
                            desglose_alumnos=desglose_alumnos
                        )
                        
                        st.download_button(
                            label=f"📥 Descargar Acta Oficial en Word (.docx) - {grp_acta}",
                            data=acta_bytes,
                            file_name=f"Acta_{par_acta.replace(' ', '_')}_{grp_acta.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.warning(f"Aún no hay calificaciones registradas para el {grp_acta}.")
                else:
                    st.info(f"No hay grupos registrados en el {par_acta}.")

        # SUB-PESTAÑA 3: GESTIÓN DE DOCENTES, GRUPOS Y CURSANTES DINÁMICOS
        with tabs_admin[2]:
            st.markdown("### 👥 Configuración de Nóminas y Tribunales")
            paralelo_gestionar = st.selectbox("Seleccione el Paralelo a configurar:", PARALELOS_DISPONIBLES)
            
            c_docentes, c_grupos = st.columns(2)
            
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
                    
            with c_grupos:
                st.markdown(f"#### 2. Grupos, Tema de Caso y Alumnos ({paralelo_gestionar})")
                with st.form("form_agregar_grupo", clear_on_submit=True):
                    nom_grupo = st.text_input("Identificador del Grupo (ej. Grupo 1):")
                    tema_grupo = st.text_input("Tema del Caso Analizado (ej. Análisis de Redes de Vínculos):")
                    alumnos_texto = st.text_area(
                        "Integrantes (escriba o pegue un nombre con grado por línea, de 4 a 6):",
                        help="Ejemplo:\nSgto. Juan Pérez\nCbos. Carlos Ruiz\nPoli. Ana Gómez",
                        placeholder="Sgto. Juan Pérez\nCbos. Carlos Ruiz\nPoli. Ana Gómez\nPoli. Luis Castro"
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
                                "INSERT INTO config_grupos (paralelo, grupo, tema_caso, alumnos_json) VALUES (?, ?, ?, ?)",
                                (paralelo_gestionar, nom_grupo.strip(), tema_grupo.strip(), json.dumps(alumnos_lista))
                            )
                            conn.commit()
                            st.success(f"Grupo registrado exitosamente con {len(alumnos_lista)} integrantes.")
                            st.rerun()
                
                cursor.execute("SELECT id, grupo, alumnos_json, tema_caso FROM config_grupos WHERE paralelo = ?", (paralelo_gestionar,))
                grps_actuales = cursor.fetchall()
                if grps_actuales:
                    st.markdown("**Grupos registrados:**")
                    for g_id, g_nom, g_alums_raw, g_tema in grps_actuales:
                        alums = json.loads(g_alums_raw)
                        with st.expander(f"📁 {g_nom} ({len(alums)} integrantes) - Tema: {g_tema}"):
                            texto_actual = "\n".join(alums)
                            nuevo_tema = st.text_input("Tema del Caso:", value=g_tema, key=f"edit_tema_{g_id}")
                            nuevo_texto_alums = st.text_area(
                                "Modificar integrantes (un nombre por línea):",
                                value=texto_actual,
                                key=f"edit_txt_{g_id}"
                            )
                            col_guardar, col_borrar = st.columns(2)
                            if col_guardar.button("💾 Guardar Cambios", key=f"btn_save_{g_id}"):
                                nueva_lista = [l.strip() for l in nuevo_texto_alums.split("\n") if l.strip()]
                                cursor.execute(
                                    "UPDATE config_grupos SET alumnos_json = ?, tema_caso = ? WHERE id = ?",
                                    (json.dumps(nueva_lista), nuevo_tema.strip(), g_id)
                                )
                                conn.commit()
                                st.success("Datos actualizados.")
                                st.rerun()
                            if col_borrar.button("🗑️ Eliminar Grupo", key=f"del_grp_{g_id}"):
                                cursor.execute("DELETE FROM config_grupos WHERE id = ?", (g_id,))
                                conn.commit()
                                st.rerun()
                else:
                    st.caption("No hay grupos agregados para este paralelo.")
                    
    elif password != "":
        st.error("Contraseña incorrecta.")
