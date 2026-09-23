import io
import json
import os
import re
import sqlite3
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ----------------- CONFIGURACIÓN DE PÁGINA -----------------
st.set_page_config(
    page_title="SURA - Censo y Tratamientos Domiciliarios",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_SURA_URL = "https://upload.wikimedia.org/wikipedia/commons/6/61/Seguros_SURA_Logo.svg"
DB_FILE = "control_pacientes.db"

# ----------------- ESTILOS CORPORATIVOS SURA -----------------
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background-color: #002270;
        color: white;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    [data-testid="stSidebar"] .stRadio label {
        background: rgba(255, 255, 255, 0.05);
        padding: 8px 12px;
        border-radius: 8px;
        margin-bottom: 6px;
        display: block;
        transition: 0.2s;
    }
    [data-testid="stSidebar"] .stRadio label:hover {
        background: rgba(0, 174, 199, 0.3);
    }
    .hospital-header {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 1.2rem 1.8rem;
        box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.05);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1.5rem;
        border-left: 6px solid #0033A0;
    }
    .hospital-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0033A0;
        margin: 0;
    }
    .hospital-sub {
        font-size: 0.9rem;
        color: #00AEC7;
        margin: 0;
        font-weight: 600;
    }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 0.9rem 0.5rem;
        box-shadow: 0px 3px 10px rgba(0,0,0,0.04);
        border: 1px solid #E6EAF0;
        text-align: center;
        margin-bottom: 0.6rem;
    }
    .metric-val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0033A0;
    }
    .metric-lbl {
        font-size: 0.76rem;
        color: #000000;
        font-weight: 700;
    }
    div.stButton > button:first-child {
        background-color: #0033A0;
        color: white;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.2rem;
        border: none;
    }
    div.stButton > button:first-child:hover {
        background-color: #00AEC7;
        color: white;
    }
    .conv-box {
        display: inline-block;
        padding: 5px 8px;
        border-radius: 6px;
        font-size: 0.78rem;
        font-weight: 600;
        margin-bottom: 4px;
        width: 100%;
        text-align: center;
        color: #000000 !important;
    }
    .footer-autor {
        text-align: center;
        padding: 1rem 0;
        margin-top: 2rem;
        border-top: 1px solid #E6EAF0;
        color: #6C757D;
        font-size: 0.85rem;
    }
    .footer-autor strong {
        color: #0033A0;
    }

    /* BARRA FIJA STICKY SUPERIOR CON DATOS DEL PACIENTE */
    .paciente-sticky-header {
        position: sticky;
        top: 0;
        z-index: 999;
        background: #002270;
        color: #FFFFFF;
        padding: 12px 20px;
        border-radius: 10px;
        box-shadow: 0px 4px 14px rgba(0,0,0,0.18);
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 3px solid #00AEC7;
    }
    .paciente-sticky-header span {
        font-size: 0.9rem;
        color: #E2E8F0;
    }
    .paciente-sticky-header strong {
        color: #FFFFFF;
    }

    /* TARJETA ELEGANTE DE TRATAMIENTO ACTIVO */
    .card-tto-activo {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-left: 6px solid #0033A0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0px 3px 10px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    .tag-turno {
        display: inline-block;
        background: #F1F5F9;
        color: #0F172A;
        border: 1px solid #CBD5E1;
        padding: 5px 12px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.88rem;
        margin-right: 8px;
    }

    /* TAGS DE MULTISELECT ESTILO SURA */
    span[data-baseweb="tag"] {
        background-color: #EBF3FC !important;
        border: 1px solid #B8D5F8 !important;
        border-radius: 6px !important;
    }
    span[data-baseweb="tag"] span {
        color: #0033A0 !important;
        font-weight: 600 !important;
    }
    span[data-baseweb="tag"] svg {
        fill: #0033A0 !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ----------------- LIMPIADOR INTELIGENTE DE TEXTO DE TURNOS -----------------
def limpiar_texto_turno(t_val):
    s = str(t_val).strip()
    if not s or s == "None" or s == "—":
        return "—"
    if s.startswith("{") and s.endswith("}"):
        try:
            d = json.loads(s)
            partes = []
            for k, v in d.items():
                partes.append(f"{k} ({v})" if str(v).upper() != "SENC" else str(k))
            return " - ".join(partes) if partes else "—"
        except Exception:
            encontrados = re.findall(r'(\d{1,2}H)', s)
            if encontrados:
                return " - ".join(encontrados)
    s = s.replace("{", "").replace("}", "").replace('"', '').replace("'", "")
    return s.strip()

# ----------------- CARGA DE CATÁLOGOS DESDE EXCEL -----------------
@st.cache_data(show_spinner=False)
def load_excel_catalog(filepath, fallback_val):
    if os.path.exists(filepath):
        try:
            df = pd.read_excel(filepath, sheet_name=0, header=None)
            if not df.empty:
                items = []
                for val in df.iloc[:, 0].dropna():
                    s = str(val).strip()
                    if s and s.lower() not in ['nan', 'none', 'medicamento', 'gestor', 'nombre']:
                        items.append(s)
                items = sorted(list(set(items)))
                if items:
                    return [""] + items
        except Exception as e:
            print(f"Error cargando {filepath}: {e}")
    return ["", fallback_val]

cat_meds_agudos = load_excel_catalog("BASE MEDICAMENTOS AGUDOS.xlsx", "MEDICAMENTO GENÉRICO")
cat_meds_be = load_excel_catalog("BASE MEDICAMENTOS BE AGUDOS.xlsx", "MEDICAMENTO BE GENÉRICO")
cat_gestores = load_excel_catalog("GESTORES DE AGUDOS.xlsx", "GESTOR GENÉRICO")

def render_selectbox(label, options, current_val, key=None, disabled=False):
    safe_options = list(options)
    if current_val and current_val not in safe_options:
        safe_options.insert(0, current_val)
    idx = safe_options.index(current_val) if current_val in safe_options else 0
    if key:
        return st.selectbox(label, safe_options, index=idx, key=key, disabled=disabled)
    return st.selectbox(label, safe_options, index=idx, disabled=disabled)

# ----------------- BASE DE DATOS Y AUTO-MIGRACIÓN -----------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS pacientes (
            documento TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            plan TEXT NOT NULL,
            programa TEXT NOT NULL,
            piso TEXT NOT NULL,
            zona TEXT NOT NULL,
            municipio TEXT DEFAULT 'Medellín',
            barrio TEXT DEFAULT '',
            aislamiento TEXT DEFAULT 'NO',
            tipo_aislamiento TEXT DEFAULT 'Ninguno',
            observaciones_clinicas TEXT DEFAULT '',
            alerta_permanente TEXT DEFAULT '',
            estado_paciente TEXT DEFAULT 'Activo',
            usuario_registro TEXT,
            fecha_ingreso TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS tratamientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT NOT NULL,
            tratamiento TEXT NOT NULL,
            dosis TEXT NOT NULL,
            via_administracion TEXT NOT NULL,
            tipo_acceso TEXT DEFAULT 'Periférico',
            frecuencia_horas INTEGER NOT NULL,
            dias INTEGER NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_fin TEXT NOT NULL,
            fecha_retiro_cateter TEXT,
            t1 TEXT,
            t2 TEXT,
            t3 TEXT,
            t1_futuro TEXT DEFAULT '',
            t2_futuro TEXT DEFAULT '',
            t3_futuro TEXT DEFAULT '',
            fecha_aplica_futuro TEXT DEFAULT '',
            puntual_fecha TEXT DEFAULT '',
            puntual_turno TEXT DEFAULT '',
            puntual_hora TEXT DEFAULT '',
            distribucion_admin TEXT DEFAULT '{}',
            visita_educacion TEXT DEFAULT '',
            dosis_base_fijas INTEGER DEFAULT 0,
            dosis_perdidas INTEGER DEFAULT 0,
            tipo_unidad_ajuste TEXT DEFAULT 'Unidosis',
            motivo_ajuste_dosis TEXT DEFAULT '',
            npt_dias_semana TEXT DEFAULT '[]',
            npt_horas_infusion INTEGER DEFAULT 12,
            npt_hora_desconexion TEXT DEFAULT '',
            be_unidosis_puente INTEGER DEFAULT 0,
            be_total_bombas INTEGER DEFAULT 0,
            be_dosis_final_cierre INTEGER DEFAULT 0,
            novedades TEXT,
            usuario_modificacion TEXT DEFAULT '',
            estado TEXT DEFAULT 'Activo',
            descanalizado TEXT DEFAULT 'NO',
            alerta_revisada TEXT DEFAULT 'NO',
            detalle_revision_alerta TEXT DEFAULT '',
            responsable_revision_alerta TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS registro_novedades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT NOT NULL,
            tipo_novedad TEXT NOT NULL,
            nivel_novedad TEXT DEFAULT 'Informativa',
            fecha_aplicacion TEXT NOT NULL,
            hora_aplicacion TEXT NOT NULL,
            detalle TEXT NOT NULL,
            responsable TEXT NOT NULL,
            estado_novedad TEXT DEFAULT 'Pendiente',
            fecha_creacion TEXT NOT NULL,
            fecha_gestion TEXT,
            responsable_gestion TEXT,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

    c.execute("PRAGMA table_info(pacientes)")
    cols_p = [col[1] for col in c.fetchall()]
    if "observaciones_clinicas" not in cols_p:
        c.execute("ALTER TABLE pacientes ADD COLUMN observaciones_clinicas TEXT DEFAULT ''")
    if "alerta_permanente" not in cols_p:
        c.execute("ALTER TABLE pacientes ADD COLUMN alerta_permanente TEXT DEFAULT ''")
    if "fecha_ingreso" not in cols_p:
        c.execute("ALTER TABLE pacientes ADD COLUMN fecha_ingreso TEXT")

    c.execute("UPDATE pacientes SET alerta_permanente = '' WHERE alerta_permanente IN ('Activo', 'Egresado', 'activo', 'egresado')")
    conn.commit()

    csv_file = "Pacientes activos agudos.csv"
    if os.path.exists(csv_file):
        try:
            df_csv = pd.read_csv(csv_file, sep=";")
            for _, row in df_csv.iterrows():
                nombres = str(row.get("Nombre", "")).strip()
                apellidos = str(row.get("Apellidos", "")).strip()
                nombre_completo = f"{nombres} {apellidos}".strip()
                doc = str(row.get("Número Documento", "")).strip()
                if not doc or doc == "nan":
                    continue
                plan = str(row.get("Plan de Salud", "POLIZA")).strip()
                programa = str(row.get("Programa", "Agudos")).strip()
                piso = str(row.get("Piso", "Piso PAS")).strip()
                municipio = str(row.get("Municipio", "MEDELLIN")).strip()
                barrio = str(row.get("Barrio", "")).strip()
                estado_p = str(row.get("Estado Paciente", "Activo")).strip()
                fec_adm = str(row.get("Fecha Admisión", datetime.now().date().isoformat())).strip()

                c.execute(
                    """
                    INSERT OR IGNORE INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso)
                    VALUES (?, ?, ?, ?, ?, 'Zona 1', ?, ?, 'NO', 'Ninguno', '', '', ?, 'Carga CSV Inicial', ?)
                    """,
                    (doc, nombre_completo, plan, programa, piso, municipio, barrio, estado_p, fec_adm[:10]),
                )
            conn.commit()
        except Exception as e:
            print("Error cargando CSV:", e)

    conn.close()

init_db()

def run_query(query, params=(), fetch=True):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(query, params)
    data = c.fetchall() if fetch else None
    conn.commit()
    conn.close()
    return data

def safe_int(valor, default=0):
    try:
        if valor is None or str(valor).strip() == "":
            return default
        return int(float(valor))
    except Exception:
        return default

# ----------------- CÁLCULOS CLÍNICOS Y BE EXACTOS -----------------
def obtener_horas_ciclo(dt_inicio, frecuencia_horas):
    frecuencia_horas = max(1, safe_int(frecuencia_horas, 8))
    dosis_diarias = 24 // frecuencia_horas
    horas_ciclo = [(dt_inicio + timedelta(hours=i * frecuencia_horas)) for i in range(dosis_diarias)]
    formato_horas = [(f"{dt.hour:02d}H" if dt.hour != 0 else "24H") for dt in horas_ciclo]
    return horas_ciclo, formato_horas

def calcular_tratamiento_mixto(dt_inicio, dias, frecuencia_horas, mapa_admin, ajuste_dosis=0):
    frecuencia_horas = max(1, safe_int(frecuencia_horas, 8))
    dias = max(1, safe_int(dias, 5))
    ajuste_dosis = safe_int(ajuste_dosis, 0)

    total_dosis_base = int((dias * 24) / frecuencia_horas)
    total_dosis = max(1, total_dosis_base + ajuste_dosis)
    fecha_fin = dt_inicio + timedelta(hours=frecuencia_horas * (total_dosis - 1))

    horas_ciclo, formato_horas = obtener_horas_ciclo(dt_inicio, frecuencia_horas)

    t1_list, t2_list, t3_list = [], [], []
    dosis_diarias_senc = 0
    dosis_diarias_cuidador = 0

    for dt, h_str in zip(horas_ciclo, formato_horas):
        responsable = mapa_admin.get(h_str, "SENC")
        if responsable == "SENC":
            dosis_diarias_senc += 1
            valor_mostrar = h_str
        else:
            dosis_diarias_cuidador += 1
            valor_mostrar = f"CUIDADOR ({h_str})"

        h = dt.hour
        if 6 <= h < 14:
            t1_list.append(valor_mostrar)
        elif 14 <= h < 22:
            t2_list.append(valor_mostrar)
        else:
            t3_list.append(valor_mostrar)

    t1_str = " - ".join(t1_list) if t1_list else "—"
    t2_str = " - ".join(t2_list) if t2_list else "—"
    t3_str = " - ".join(t3_list) if t3_list else "—"

    dosis_totales_senc = dosis_diarias_senc * dias
    dosis_totales_cuidador = dosis_diarias_cuidador * dias

    return (
        fecha_fin,
        t1_str,
        t2_str,
        t3_str,
        total_dosis_base,
        total_dosis,
        dosis_totales_senc,
        dosis_totales_cuidador,
    )

def calcular_esquema_be_exacto(dias_ordenados, frecuencia_dosis, unidosis_puente):
    dosis_por_bomba = 6 if frecuencia_dosis == 4 else (4 if frecuencia_dosis == 6 else 3)
    total_unidosis_ordenadas = int((dias_ordenados * 24) / frecuencia_dosis)

    dosis_restantes = max(0, total_unidosis_ordenadas - unidosis_puente)
    bombas_completas = dosis_restantes // dosis_por_bomba
    dosis_final_cierre = dosis_restantes % dosis_por_bomba

    return (
        total_unidosis_ordenadas,
        bombas_completas,
        dosis_final_cierre,
        dosis_por_bomba,
    )

def evaluar_alerta_fin(fecha_fin_str, frecuencia_horas=8, estado="Activo", descanalizado="NO", novedades="", alerta_revisada="NO"):
    if estado == "Completado" or descanalizado == "SI":
        return "✅ DESCANALIZADO / CERRADO"

    if "CAMBIADO" in str(estado).upper() or "SUSPENDIDO" in str(estado).upper():
        return f"🔄 {str(estado).upper()}"

    nov_up = str(novedades).upper() if novedades else ""

    if alerta_revisada != "SI" and "NOVEDAD INFORMATIVA" in nov_up:
        return "📢 NOVEDAD ACTIVA (POR GESTIONAR)"

    try:
        dt_fin = datetime.fromisoformat(fecha_fin_str)
        ahora = datetime.now()
        horas_restantes = (dt_fin - ahora).total_seconds() / 3600
        frec_val = max(1, safe_int(frecuencia_horas, 8))

        if horas_restantes <= frec_val:
            return "⚠️ FALTA 1 DOSIS (DESCANALIZAR)"
        elif horas_restantes <= 48:
            return "🔔 PRÓXIMO A FIN (<=48h - Programar)"
        return "En curso"
    except Exception:
        return "En curso"

COLOR_PICC = "#FFD1D1"
COLOR_NPT = "#FFF3CD"
COLOR_BE = "#D1E7DD"
COLOR_ANALGESIA = "#E2D9F3"
COLOR_TB = "#FFE5D0"
COLOR_PUSH = "#CFF4FC"
COLOR_AISLA = "#E8D7F1"
COLOR_FALTA1 = "#FFE0B2"

def get_color_fila_hex(row):
    via = str(row.get("Vía", "")).upper()
    acceso = str(row.get("Tipo Acceso", "")).upper()
    alerta = str(row.get("Estado Tratamiento", ""))
    aisla = str(row.get("Aislamiento", "")).upper()

    if "NOVEDAD ACTIVA" in alerta:
        return "#FFE0B2"
    if "FALTA 1 DOSIS" in alerta:
        return COLOR_FALTA1
    if acceso == "PICC":
        return COLOR_PICC
    if "NPT" in via:
        return COLOR_NPT
    if "ANALGESIA" in via:
        return COLOR_ANALGESIA
    if "BE" in via:
        return COLOR_BE
    if "TUBERCULOSIS" in via or "TB" in via:
        return COLOR_TB
    if via == "PUSH":
        return COLOR_PUSH
    if aisla == "SI":
        return COLOR_AISLA
    if "PRÓXIMO A FIN" in alerta:
        return "#FFF3CD"
    return "#FFFFFF"

# ----------------- EXPORTACIÓN DE EXCEL CON COLORES -----------------
def exportar_excel_con_colores(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilla_SURA")
        workbook = writer.book
        worksheet = writer.sheets["Planilla_SURA"]

        fill_picc = PatternFill(start_color="FFD1D1", end_color="FFD1D1", fill_type="solid")
        fill_npt = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        fill_be = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
        fill_analgesia = PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid")
        fill_tb = PatternFill(start_color="FFE5D0", end_color="FFE5D0", fill_type="solid")
        fill_push = PatternFill(start_color="CFF4FC", end_color="CFF4FC", fill_type="solid")
        fill_aisla = PatternFill(start_color="E8D7F1", end_color="E8D7F1", fill_type="solid")

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="0033A0", end_color="0033A0", fill_type="solid")
        data_font = Font(name="Calibri", size=10, bold=False, color="000000")

        for col in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill

        for row_idx, row in df.iterrows():
            via = str(row.get("Vía", "")).upper()
            acceso = str(row.get("Tipo Acceso", "")).upper()
            aisla = str(row.get("Aislamiento", "")).upper()

            fill_row = None
            if acceso == "PICC":
                fill_row = fill_picc
            elif "NPT" in via:
                fill_row = fill_npt
            elif "ANALGESIA" in via:
                fill_row = fill_analgesia
            elif "BE" in via:
                fill_row = fill_be
            elif "TUBERCULOSIS" in via or "TB" in via:
                fill_row = fill_tb
            elif via == "PUSH":
                fill_row = fill_push
            elif aisla == "SI":
                fill_row = fill_aisla

            for col_idx in range(1, len(df.columns) + 1):
                cell_d = worksheet.cell(row=row_idx + 2, column=col_idx)
                cell_d.font = data_font
                if fill_row:
                    cell_d.fill = fill_row

    return output.getvalue()

# GENERADOR NATIVO DE GRÁFICO CIRCULAR EN SVG
def generar_pie_chart_svg(datos, etiquetas, colores, titulo):
    total = sum(datos)
    if total == 0:
        return f"<div style='text-align: center; color: #6C757D; padding: 20px;'><strong>{titulo}</strong><br>Sin datos en este rango</div>"

    cx, cy, r = 100, 100, 80
    current_angle = 0
    slices_svg = ""
    legend_html = "<div style='margin-top: 10px; font-size: 0.8rem;'>"

    import math

    for val, label, color in zip(datos, etiquetas, colores):
        if val <= 0:
            continue
        pct = (val / total) * 100
        angle = (val / total) * 360

        if angle >= 360:
            slices_svg += f"<circle cx='{cx}' cy='{cy}' r='{r}' fill='{color}' stroke='#fff' stroke-width='2'/>"
        else:
            start_rad = math.radians(current_angle - 90)
            end_rad = math.radians(current_angle + angle - 90)
            x1 = cx + r * math.cos(start_rad)
            y1 = cy + r * math.sin(start_rad)
            x2 = cx + r * math.cos(end_rad)
            y2 = cy + r * math.sin(end_rad)
            large_arc = 1 if angle > 180 else 0
            d = f"M {cx} {cy} L {x1:.2f} {y1:.2f} A {r} {r} 0 {large_arc} 1 {x2:.2f} {y2:.2f} Z"
            slices_svg += f"<path d='{d}' fill='{color}' stroke='#FFFFFF' stroke-width='1.5'/>"

        current_angle += angle
        legend_html += f"<span style='display:inline-block; margin-right: 12px; margin-bottom: 4px;'><span style='display:inline-block; width:10px; height:10px; background-color:{color}; border-radius:50%; margin-right: 4px;'></span><strong>{label}:</strong> {val} ({pct:.1f}%)</span>"

    legend_html += "</div>"

    svg_code = f"""
    <div style="background: white; border-radius: 12px; padding: 16px; border: 1px solid #E6EAF0; text-align: center; box-shadow: 0px 3px 8px rgba(0,0,0,0.03);">
        <strong style="font-size: 0.95rem; color: #0033A0;">{titulo}</strong><br><br>
        <svg width="200" height="200" viewBox="0 0 200 200" style="display: block; margin: 0 auto;">
            {slices_svg}
        </svg>
        {legend_html}
    </div>
    """
    return svg_code

# ----------------- BARRA LATERAL (SIDEBAR) -----------------
with st.sidebar:
    st.markdown(
        f"""
        <div style="background-color: #FFFFFF; padding: 12px 16px; border-radius: 12px; margin-bottom: 12px; text-align: center; box-shadow: 0px 4px 10px rgba(0,0,0,0.2);">
            <img src="{LOGO_SURA_URL}" style="width: 100%; max-width: 180px; height: auto; display: block; margin: 0 auto;">
        </div>
    """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; font-size: 0.85rem; color: #00AEC7 !important; font-weight: 600;'>SALUD EN CASA (SENC) · GESTIÓN AGUDOS</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    menu = st.radio(
        "MENÚ OPERATIVO",
        [
            "📋 Gestión de Pacientes",
            "📢 Novedades y Ajustes de Ruta",
            "📊 Censo y Planilla en Vivo",
            "💾 Respaldo General en Excel",
            "📈 Analítica y Gráficos del Censo",
        ],
    )
    st.markdown("---")
    st.markdown(
        """
        <div style='background: rgba(255,255,255,0.08); padding: 12px; border-radius: 10px; text-align: center; margin-top: 40px; border: 1px solid rgba(255,255,255,0.15);'>
            <span style='font-size: 0.75rem; color: #00AEC7; text-transform: uppercase; font-weight: bold;'>Autor del Sistema</span><br>
            <strong style='font-size: 0.95rem; color: #FFFFFF;'>Andrés Medina</strong><br>
            <span style='font-size: 0.75rem; color: #E6EAF0;'>Salud en Casa · SURA Colombia</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ----------------- ENCABEZADO SUPERIOR -----------------
st.markdown(
    f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Panel de Control Clínico y Cronograma SENC</h2>
            <p class="hospital-sub">Programa de Hospitalización Domiciliaria · SURA Colombia</p>
        </div>
        <div>
            <img src="{LOGO_SURA_URL}" style="height: 48px; width: auto;">
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

if "id_tto_editando" not in st.session_state:
    st.session_state.id_tto_editando = None
if "id_tto_cambio" not in st.session_state:
    st.session_state.id_tto_cambio = None
if "id_tto_ver" not in st.session_state:
    st.session_state.id_tto_ver = None
if "id_tto_suspender" not in st.session_state:
    st.session_state.id_tto_suspender = None
if "doc_previo_gestion" not in st.session_state:
    st.session_state.doc_previo_gestion = ""

# ----------------- SECCIÓN 1: GESTIÓN DE PACIENTES Y FORMULACIÓN -----------------
if menu == "📋 Gestión de Pacientes":
    st.subheader("🔍 Localizador de Pacientes")
    doc_busqueda = st.text_input("Documento de Identidad del Paciente:", "").strip()

    if doc_busqueda != st.session_state.doc_previo_gestion:
        st.session_state.id_tto_editando = None
        st.session_state.id_tto_cambio = None
        st.session_state.id_tto_ver = None
        st.session_state.id_tto_suspender = None
        st.session_state.doc_previo_gestion = doc_busqueda

    programas_disponibles = ["Agudos"]
    municipios_disponibles = [
        "Medellín", "Bello", "Itagüí", "Envigado", "Sabaneta", "Caldas",
        "La Estrella", "Copacabana", "Girardota", "Barbosa", "Rionegro"
    ]
    pisos = ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
    zonas = ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"]
    aislamientos_tipos = ["Contacto", "Gotas", "Aerosoles", "Protector (Invertido)"]

    if doc_busqueda:
        paciente_res = run_query(
            """SELECT documento, nombre, plan, programa, piso, zona, municipio, barrio, 
                      aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, 
                      usuario_registro FROM pacientes WHERE documento = ?""",
            (doc_busqueda,),
        )

        p_doc = doc_busqueda
        p_nom = ""
        p_plan = "POS"
        p_prog = "Agudos"
        p_piso = "Norte"
        p_zona = "Zona 1"
        p_muni = "Medellín"
        p_barrio = ""
        p_aisla = "NO"
        p_tipo_aisla = "Ninguno"
        p_obs = ""
        p_alerta_fija = ""
        p_user = ""

        if paciente_res:
            (
                p_doc, p_nom, p_plan, p_prog, p_piso, p_zona, p_muni,
                p_barrio, p_aisla, p_tipo_aisla, p_obs, p_alerta_fija, p_user,
            ) = paciente_res[0]

            p_obs = str(p_obs) if p_obs else ""
            p_alerta_fija = str(p_alerta_fija) if p_alerta_fija else ""
            p_user = str(p_user) if p_user else ""

            alerta_badge_txt = f" | ⚠️ ALERTA: {p_alerta_fija.strip()}" if p_alerta_fija.strip() else ""
            st.markdown(
                f"""
                <div class="paciente-sticky-header">
                    <div>
                        <strong style="font-size: 1.05rem;">👤 {p_nom}</strong> 
                        <span style="margin-left: 8px;">(CC {p_doc})</span>
                    </div>
                    <div>
                        <span><strong>Plan:</strong> {p_plan} | <strong>Piso:</strong> {p_piso} | <strong>Zona:</strong> {p_zona} | <strong>Mpio:</strong> {p_muni}{alerta_badge_txt}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.expander("👤 Ficha Demográfica, Observaciones y Alertas Fijas", expanded=not bool(paciente_res)):
            col1, col2, col3 = st.columns(3)

            if paciente_res:
                nombre = col1.text_input("Nombre y Apellido", value=p_nom)
                plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"], index=(["POS", "Póliza", "ARL"].index(p_plan) if p_plan in ["POS", "Póliza", "ARL"] else 0))
                programa = col3.selectbox("Programa", programas_disponibles, index=0)

                col4, col5, col6, col7 = st.columns(4)
                muni_idx = municipios_disponibles.index(p_muni) if p_muni in municipios_disponibles else 0
                municipio = col4.selectbox("Municipio", municipios_disponibles, index=muni_idx)
                barrio = col5.text_input("Barrio", value=p_barrio)
                piso = col6.selectbox("Piso", pisos, index=pisos.index(p_piso) if p_piso in pisos else 0)
                zona = col7.selectbox("Zona", zonas, index=zonas.index(p_zona) if p_zona in zonas else 0)

                st.markdown("#### 📝 Observaciones Clínicas y Alerta Fija Permanente")
                col_obs1, col_obs2 = st.columns(2)
                obs_clinica = col_obs1.text_area("Observación Clínica General (En blanco si no aplica):", value=p_obs)
                alerta_fija = col_obs2.text_area("📌 Alerta Fija Permanente:", value=p_alerta_fija)

                st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
                c_ais1, c_ais2 = st.columns(2)
                tiene_aislamiento = c_ais1.radio("¿Tiene aislamiento? (Obligatorio)", ["NO", "SI"], index=0 if p_aisla == "NO" else 1, horizontal=True)
                tipo_aisla_val = "Ninguno"
                if tiene_aislamiento == "SI":
                    idx_t = aislamientos_tipos.index(p_tipo_aisla) if p_tipo_aisla in aislamientos_tipos else 0
                    tipo_aisla_val = c_ais2.selectbox("Tipo de Aislamiento:", aislamientos_tipos, index=idx_t)

                usuario_edita = render_selectbox("Tu Nombre / Profesional responsable:", cat_gestores, p_user)

                if st.button("Actualizar Ficha Paciente"):
                    if usuario_edita.strip():
                        run_query(
                            """UPDATE pacientes SET nombre=?, plan=?, programa=?, piso=?, zona=?, 
                                       municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, 
                                       observaciones_clinicas=?, alerta_permanente=?, estado_paciente='Activo', usuario_registro=? WHERE documento=?""",
                            (nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_edita.strip(), doc_busqueda),
                            fetch=False,
                        )
                        st.success("Ficha guardada y actualizada correctamente.")
                        st.rerun()
                    else:
                        st.error("Debes ingresar tu nombre como responsable.")
            else:
                st.info("Paciente nuevo. Formulario en blanco:")
                nombre = col1.text_input("Nombre y Apellido", value="")
                plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"], index=0)
                programa = col3.selectbox("Programa", programas_disponibles, index=0)

                col4, col5, col6, col7 = st.columns(4)
                municipio = col4.selectbox("Municipio", municipios_disponibles, index=0)
                barrio = col5.text_input("Barrio", value="")
                piso = col6.selectbox("Piso", pisos, index=0)
                zona = col7.selectbox("Zona", zonas, index=0)

                st.markdown("#### 📝 Observaciones Clínicas y Alerta Fija Permanente")
                col_obs1, col_obs2 = st.columns(2)
                obs_clinica = col_obs1.text_area("Observación Clínica General:", value="")
                alerta_fija = col_obs2.text_area("📌 Alerta Fija Permanente:", value="")

                st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
                c_ais1, c_ais2 = st.columns(2)
                tiene_aislamiento = c_ais1.radio("¿Tiene aislamiento?", ["NO", "SI"], index=0, horizontal=True)
                tipo_aisla_val = "Ninguno"
                if tiene_aislamiento == "SI":
                    tipo_aisla_val = c_ais2.selectbox("Tipo de Aislamiento:", aislamientos_tipos)

                usuario_crea = render_selectbox("Tu Nombre / Profesional que registra:", cat_gestores, "")

                if st.button("Admitir Paciente a la Base"):
                    if nombre and usuario_crea.strip():
                        fec_hoy = datetime.now().date().isoformat()
                        run_query(
                            """INSERT INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso) 
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?, ?)""",
                            (doc_busqueda, nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_crea.strip(), fec_hoy),
                            fetch=False,
                        )
                        st.success("Paciente admitido exitosamente.")
                        st.rerun()
                    else:
                        st.error("Diligencie el nombre del paciente y su nombre.")

        if paciente_res:
            st.divider()

            ttos_activos = run_query(
                "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo' ORDER BY id DESC",
                (doc_busqueda,),
            )

            # ----------------- ACCIONES DE TRATAMIENTO (CAMBIO / SUSPENDER / VER / VOLVER) -----------------
            if st.session_state.id_tto_cambio:
                tc = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_cambio,))[0]
                st.error(f"🔄 **CAMBIO DE TRATAMIENTO MÉDICO (SUSPENDER ANTERIOR E INICIAR NUEVO)**\n\nMedicamento: **{tc[2]}** ({tc[3]}) - Vía: {tc[4]}")
                motivo_cambio_med = st.text_input("Motivo médico del cambio / rotación (Obligatorio):", value="Rotación antibiótica / Cambio de orden médica")

                col_cbtn1, col_cbtn2 = st.columns([2, 1])
                with col_cbtn1:
                    if st.button("Confirmar Cierre y Formular Nuevo"):
                        fecha_cambio_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                        nov_cierre = f"[CAMBIO DE TTO el {fecha_cambio_str}]: Se suspende {tc[2]}. Motivo: {motivo_cambio_med}"
                        run_query("UPDATE tratamientos SET estado = 'Cambiado por Orden Médica', novedades = novedades || ' // ' || ? WHERE id = ?", (nov_cierre, tc[0]), fetch=False)
                        st.session_state.id_tto_cambio = None
                        st.success(f"✅ {tc[2]} cerrado exitosamente. Formule el nuevo medicamento abajo.")
                        st.rerun()
                with col_cbtn2:
                    if st.button("Cancelar Cambio"):
                        st.session_state.id_tto_cambio = None
                        st.rerun()
                st.divider()

            if st.session_state.id_tto_suspender:
                ts = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_suspender,))[0]
                st.error(f"⛔ **SUSPENDER TRATAMIENTO DEFINITIVAMENTE**\n\nMedicamento: **{ts[2]}** ({ts[3]}) - Vía: {ts[4]}")
                col_s1, col_s2 = st.columns(2)
                motivo_suspension = col_s1.text_input("Motivo de la suspensión médica:", value="Suspensión médica definitiva / Finalizado anticipado")
                prof_susp = render_selectbox("Profesional responsable:", cat_gestores, "", key="prof_susp")

                col_sbtn1, col_sbtn2 = st.columns([2, 1])
                with col_sbtn1:
                    if st.button("Confirmar Suspensión Definitiva"):
                        if not prof_susp.strip():
                            st.error("Ingrese profesional responsable.")
                        else:
                            fecha_susp_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            nov_susp = f"[SUSPENDIDO el {fecha_susp_str} por {prof_susp.strip()}]: {motivo_suspension}"
                            run_query("UPDATE tratamientos SET estado = 'Suspendido Médicamente', novedades = novedades || ' // ' || ?, usuario_modificacion = ? WHERE id = ?", (nov_susp, prof_susp.strip(), ts[0]), fetch=False)
                            st.session_state.id_tto_suspender = None
                            st.success(f"✅ {ts[2]} suspendido permanentemente.")
                            st.rerun()
                with col_sbtn2:
                    if st.button("Cancelar"):
                        st.session_state.id_tto_suspender = None
                        st.rerun()
                st.divider()

            if st.session_state.id_tto_ver:
                tv = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_ver,))[0]
                
                # Botón superior de retorno
                if st.button("🔙 Regresar al Listado de Tratamientos"):
                    st.session_state.id_tto_ver = None
                    st.rerun()

                st.info(f"👁️ **DETALLE COMPLETO DEL TRATAMIENTO ID #{tv[0]}**")
                st.write(f"**Medicamento:** {tv[2]} | **Dosis:** {tv[3]} | **Vía:** {tv[4]} ({tv[5]})")
                st.write(f"**Frecuencia:** Cada {tv[6]}h por {tv[7]} días")
                
                fin_formateada_ver = datetime.fromisoformat(tv[9]).strftime("%d/%m %I:%M %p") if tv[9] and tv[9] != "-" else "-"
                
                retiro_formateado_ver = "N/A"
                if tv[10] and str(tv[10]).strip() and str(tv[10]).strip() != "-":
                    try:
                        retiro_formateado_ver = datetime.fromisoformat(tv[10]).strftime("%d/%m %I:%M %p")
                    except Exception:
                        retiro_formateado_ver = str(tv[10])

                st.write(f"**Fin Tratamiento:** {fin_formateada_ver} | **Retiro Catéter:** {retiro_formateado_ver}")
                st.write(f"**Historial de Novedades:** {tv[24]}")
                
                # Botón inferior de retorno
                if st.button("🔙 Regresar al Listado de Tratamientos", key="btn_volver_abajo"):
                    st.session_state.id_tto_ver = None
                    st.rerun()
                st.divider()

            # ----------------- LISTADO DE TRATAMIENTOS ACTIVOS -----------------
            st.markdown("### 💊 Tratamientos Activos del Paciente")
            if ttos_activos:
                for t in ttos_activos:
                    alerta_t = evaluar_alerta_fin(t[9], t[6], t[20], t[21], t[18], t[22])
                    fin_dt_fmt = datetime.fromisoformat(t[9]).strftime('%d/%m %I:%M %p') if t[9] != "-" else "-"
                    
                    t1_limpio = limpiar_texto_turno(t[11])
                    t2_limpio = limpiar_texto_turno(t[12])
                    t3_limpio = limpiar_texto_turno(t[13])

                    st.markdown(
                        f"""
                        <div class="card-tto-activo">
                            <strong style="font-size: 1.05rem; color: #0033A0;">💊 {t[2]} ({t[3]})</strong> · 
                            <span><strong>Vía:</strong> {t[4]}</span> · 
                            <span><strong>Acceso:</strong> {t[5]}</span> · 
                            <span><strong>Frec:</strong> C/{t[6]}h</span> · 
                            <span><strong>Fin:</strong> {fin_dt_fmt}</span>
                            <div style="margin-top: 10px; margin-bottom: 8px;">
                                <span class="tag-turno">T1: {t1_limpio}</span>
                                <span class="tag-turno">T2: {t2_limpio}</span>
                                <span class="tag-turno">T3: {t3_limpio}</span>
                            </div>
                            <span style="font-size: 0.85rem; font-weight: 700; color: #475569;">Estado Clínico:</span> 
                            <span style="font-size: 0.85rem; font-weight: 700; color: #0033A0;">{alerta_t}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    c_act1, c_act2, c_act3, c_act4 = st.columns(4)
                    if c_act1.button("✏️ Modificar Tratamiento", key=f"btn_edit_{t[0]}"):
                        st.session_state.id_tto_editando = t[0]
                        st.session_state.id_tto_cambio = None
                        st.session_state.id_tto_ver = None
                        st.session_state.id_tto_suspender = None
                        st.rerun()

                    if c_act2.button("🔄 Cambio de Tratamiento", key=f"btn_cambio_{t[0]}"):
                        st.session_state.id_tto_cambio = t[0]
                        st.session_state.id_tto_editando = None
                        st.session_state.id_tto_ver = None
                        st.session_state.id_tto_suspender = None
                        st.rerun()

                    if c_act3.button("⛔ Suspender Tratamiento", key=f"btn_susp_{t[0]}"):
                        st.session_state.id_tto_suspender = t[0]
                        st.session_state.id_tto_editando = None
                        st.session_state.id_tto_cambio = None
                        st.session_state.id_tto_ver = None
                        st.rerun()

                    if c_act4.button("👁️ Ver Historial", key=f"btn_ver_{t[0]}"):
                        st.session_state.id_tto_ver = t[0]
                        st.rerun()

                    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            else:
                st.info("El paciente no tiene tratamientos activos en este momento.")

            tto_edit_data = None
            if st.session_state.id_tto_editando:
                tto_edit_res = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_editando,))
                if tto_edit_res:
                    tto_edit_data = tto_edit_res[0]
                    st.warning(f"✏️ **Modificando Tratamiento Activo ID #{tto_edit_data[0]}:** {tto_edit_data[2]}")

            # ----------------- FORMULARIO DE FORMULACIÓN (BLINDADO CONTRA NAMEERROR) -----------------
            st.markdown("### ➕ Formular Nuevo Tratamiento (Inicia en Dosis 1) o Guardar Cambios")

            vias = ["IV", "VO", "SC", "IM", "BE", "BE analgesia", "NPT", "Tuberculosis", "Materna", "NBZ", "PUSH"]
            accesos = ["Periférico", "PICC", "SC", "Ninguno"]

            val_via = tto_edit_data[4] if (tto_edit_data and tto_edit_data[4] in vias) else "IV"

            col_v, col_a = st.columns([1, 1])
            t_via = col_v.selectbox("Vía de Administración:", vias, index=vias.index(val_via) if val_via in vias else 0)

            val_acc = tto_edit_data[5] if (tto_edit_data and tto_edit_data[5] in accesos) else "Periférico"
            if t_via == "SC":
                con_cateter_sc = col_a.checkbox("¿La administración SC requiere catéter?", value=True if "SC" in val_acc else False)
                t_acceso = "SC" if con_cateter_sc else "Ninguno"
            elif t_via in ["VO", "IM", "Tuberculosis", "Materna", "NBZ"]:
                t_acceso = "Ninguno"
                col_a.info(f"Acceso Vascular: **No aplica para vía {t_via}**")
            else:
                t_acceso = col_a.selectbox("Acceso Vascular:", accesos, index=accesos.index(val_acc) if val_acc in accesos else 0)

            # INICIALIZACIÓN GLOBAL DE VARIABLES DE PROTOCOLO PARA EVITAR NAMEERROR
            visita_educacion_sc = ""
            npt_dias_semana_json = "[]"
            npt_horas_infusion = 12
            npt_hora_desconexion_str = ""

            # MEDICAMENTO Y DOSIS EN BLANCO (CERO PRELLENADO)
            col_med, col_dos = st.columns([2.2, 1.2])
            val_med = tto_edit_data[2] if tto_edit_data else ""
            val_dos = tto_edit_data[3] if tto_edit_data else ""

            t_nombre = ""
            bloquear_name = False

            if t_via == "BE analgesia":
                t_nombre = "BOMBA DE ANALGESIA / VIGILANCIA"
                bloquear_name = True
            elif t_via == "NPT":
                t_nombre = "NUTRICIÓN PARENTERAL TOTAL (NPT)"
                bloquear_name = True
                if not val_dos: val_dos = "1 Bolsa"
            elif t_via == "Materna":
                t_nombre = "PROGRAMA MATERNO / EDUCACIÓN"
                bloquear_name = True
                val_dos = "N/A"
            elif t_via == "Tuberculosis":
                t_nombre = "TRATAMIENTO TUBERCULOSIS (TB)"
                bloquear_name = True
                if not val_dos: val_dos = "1 Dosis Diaria"
            elif t_via == "BE":
                lista_meds = cat_meds_be
            else:
                lista_meds = cat_meds_agudos

            with col_med:
                if bloquear_name:
                    t_nombre = st.text_input("Medicamento / Terapia:", value=t_nombre, disabled=True)
                else:
                    t_nombre = render_selectbox("Medicamento / Terapia (Autocompletable):", lista_meds, val_med, key="med_select")

            with col_dos:
                t_dosis = st.text_input("Dosis (Ej: 2700mg, 1g, 500mg):", value=val_dos)

            col_f1, col_f2 = st.columns(2)
            val_frec = safe_int(tto_edit_data[6], 8) if tto_edit_data else 8
            val_dias = safe_int(tto_edit_data[7], 5) if tto_edit_data else 5

            if t_via == "BE":
                frecuencia = 6 if "c/6" in t_nombre.lower() or "6h" in t_nombre.lower() else 8
            elif t_via == "Tuberculosis":
                frecuencia = 24
            elif t_via == "BE analgesia":
                frecuencia = 12
            elif t_via == "Materna":
                frecuencia = 24
            else:
                frecuencias = [4, 6, 8, 12, 24]
                def_frec_idx = frecuencias.index(val_frec) if val_frec in frecuencias else 2
                frecuencia = col_f1.selectbox("Frecuencia (Horas):", frecuencias, index=def_frec_idx)

            dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)

            val_ini = datetime.fromisoformat(tto_edit_data[8]) if tto_edit_data else datetime.now()
            col_h1, col_h2 = st.columns(2)
            fecha_inicio = col_h1.date_input("Fecha Inicio", value=val_ini.date())
            hora_inicio = col_h2.time_input("Hora Inicio", value=val_ini.time())
            dt_inicio = datetime.combine(fecha_inicio, hora_inicio)

            # BALANCE BE Y REGLA ESTRICTA DE PUENTE NOCTURNO (23:00 - 05:00)
            unidosis_puente = 0
            bombas_calc = 0
            dosis_final_cierre = 0
            veredicto_cierre_be = ""
            cronograma_visitas = []

            if t_via == "BE":
                st.markdown(f"#### 🧪 Protocolo Clínico BE C/{frecuencia}h (1 Día = 1 Bomba Elastomérica)")
                col_be1, col_be2 = st.columns(2)
                val_puente_prev = safe_int(tto_edit_data[29], 0) if (tto_edit_data and len(tto_edit_data) > 29) else 0
                unidosis_puente = col_be1.number_input("Unidosis IV administradas como puente previo (si inició en franja nocturna 23:00 - 05:00):", min_value=0, max_value=50, value=val_puente_prev)

                total_unidosis_ord, bombas_calc, dosis_final_cierre, dosis_por_bomba = calcular_esquema_be_exacto(dias, frecuencia, unidosis_puente)
                
                if dosis_final_cierre == 0:
                    veredicto_cierre_be = "✅ El ciclo CONCLUYE CON BE COMPLETA. Retiro de catéter al día siguiente de la última bomba."
                else:
                    veredicto_cierre_be = f"⚠️ El ciclo CONCLUYE EN UNIDOSIS DE CIERRE ({dosis_final_cierre} dosis finales). Retiro de catéter el mismo día de la última unidosis de cierre."

                col_be2.markdown(
                    f"""
                    <div style='background-color: #D1E7DD; color: #000000; padding: 12px; border-radius: 8px; font-size: 0.85rem;'>
                        <strong>📦 BALANCE BE EXACTO Y VEREDICTO CLÍNICO:</strong><br>
                        • Días ordenados: <strong>{dias} Días</strong> = <strong>{bombas_calc} Bombas Completas</strong>.<br>
                        • Total unidosis equivalentes: {total_unidosis_ord} dosis.<br>
                        • Puente previo aplicado: {unidosis_puente} dosis.<br>
                        • Veredicto: <strong>{veredicto_cierre_be}</strong>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            col_aj1, col_aj2, col_aj3 = st.columns([1, 1, 2])
            tipo_unidad_ajuste = col_aj1.selectbox("Unidad de Ajuste:", ["Unidosis", "BE Completa"], index=0 if "BE" not in t_via else 1)
            val_ajuste_prev = safe_int(tto_edit_data[23], 0) if (tto_edit_data and len(tto_edit_data) > 23) else 0
            cant_ajuste = col_aj2.number_input("Cantidad (+ Aumentar / - Reducir):", min_value=-99, max_value=99, value=val_ajuste_prev)

            dosis_equiv_ajuste = cant_ajuste
            if tipo_unidad_ajuste == "BE Completa":
                dosis_por_be_adj = 4 if frecuencia == 6 else 3
                dosis_equiv_ajuste = cant_ajuste * dosis_por_be_adj

            motivo_ajuste = col_aj3.selectbox("Motivo del Ajuste / Modificación Médica:", ["Ninguno / Dosis exactas de orden médica", "Aumento de dosis ordenado por médico", "Extensión de días de tratamiento", "Paciente Ausente", "Suspensión anticipada"])

            requiere_retiro_cateter = t_via in ["IV", "BE", "BE analgesia", "PUSH"] or (t_via == "SC" and t_acceso == "SC")

            # ----------------- MOTOR DE CRONOGRAMA CLÍNICO SIN SALTO A 24H -----------------
            if t_via == "BE":
                dt_cursor = dt_inicio
                es_nocturno = (dt_inicio.hour >= 23 or dt_inicio.hour < 5) and unidosis_puente > 0
                
                if es_nocturno:
                    st.warning(f"⚠️ **ALERTA DE INICIO MIXTO:** El paciente inicia en unidosis de puente en horario nocturno (23:00 - 05:00). La Bomba Elastomérica (BE) se instalará automáticamente en la siguiente dosis correspondiente (cada {frecuencia}h).")

                for p_idx in range(1, unidosis_puente + 1):
                    cronograma_visitas.append({
                        "Secuencia": f"Puente #{p_idx}",
                        "Fecha y Hora Programada": dt_cursor.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Administración de Unidosis de Puente IV (Respetando intervalo de {frecuencia}h)"
                    })
                    dt_cursor += timedelta(hours=frecuencia)

                dt_bomba_inicio = dt_cursor if es_nocturno else dt_inicio
                
                horario_visita_be = f"{dt_bomba_inicio.strftime('%H:%M')}H (RECAMBIO BE C/24H)"
                t1_calc = horario_visita_be if 6 <= dt_bomba_inicio.hour < 14 else "—"
                t2_calc = horario_visita_be if 14 <= dt_bomba_inicio.hour < 22 else "—"
                t3_calc = horario_visita_be if (dt_bomba_inicio.hour >= 22 or dt_bomba_inicio.hour < 6) else "—"

                dt_b_iter = dt_bomba_inicio
                for b_num in range(1, bombas_calc + 1):
                    cronograma_visitas.append({
                        "Secuencia": f"Bomba BE #{b_num}",
                        "Fecha y Hora Programada": dt_b_iter.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Instalación / Recambio de Bomba Elastomérica (BE #{b_num} - C/24h)"
                    })
                    dt_b_iter += timedelta(hours=24)

                fecha_fin_calculada = dt_b_iter - timedelta(hours=24)

                if dosis_final_cierre > 0:
                    dt_cierre = dt_b_iter - timedelta(hours=24)
                    for c_num in range(1, dosis_final_cierre + 1):
                        es_ultima = (c_num == dosis_final_cierre)
                        txt_acc = f"Administración de Unidosis de Cierre #{c_num} (Intervalo 8h)"
                        if es_ultima:
                            txt_acc += " + ⚠️ RETIRO DE CATÉTER EN ESTA ÚLTIMA VISITA"
                        cronograma_visitas.append({
                            "Secuencia": f"Cierre #{c_num}",
                            "Fecha y Hora Programada": dt_cierre.strftime('%d/%m/%Y %I:%M %p'),
                            "Acción Clínica / Intervención en Domicilio": txt_acc
                        })
                        dt_cierre += timedelta(hours=8)
                    
                    fecha_fin_calculada = dt_cierre - timedelta(hours=8)
                    fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                else:
                    fecha_retiro_cateter_calc = dt_b_iter if requiere_retiro_cateter else None

            else:
                fecha_fin_calculada = dt_inicio + timedelta(days=dias)
                fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                t1_calc, t2_calc, t3_calc = "—", "—", "—"
                cronograma_visitas.append({
                    "Secuencia": "Tratamiento Estándar",
                    "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                    "Acción Clínica / Intervención en Domicilio": f"Inicio de tratamiento vía {t_via} por {dias} días."
                })

            usuario_responsable = render_selectbox("✍️ Profesional responsable que guarda/modifica (Obligatorio):", cat_gestores, "", key="prof_formular")

            st.markdown("#### ⚖️ Conciliación Final y Fechas Clínicas")
            mb1, mb2, mb3, mb4 = st.columns(4)
            mb1.metric("Días Ordenados", f"{dias} Días")
            mb2.metric("Ajuste Dosis", f"{cant_ajuste} ({tipo_unidad_ajuste})" if cant_ajuste != 0 else "0")
            mb3.metric("Fecha Fin Tratamiento", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))
            mb4.metric("Fecha Retiro Catéter", fecha_retiro_cateter_calc.strftime("%d/%m %I:%M %p") if fecha_retiro_cateter_calc else "N/A")

            # ----------------- RENDERIZADO SEGURO CON st.dataframe -----------------
            st.markdown("#### 📅 Cronograma y Visitas Día a Día (Seguimiento de Ruta)")
            df_cronograma = pd.DataFrame(cronograma_visitas)
            st.dataframe(df_cronograma, use_container_width=True, hide_index=True)

            col_btn_guardar, col_btn_cancelar = st.columns([3, 1])

            with col_btn_guardar:
                btn_label = "💾 Actualizar Cambios del Tratamiento" if tto_edit_data else "💾 Guardar Formulación de Tratamiento (Inicia en Dosis 1)"
                if st.button(btn_label):
                    if not usuario_responsable.strip():
                        st.error("❌ Indique el profesional responsable.")
                    elif not t_nombre.strip():
                        st.error("❌ Ingrese el medicamento o terapia.")
                    else:
                        fecha_reg_str = datetime.now().isoformat()
                        nov_txt = f"[{'ACTUALIZACIÓN' if tto_edit_data else 'INICIO NUEVO TTO'} {t_via}]: {t_nombre} {t_dosis} por {usuario_responsable.strip()}"
                        if t_via == "BE":
                            nov_txt += f" // {veredicto_cierre_be}"

                        if tto_edit_data:
                            run_query(
                                """
                                UPDATE tratamientos SET 
                                    tratamiento=?, dosis=?, via_administracion=?, tipo_acceso=?, frecuencia_horas=?, 
                                    dias=?, fecha_inicio=?, fecha_fin=?, fecha_retiro_cateter=?, t1=?, t2=?, t3=?, distribucion_admin=?, 
                                    visita_educacion=?, dosis_base_fijas=?, dosis_perdidas=?, tipo_unidad_ajuste=?, motivo_ajuste_dosis=?, 
                                    npt_dias_semana=?, npt_horas_infusion=?, npt_hora_desconexion=?, be_unidosis_puente=?, 
                                    be_total_bombas=?, be_dosis_final_cierre=?, novedades=novedades || ' // ' || ?, 
                                    usuario_modificacion=?, alerta_revisada='NO'
                                WHERE id=?
                                """,
                                (
                                    t_nombre.strip(), t_dosis, t_via, t_acceso, frecuencia, dias,
                                    dt_inicio.isoformat(), fecha_fin_calculada.isoformat(),
                                    fecha_retiro_cateter_calc.isoformat() if fecha_retiro_cateter_calc else "",
                                    t1_calc, t2_calc, t3_calc, json.dumps(nuevo_mapa_admin) if 'nuevo_mapa_admin' in locals() else "{}",
                                    visita_educacion_sc, dias, dosis_equiv_ajuste, tipo_unidad_ajuste, motivo_ajuste,
                                    npt_dias_semana_json, npt_horas_infusion, npt_hora_desconexion_str,
                                    unidosis_puente, bombas_calc, dosis_final_cierre,
                                    nov_txt, usuario_responsable.strip(), tto_edit_data[0],
                                ),
                                fetch=False,
                            )
                            st.session_state.id_tto_editando = None
                            st.success("✅ Tratamiento actualizado con éxito.")
                        else:
                            run_query(
                                """
                                INSERT INTO tratamientos (
                                    documento, tratamiento, dosis, via_administracion, tipo_acceso, frecuencia_horas, 
                                    dias, fecha_inicio, fecha_fin, fecha_retiro_cateter, t1, t2, t3, distribucion_admin, visita_educacion, 
                                    dosis_base_fijas, dosis_perdidas, tipo_unidad_ajuste, motivo_ajuste_dosis, 
                                    npt_dias_semana, npt_horas_infusion, npt_hora_desconexion, be_unidosis_puente, 
                                    be_total_bombas, be_dosis_final_cierre, novedades, usuario_modificacion, 
                                    estado, descanalizado, alerta_revisada, fecha_registro
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', 'NO', 'NO', ?)
                                """,
                                (
                                    doc_busqueda, t_nombre.strip(), t_dosis, t_via, t_acceso, frecuencia, dias,
                                    dt_inicio.isoformat(), fecha_fin_calculada.isoformat(),
                                    fecha_retiro_cateter_calc.isoformat() if fecha_retiro_cateter_calc else "",
                                    t1_calc, t2_calc, t3_calc, json.dumps(nuevo_mapa_admin) if 'nuevo_mapa_admin' in locals() else "{}",
                                    visita_educacion_sc, dias, dosis_equiv_ajuste, tipo_unidad_ajuste, motivo_ajuste,
                                    npt_dias_semana_json, npt_horas_infusion, npt_hora_desconexion_str,
                                    unidosis_puente, bombas_calc, dosis_final_cierre,
                                    nov_txt, usuario_responsable.strip(), fecha_reg_str,
                                ),
                                fetch=False,
                            )
                            st.success("✅ Nuevo tratamiento formulado exitosamente.")

                        st.rerun()

            with col_btn_cancelar:
                if st.button("❌ Cancelar / Limpiar"):
                    st.session_state.id_tto_editando = None
                    st.rerun()

# ----------------- SECCIÓN 2: NOVEDADES Y AJUSTES DE RUTA -----------------
elif menu == "📢 Novedades y Ajustes de Ruta":
    st.subheader("📢 Gestión Operativa de Novedades Clínicas")

    conn = sqlite3.connect(DB_FILE)
    query_novs_operativas = """
        SELECT 
            rn.id AS 'ID',
            rn.documento AS 'Cédula',
            p.nombre AS 'Nombre y Apellido',
            p.piso AS 'Piso',
            p.zona AS 'Zona',
            rn.tipo_novedad AS 'Tipo Novedad',
            rn.nivel_novedad AS 'Nivel / Tipo de Acción',
            rn.fecha_aplicacion AS 'Fecha Aplica',
            rn.hora_aplicacion AS 'Hora Aplica',
            rn.detalle AS 'Detalle Novedad',
            rn.responsable AS 'Registrado Por'
        FROM registro_novedades rn
        JOIN pacientes p ON rn.documento = p.documento
        WHERE rn.estado_novedad = 'Pendiente'
        ORDER BY rn.id DESC
    """
    df_novs = pd.read_sql_query(query_novs_operativas, conn)
    conn.close()

    ahora = datetime.now()

    def clasificar_urgencia(row):
        f_str = str(row['Fecha Aplica']).strip()
        h_str = str(row['Hora Aplica']).strip()
        try:
            if len(f_str) == 10 and "/" in f_str:
                d, m, y = f_str.split("/")
                dt_aplica = datetime(int(y), int(m), int(d))
            else:
                dt_aplica = datetime.fromisoformat(f_str[:10])

            if ":" in h_str:
                partes = h_str.split(":")
                hh = int(partes[0].strip()[-2:])
                mm = int(partes[1].strip()[:2])
                if "PM" in h_str.upper() and hh < 12:
                    hh += 12
                if "AM" in h_str.upper() and hh == 12:
                    hh = 0
                dt_aplica = dt_aplica.replace(hour=hh, minute=mm)
            else:
                dt_aplica = dt_aplica.replace(hour=8, minute=0)

            diff_horas = (dt_aplica - ahora).total_seconds() / 3600.0

            if diff_horas <= 3:
                return "🔴 CRÍTICA / EN TURNO", diff_horas
            elif 3 < diff_horas <= 12:
                return "🟡 ALERTA PRÓXIMA (VENCE EN 9-12H)", diff_horas
            else:
                return "🔵 PROGRAMADA", diff_horas
        except Exception:
            return "🔴 CRÍTICA / EN TURNO", 0

    if not df_novs.empty:
        df_novs[['Semaforo', 'Horas_Restantes']] = df_novs.apply(clasificar_urgencia, axis=1, result_type='expand')

        total_criticas = len(df_novs[df_novs['Semaforo'].str.contains("🔴")])
        total_proximas = len(df_novs[df_novs['Semaforo'].str.contains("🟡")])
        total_prog = len(df_novs[df_novs['Semaforo'].str.contains("🔵")])

        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        c_m1.metric("Pendientes Totales", f"{len(df_novs)}")
        c_m2.metric("🔴 Críticas / En Turno", f"{total_criticas}")
        c_m3.metric("🟡 Próximas (9 - 12h)", f"{total_proximas}")
        c_m4.metric("🔵 Programadas Futuras", f"{total_prog}")

        st.markdown("---")

        f1, f2, f3 = st.columns([1.5, 1.5, 2])
        pisos_disponibles = sorted(list(df_novs['Piso'].dropna().unique()))
        piso_filtro = f1.multiselect("Filtrar por tu Piso:", pisos_disponibles, default=pisos_disponibles)

        zonas_disponibles = sorted(list(df_novs['Zona'].dropna().unique()))
        zona_filtro = f2.multiselect("Filtrar por Zona:", zonas_disponibles, default=zonas_disponibles)

        semaforo_filtro = f3.multiselect(
            "Filtrar por Prioridad:",
            ["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (VENCE EN 9-12H)", "🔵 PROGRAMADA"],
            default=["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (VENCE EN 9-12H)", "🔵 PROGRAMADA"]
        )

        df_filtradas = df_novs[
            (df_novs['Piso'].isin(piso_filtro)) &
            (df_novs['Zona'].isin(zona_filtro)) &
            (df_novs['Semaforo'].isin(semaforo_filtro))
        ]

        st.markdown(f"#### 🛎️ Novedades a Gestionar en tu Servicio ({len(df_filtradas)})")

        for _, nov in df_filtradas.iterrows():
            n_id = nov['ID']
            n_doc = nov['Cédula']
            n_nom = nov['Nombre y Apellido']
            n_piso = nov['Piso']
            n_zona = nov['Zona']
            n_tipo = nov['Tipo Novedad']
            n_sem = nov['Semaforo']
            n_fec = nov['Fecha Aplica']
            n_hor = nov['Hora Aplica']
            n_det = nov['Detalle Novedad']
            n_resp = nov['Registrado Por']

            tipo_card = "urgente" if "🔴" in n_sem else ("proxima" if "🟡" in n_sem else "normal")

            with st.expander(f"{n_sem} | #{n_id} · {n_nom} (CC {n_doc}) - {n_piso} / {n_zona}", expanded=("🔴" in n_sem or "🟡" in n_sem)):
                st.markdown(
                    f"""
                    <div class="card-novedad-clinica {tipo_card}">
                        <div class="meta-linea">📋 <strong>Categoría:</strong> {n_tipo}</div>
                        <div class="meta-linea">⏰ <strong>Fecha y Hora de Aplicación:</strong> {n_fec} a las {n_hor}</div>
                        <div class="meta-linea">✍️ <strong>Registrado Por:</strong> {n_resp}</div>
                        <div style="margin-top: 10px; font-weight: 700; color: #0033A0; font-size: 0.88rem;">📝 DETALLE CLÍNICO / OPERATIVO DE LA NOVEDAD:</div>
                        <div class="detalle-caja-destacada">
                            {n_det}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                c_g1, c_g2 = st.columns([2, 2])
                prof_ges = render_selectbox(
                    "Profesional Gestor que Atiende:",
                    cat_gestores,
                    "",
                    key=f"prof_gestion_{n_id}"
                )
                acc_tomada = c_g2.text_input("Acción realizada / Nota:", value="Novedad atendida y programada en ruta.", key=f"acc_{n_id}")

                if st.button(f"✅ Resolver y Aplicar a Ruta (#{n_id})", key=f"btn_res_directo_{n_id}"):
                    if not prof_ges.strip():
                        st.error("Seleccione el profesional gestor responsable.")
                    else:
                        fec_ahora_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                        run_query(
                            """UPDATE registro_novedades SET estado_novedad = 'Gestionada', 
                                       fecha_gestion = ?, responsable_gestion = ? WHERE id = ?""",
                            (fec_ahora_str, prof_ges.strip(), n_id),
                            fetch=False,
                        )

                        restantes = run_query("SELECT COUNT(*) FROM registro_novedades WHERE documento = ? AND estado_novedad = 'Pendiente'", (n_doc,))
                        if restantes and restantes[0][0] == 0:
                            run_query(
                                """UPDATE tratamientos SET alerta_revisada = 'SI', 
                                           novedades = novedades || ' // [GESTIONADA ' || ? || ' por ' || ? || ']' WHERE documento = ? AND estado = 'Activo'""",
                                (fec_ahora_str, prof_ges.strip(), n_doc),
                                fetch=False,
                            )
                        st.success(f"✅ Novedad #{n_id} gestionada exitosamente.")
                        st.rerun()

    else:
        st.success("🎉 No hay novedades pendientes en este momento. Todas las rutas se encuentran al día.")

    st.markdown("---")

    with st.expander("➕ Registrar Nueva Novedad a un Paciente"):
        doc_manual = st.text_input("Digite la Cédula del Paciente:", "").strip()
        if doc_manual:
            p_data = run_query("SELECT documento, nombre, piso, zona FROM pacientes WHERE documento = ?", (doc_manual,))
            if p_data:
                p_doc, p_nom, p_piso, p_zona = p_data[0]
                st.info(f"Paciente: **{p_nom}** | Piso: **{p_piso}** | Zona: **{p_zona}**")

                c_reg1, c_reg2 = st.columns(2)
                tipo_nov_reg = c_reg1.selectbox("Tipo de Novedad:", ["Cambio de Horario", "Visita Fallida", "Descanalización / Salida", "Modificación Orden Médica", "Paciente Ausente", "Otra"], key="reg_tipo")
                nivel_nov_reg = c_reg2.selectbox("Nivel:", ["Nivel 1: Informativa", "Nivel 2: Visita Puntual", "Nivel 3: Cambio Permanente"], key="reg_nivel")

                c_fec1, c_fec2 = st.columns(2)
                fec_aplica_reg = c_fec1.date_input("Fecha en que Aplica:", value=datetime.now().date(), key="reg_fec")
                hor_aplica_reg = c_fec2.time_input("Hora en que Aplica:", value=datetime.now().time(), key="reg_hor")

                det_reg = st.text_area("Detalle de la Novedad (Escriba con claridad lo sucedido):", key="reg_det")
                prof_reg = render_selectbox("Profesional que Registra:", cat_gestores, "", key="reg_prof")

                if st.button("💾 Guardar Novedad en Bandeja"):
                    if not prof_reg.strip() or not det_reg.strip():
                        st.error("Diligencie todos los campos.")
                    else:
                        run_query(
                            """INSERT INTO registro_novedades (documento, tipo_novedad, nivel_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                                       VALUES (?, ?, ?, ?, ?, ?, ?, 'Pendiente', ?)""",
                            (p_doc, tipo_nov_reg, nivel_nov_reg.split(":")[0], fec_aplica_reg.strftime("%d/%m/%Y"), hor_aplica_reg.strftime("%I:%M %p"), det_reg, prof_reg.strip(), datetime.now().isoformat()),
                            fetch=False,
                        )
                        run_query("UPDATE tratamientos SET alerta_revisada = 'NO' WHERE documento = ? AND estado = 'Activo'", (p_doc,), fetch=False)
                        st.success("✅ Novedad registrada con éxito.")
                        st.rerun()
            else:
                st.error("Cédula no encontrada en la base de pacientes.")

# ----------------- SECCIÓN 3: CENSO Y PLANILLA EN VIVO -----------------
elif menu == "📊 Censo y Planilla en Vivo":
    col_vista1, col_vista2 = st.columns([3, 1])
    filtro_estado = col_vista2.selectbox("Filtrar por Estado:", ["Activos", "Completados / Descanalizados", "Todos"])

    condicion_estado = "WHERE t.estado = 'Activo'"
    if filtro_estado == "Completados / Descanalizados":
        condicion_estado = "WHERE t.estado IN ('Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"
    elif filtro_estado == "Todos":
        condicion_estado = "WHERE t.estado IN ('Activo', 'Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"

    query_general = f"""
        SELECT 
            p.documento AS 'Documento',
            p.nombre AS 'Nombre y Apellido',
            p.municipio AS 'Municipio',
            p.barrio AS 'Barrio',
            p.piso AS 'Piso',
            p.aislamiento AS 'Aislamiento',
            COALESCE(p.alerta_permanente, '') AS 'Alerta_Permanente_Texto',
            COALESCE(t.tratamiento, 'Sin Tratamiento') AS 'Tratamiento',
            COALESCE(t.dosis, '-') AS 'Dosis',
            COALESCE(t.via_administracion, '-') AS 'Vía',
            COALESCE(t.tipo_acceso, 'Periférico') AS 'Tipo Acceso',
            COALESCE(t.frecuencia_horas, 0) AS 'Frec (h)',
            COALESCE(t.dias, 0) AS 'Días',
            COALESCE(t.fecha_inicio, '-') AS 'Fecha Inicio',
            COALESCE(t.fecha_fin, '-') AS 'Fecha Fin',
            COALESCE(t.fecha_retiro_cateter, '') AS 'Fecha Retiro Cateter',
            COALESCE(t.t1, '-') AS 'T1',
            COALESCE(t.t2, '-') AS 'T2',
            COALESCE(t.t3, '-') AS 'T3',
            COALESCE(t.estado, 'Sin TTO') AS 'Estado',
            COALESCE(t.descanalizado, 'NO') AS 'Descanalizado',
            COALESCE(t.alerta_revisada, 'NO') AS 'Alerta Revisada',
            COALESCE(t.novedades, '') AS 'NOVEDAD'
        FROM pacientes p
        JOIN tratamientos t ON p.documento = t.documento
        {condicion_estado}
        ORDER BY p.piso ASC, p.nombre ASC
    """

    conn = sqlite3.connect(DB_FILE)
    df_general = pd.read_sql_query(query_general, conn)
    conn.close()

    if not df_general.empty:
        df_general["Estado Tratamiento"] = df_general.apply(
            lambda row: evaluar_alerta_fin(
                row["Fecha Fin"],
                row["Frec (h)"],
                row["Estado"],
                row["Descanalizado"],
                row["NOVEDAD"],
                row["Alerta Revisada"],
            )
            if row["Fecha Fin"] != "-"
            else "Sin tratamiento",
            axis=1,
        )

        df_general["Alerta Fija"] = df_general["Alerta_Permanente_Texto"].apply(lambda txt: "SI" if txt.strip() else "NO")

        c_f1, c_f2, c_f3, c_f4 = st.columns(4)
        filtro_piso = c_f1.multiselect("Piso", options=df_general["Piso"].dropna().unique())
        filtro_muni = c_f2.multiselect("Municipio", options=df_general["Municipio"].dropna().unique())
        filtro_via = c_f3.multiselect("Vía", options=df_general["Vía"].dropna().unique())
        filtro_alerta = c_f4.multiselect("Alertas", options=df_general["Estado Tratamiento"].dropna().unique())

        df_filtrado = df_general.copy()
        if filtro_piso:
            df_filtrado = df_filtrado[df_filtrado["Piso"].isin(filtro_piso)]
        if filtro_muni:
            df_filtrado = df_filtrado[df_filtrado["Municipio"].isin(filtro_muni)]
        if filtro_via:
            df_filtrado = df_filtrado[df_filtrado["Vía"].isin(filtro_via)]
        if filtro_alerta:
            df_filtrado = df_filtrado[df_filtrado["Estado Tratamiento"].isin(filtro_alerta)]

        total_filtrados = len(df_filtrado[df_filtrado["Estado"] == "Activo"])
        total_picc = len(df_filtrado[(df_filtrado["Tipo Acceso"].str.upper() == "PICC") & (df_filtrado["Estado"] == "Activo")])
        total_npt = len(df_filtrado[df_filtrado["Vía"].str.upper().str.contains("NPT") & (df_filtrado["Estado"] == "Activo")])
        total_be = len(df_filtrado[df_filtrado["Vía"].str.upper().str.contains("BE") & (df_filtrado["Estado"] == "Activo")])
        total_aislados = len(df_filtrado[(df_filtrado["Aislamiento"] == "SI") & (df_filtrado["Estado"] == "Activo")])
        total_nov_act = len(df_filtrado[df_filtrado["Estado Tratamiento"].str.contains("NOVEDAD ACTIVA") & (df_filtrado["Estado"] == "Activo")])
        total_descanalizar = len(df_filtrado[df_filtrado["Estado Tratamiento"].str.contains("FALTA 1 DOSIS") & (df_filtrado["Estado"] == "Activo")])

        m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
        m1.markdown(f"<div class='metric-card'><div class='metric-val'>{total_filtrados}</div><div class='metric-lbl'>ACTIVOS</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-card'><div class='metric-val'>{total_picc}</div><div class='metric-lbl'>🩸 PICC</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-card'><div class='metric-val'>{total_npt}</div><div class='metric-lbl'>🥣 NPT</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-card'><div class='metric-val'>{total_be}</div><div class='metric-lbl'>💉 BE / DOLOR</div></div>", unsafe_allow_html=True)
        m5.markdown(f"<div class='metric-card'><div class='metric-val'>{total_aislados}</div><div class='metric-lbl'>☣️ AISLADOS</div></div>", unsafe_allow_html=True)
        m6.markdown(f"<div class='metric-card'><div class='metric-val'>{total_nov_act}</div><div class='metric-lbl'>📢 NOVEDADES</div></div>", unsafe_allow_html=True)
        m7.markdown(f"<div class='metric-card'><div class='metric-val'>{total_descanalizar}</div><div class='metric-lbl'>⚠️ PEND. RETIRO</div></div>", unsafe_allow_html=True)

        st.markdown("##### 🎨 Convenciones de Color en la Planilla")
        c_c1, c_c2, c_c3, c_c4 = st.columns(4)
        with c_c1:
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_PICC};'>🩸 Catéter PICC</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_NPT};'>🥣 Nutrición Parenteral (NPT)</div>", unsafe_allow_html=True)
        with c_c2:
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_BE};'>🟩 Bomba Elastomérica (BE)</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_ANALGESIA};'>🟪 BE Analgesia / Dolor</div>", unsafe_allow_html=True)
        with c_c3:
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_TB};'>🟧 Tuberculosis (TB)</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_PUSH};'>🟦 Vía PUSH</div>", unsafe_allow_html=True)
        with c_c4:
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_AISLA};'>🟣 Paciente Aislado</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_FALTA1};'>⚠️ Falta 1 Dosis / Salida</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        html_filas = ""
        for num_idx, (_, r) in enumerate(df_filtrado.iterrows(), start=1):
            bg_color = get_color_fila_hex(r)
            alerta_fija_texto = str(r["Alerta_Permanente_Texto"]).strip()

            if alerta_fija_texto:
                celda_alerta_html = f"""
                <span class="tooltip-container">
                    <span class="badge-alerta-si">SI ⚠️</span>
                    <span class="tooltip-text">📌 <strong>ALERTA FIJA PERMANENTE:</strong><br>{alerta_fija_texto}</span>
                </span>
                """
            else:
                celda_alerta_html = '<span class="badge-alerta-no">NO</span>'

            fin_formateada = datetime.fromisoformat(r["Fecha Fin"]).strftime("%d/%m %I:%M %p") if r["Fecha Fin"] != "-" else "-"
            retiro_formateado = datetime.fromisoformat(r["Fecha Retiro Cateter"]).strftime("%d/%m %I:%M %p") if str(r["Fecha Retiro Cateter"]).strip() else "N/A"

            t1_limpio_c = limpiar_texto_turno(r['T1'])
            t2_limpio_c = limpiar_texto_turno(r['T2'])
            t3_limpio_c = limpiar_texto_turno(r['T3'])

            html_filas += f"""
            <tr style="background-color: {bg_color};">
                <td><strong>{num_idx}</strong></td>
                <td>{celda_alerta_html}</td>
                <td><strong>{r['Documento']}</strong></td>
                <td style="text-align: left;">{r['Nombre y Apellido']}</td>
                <td>{r['Municipio']}</td>
                <td>{r['Barrio']}</td>
                <td>{r['Piso']}</td>
                <td>{r['Aislamiento']}</td>
                <td><strong>{r['Tratamiento']}</strong></td>
                <td>{r['Dosis']}</td>
                <td>{r['Vía']}</td>
                <td>{r['Tipo Acceso']}</td>
                <td>C/{r['Frec (h)']}h</td>
                <td>{r['Días']}</td>
                <td>{fin_formateada}</td>
                <td style="background-color: #FFF3CD; font-weight: bold;">{retiro_formateado}</td>
                <td><code>{t1_limpio_c}</code></td>
                <td><code>{t2_limpio_c}</code></td>
                <td><code>{t3_limpio_c}</code></td>
                <td><strong>{r['Estado Tratamiento']}</strong></td>
            </tr>
            """

        tabla_html_final = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        .tabla-censo-container {{
            width: 100%;
            height: 520px;
            overflow: auto;
            border-radius: 8px;
            box-shadow: 0px 2px 8px rgba(0,0,0,0.05);
        }}
        .tabla-censo {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.82rem;
            color: #000000;
            background-color: #FFFFFF;
            font-family: sans-serif;
        }}
        .tabla-censo th {{
            background-color: #0033A0;
            color: #FFFFFF !important;
            padding: 10px 8px;
            text-align: center;
            font-weight: 700;
            white-space: nowrap;
            border: 1px solid #D9E2EC;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        .tabla-censo td {{
            padding: 8px 6px;
            text-align: center;
            border: 1px solid #E6EAF0;
            white-space: nowrap;
            color: #000000;
        }}
        .tabla-censo tr:hover {{
            filter: brightness(0.97);
        }}
        .tooltip-container {{
            position: relative;
            display: inline-block;
            cursor: pointer;
        }}
        .badge-alerta-si {{
            background-color: #FFE082;
            color: #856404;
            font-weight: 800;
            padding: 3px 8px;
            border-radius: 6px;
            border: 1px solid #FFD54F;
        }}
        .badge-alerta-no {{
            color: #6C757D;
            font-weight: 600;
        }}
        .tooltip-container .tooltip-text {{
            visibility: hidden;
            width: 280px;
            background-color: #212529;
            color: #FFFFFF;
            text-align: left;
            border-radius: 8px;
            padding: 10px 14px;
            position: absolute;
            z-index: 999999;
            top: 50%;
            left: 100%;
            transform: translateY(-50%);
            margin-left: 12px;
            opacity: 0;
            transition: opacity 0.2s ease-in-out;
            box-shadow: 0px 6px 18px rgba(0,0,0,0.3);
            font-size: 0.82rem;
            line-height: 1.4;
            border-left: 5px solid #FFC107;
            white-space: normal;
        }}
        .tooltip-container .tooltip-text::after {{
            content: "";
            position: absolute;
            top: 50%;
            right: 100%;
            margin-top: -6px;
            border-width: 6px;
            border-style: solid;
            border-color: transparent #212529 transparent transparent;
        }}
        .tooltip-container:hover .tooltip-text {{
            visibility: visible;
            opacity: 1;
        }}
        </style>
        </head>
        <body>
            <div class="tabla-censo-container">
                <table class="tabla-censo">
                    <thead>
                        <tr>
                            <th>N°</th>
                            <th>Alerta Fija</th>
                            <th>Documento</th>
                            <th>Nombre y Apellido</th>
                            <th>Municipio</th>
                            <th>Barrio</th>
                            <th>Piso</th>
                            <th>Aisla</th>
                            <th>Tratamiento</th>
                            <th>Dosis</th>
                            <th>Vía</th>
                            <th>Acceso</th>
                            <th>Frec</th>
                            <th>Días</th>
                            <th>Fin TTO</th>
                            <th style="background-color: #D39E00; color: #FFFFFF !important;">Retiro Catéter</th>
                            <th>Turno T1</th>
                            <th>Turno T2</th>
                            <th>Turno T3</th>
                            <th>Estado TTO</th>
                        </tr>
                    </thead>
                    <tbody>
                        {html_filas}
                    </tbody>
                </table>
            </div>
        </body>
        </html>
        """

        components.html(tabla_html_final, height=520, scrolling=True)
        st.markdown("<br>", unsafe_allow_html=True)

        df_descarga = df_filtrado.copy()
        df_descarga["Alerta Fija"] = df_descarga["Alerta_Permanente_Texto"].apply(lambda t: "SI" if str(t).strip() else "NO")
        cols_excel = [
            "Alerta Fija", "Documento", "Nombre y Apellido", "Municipio", "Barrio", "Piso",
            "Aislamiento", "Alerta_Permanente_Texto", "Tratamiento", "Dosis", "Vía", "Tipo Acceso",
            "Frec (h)", "Días", "Fecha Fin", "Fecha Retiro Cateter", "T1", "T2", "T3", "Estado Tratamiento", "NOVEDAD"
        ]
        df_descarga = df_descarga[cols_excel]
        df_descarga.rename(columns={"Alerta_Permanente_Texto": "Detalle Alerta Fija"}, inplace=True)

        excel_coloreado = exportar_excel_con_colores(df_descarga)
        st.download_button(
            label="📥 Descargar Planilla en Excel con Colores (.xlsx)",
            data=excel_coloreado,
            file_name=f"Planilla_SENC_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No hay registros activos.")

# ----------------- SECCIÓN 4: RESPALDO GENERAL EN EXCEL -----------------
elif menu == "💾 Respaldo General en Excel":
    st.subheader("💾 Copia de Seguridad Consolidada en Excel (Con Colores)")
    opcion_backup = st.radio(
        "Modalidad de Respaldo:",
        ["1. Planilla Consolidada Limpia (Solo activos)", "2. Planilla Completa con Historial (Auditoría)"],
        index=0,
    )

    if st.button("📊 Generar y Descargar Archivo Excel"):
        conn = sqlite3.connect(DB_FILE)
        if "Solo activos" in opcion_backup:
            query_b = """
                SELECT 
                    p.documento AS 'Documento', p.nombre AS 'Nombre y Apellido', p.plan AS 'Plan', 
                    p.programa AS 'Programa', p.municipio AS 'Municipio', p.barrio AS 'Barrio', 
                    p.piso AS 'Piso', p.zona AS 'Zona', p.alerta_permanente AS 'Alerta Fija', 
                    p.observaciones_clinicas AS 'Obs Clínica', p.aislamiento AS 'Aislamiento', 
                    t.tratamiento AS 'Medicamento', t.dosis AS 'Dosis', t.via_administracion AS 'Vía', 
                    t.tipo_acceso AS 'Tipo Acceso', t.frecuencia_horas AS 'Frec (h)', t.dias AS 'Días', 
                    t.fecha_inicio AS 'Fecha Inicio', t.fecha_fin AS 'Fecha Fin', t.fecha_retiro_cateter AS 'Fecha Retiro Cateter',
                    t.t1 AS 'T1', t.t2 AS 'T2', t.t3 AS 'T3', t.novedades AS 'NOVEDAD',
                    t.usuario_modificacion AS 'Gestor Responsable'
                FROM pacientes p
                JOIN tratamientos t ON p.documento = t.documento
                WHERE t.estado = 'Activo'
                ORDER BY p.piso ASC, p.nombre ASC
            """
            nombre_archivo = f"CENSO_ACTIVO_LIMPIO_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        else:
            query_b = """
                SELECT 
                    p.documento AS 'Documento', p.nombre AS 'Nombre y Apellido', p.plan AS 'Plan', 
                    p.programa AS 'Programa', p.municipio AS 'Municipio', p.barrio AS 'Barrio', 
                    p.piso AS 'Piso', p.zona AS 'Zona', p.alerta_permanente AS 'Alerta Fija', 
                    p.observaciones_clinicas AS 'Obs Clínica', p.estado_paciente AS 'Estado Paciente',
                    t.tratamiento AS 'Medicamento', t.dosis AS 'Dosis', t.via_administracion AS 'Vía', 
                    t.tipo_acceso AS 'Tipo Acceso', t.dias AS 'Días', t.fecha_inicio AS 'Fecha Inicio', 
                    t.fecha_fin AS 'Fecha Fin', t.fecha_retiro_cateter AS 'Fecha Retiro Cateter',
                    t.t1 AS 'T1', t.t2 AS 'T2', t.t3 AS 'T3', 
                    t.novedades AS 'NOVEDAD', t.estado AS 'Estado TTO', t.descanalizado AS 'Descanalizado',
                    t.usuario_modificacion AS 'Gestor Responsable', t.fecha_registro AS 'Fecha Movimiento'
                FROM pacientes p
                LEFT JOIN tratamientos t ON p.documento = t.documento
                ORDER BY p.piso ASC, p.documento ASC, t.id DESC
            """
            nombre_archivo = f"HISTORIAL_AUDITORIA_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        df_b = pd.read_sql_query(query_b, conn)
        conn.close()

        excel_b_coloreado = exportar_excel_con_colores(df_b)
        st.download_button(
            label="⬇️ Descargar Archivo Excel (.xlsx)",
            data=excel_b_coloreado,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----------------- SECCIÓN 5: ANALÍTICA Y GRÁFICOS INTERACTIVOS (SVG NATIVO) -----------------
elif menu == "📈 Analítica y Gráficos del Censo":
    st.subheader("📈 Tablero Analítico y Gráficos Circulares SENC")
    conn = sqlite3.connect(DB_FILE)
    df_pacientes_all = pd.read_sql_query("SELECT * FROM pacientes", conn)
    df_ttos_all = pd.read_sql_query("SELECT * FROM tratamientos", conn)
    conn.close()

    if not df_pacientes_all.empty:
        col_fec1, col_fec2 = st.columns(2)
        f_desde = col_fec1.date_input("Fecha Ingreso Desde:", value=(datetime.now() - timedelta(days=30)).date())
        f_hasta = col_fec2.date_input("Fecha Ingreso Hasta:", value=datetime.now().date())

        df_pacientes_all["fecha_ingreso_dt"] = pd.to_datetime(df_pacientes_all["fecha_ingreso"], errors="coerce")
        df_p_rango = df_pacientes_all[(df_pacientes_all["fecha_ingreso_dt"].dt.date >= f_desde) & (df_pacientes_all["fecha_ingreso_dt"].dt.date <= f_hasta)]

        total_ingresos_rango = len(df_p_rango)
        activos_rango = len(df_p_rango[df_p_rango["estado_paciente"] == "Activo"])
        egresados_rango = len(df_p_rango[df_p_rango["estado_paciente"] == "Egresado"])

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Ingresos en Periodo", f"{total_ingresos_rango} Pacientes")
        col_m2.metric("Pacientes Activos", f"{activos_rango} Activos")
        col_m3.metric("Pacientes Egresados / Alta", f"{egresados_rango} Egresados")
        pct_alta = round((egresados_rango / total_ingresos_rango) * 100, 1) if total_ingresos_rango > 0 else 0
        col_m4.metric("% Tasa de Egreso / Alta", f"{pct_alta}%")

        st.markdown("---")
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            data_estado = [activos_rango, egresados_rango]
            labels_estado = ["Activos", "Egresados"]
            colores_estado = ["#0033A0", "#6C757D"]
            st.markdown(generar_pie_chart_svg(data_estado, labels_estado, colores_estado, "Distribución por Estado de Pacientes"), unsafe_allow_html=True)

        with col_g2:
            df_ttos_active = df_ttos_all[df_ttos_all["estado"] == "Activo"]
            if not df_ttos_active.empty:
                vias_cuenta = df_ttos_active["via_administracion"].value_counts()
                palette_vias = ["#0033A0", "#00AEC7", "#28A745", "#FFC107", "#DC3545", "#6F42C1", "#FD7E14", "#20C997"]
                st.markdown(generar_pie_chart_svg(list(vias_cuenta.values), list(vias_cuenta.index), palette_vias[: len(vias_cuenta)], "Distribución por Vía / Terapia"), unsafe_allow_html=True)
            else:
                st.info("No hay tratamientos activos.")

        st.markdown("<br>", unsafe_allow_html=True)
        col_g3, col_g4 = st.columns(2)

        with col_g3:
            if not df_ttos_active.empty:
                acc_cuenta = df_ttos_active["tipo_acceso"].value_counts()
                palette_acc = ["#17A2B8", "#E83E8C", "#6C757D", "#FFC107"]
                st.markdown(generar_pie_chart_svg(list(acc_cuenta.values), list(acc_cuenta.index), palette_acc[: len(acc_cuenta)], "Distribución por Acceso Vascular"), unsafe_allow_html=True)

        with col_g4:
            pisos_cuenta = df_pacientes_all[df_pacientes_all["estado_paciente"] == "Activo"]["piso"].value_counts()
            if not pisos_cuenta.empty:
                palette_pisos = ["#0033A0", "#00AEC7", "#20C997", "#28A745", "#FFC107"]
                st.markdown(generar_pie_chart_svg(list(pisos_cuenta.values), list(pisos_cuenta.index), palette_pisos[: len(pisos_cuenta)], "Distribución por Piso Hospitalario"), unsafe_allow_html=True)
    else:
        st.info("No hay pacientes registrados.")

# ----------------- PIE DE PÁGINA INSTITUCIONAL -----------------
st.markdown(
    """
    <div class="footer-autor">
        🏥 <strong>Sistema de Gestión y Control de Tratamientos Domiciliarios (SENC)</strong><br>
        Diseñado y desarrollado por <strong>Andrés Medina</strong> · Salud en Casa SURA Colombia
    </div>
""",
    unsafe_allow_html=True,
)
