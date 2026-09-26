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
LOGO_SURA_FAVICON = "https://upload.wikimedia.org/wikipedia/commons/6/61/Seguros_SURA_Logo.svg"

st.set_page_config(
    page_title="SURA - Base Agudos",
    page_icon=LOGO_SURA_FAVICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_SURA_URL = "https://upload.wikimedia.org/wikipedia/commons/6/61/Seguros_SURA_Logo.svg"
DB_FILE = "control_pacientes.db"

# ----------------- ESTILOS CORPORATIVOS DASHBOARD CLÍNICO DE ALTA GAMA -----------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F4F7FB;
        color: #0F172A;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #002270 0%, #00123D 100%);
        color: white;
        padding-top: 1.5rem;
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    [data-testid="stSidebar"] .stRadio label {
        background: rgba(255, 255, 255, 0.03);
        padding: 12px 16px;
        border-radius: 10px;
        margin-bottom: 8px;
        display: block;
        transition: all 0.25s ease-in-out;
        border: 1px solid rgba(255, 255, 255, 0.06);
        font-weight: 500;
        font-size: 0.92rem;
    }
    [data-testid="stSidebar"] .stRadio label:hover {
        background: rgba(0, 174, 199, 0.25);
        border-color: rgba(0, 174, 199, 0.4);
        transform: translateX(4px);
    }

    .hospital-header {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 1.2rem 2rem;
        box-shadow: 0px 4px 16px rgba(0, 51, 160, 0.04);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1.5rem;
        border-left: 5px solid #0033A0;
        border: 1px solid #E2E8F0;
    }
    .hospital-title {
        font-size: 1.45rem;
        font-weight: 700;
        color: #0033A0;
        margin: 0;
        letter-spacing: -0.3px;
    }
    .hospital-sub {
        font-size: 0.85rem;
        color: #00AEC7;
        margin: 0;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.6px;
    }

    .metric-card {
        background: #FFFFFF;
        border-radius: 10px;
        padding: 1rem 0.5rem;
        box-shadow: 0px 3px 10px rgba(0,0,0,0.02);
        border: 1px solid #E2E8F0;
        text-align: center;
        margin-bottom: 0.6rem;
        transition: transform 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0px 5px 14px rgba(0, 51, 160, 0.06);
        border-color: #00AEC7;
    }
    .metric-val {
        font-size: 1.6rem;
        font-weight: 800;
        color: #0033A0;
    }
    .metric-lbl {
        font-size: 0.74rem;
        color: #64748B;
        font-weight: 700;
        text-transform: uppercase;
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #0033A0 0%, #002270 100%);
        color: white;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.2rem;
        border: none;
        box-shadow: 0px 2px 6px rgba(0, 51, 160, 0.15);
        transition: all 0.2s ease-in-out;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(135deg, #00AEC7 0%, #008B9E 100%);
        color: white;
        box-shadow: 0px 4px 10px rgba(0, 174, 199, 0.25);
        transform: translateY(-1px);
    }

    .conv-box {
        display: inline-block;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: 0.78rem;
        font-weight: 700;
        margin-bottom: 4px;
        width: 100%;
        text-align: center;
        color: #0F172A !important;
        box-shadow: 0px 1px 3px rgba(0,0,0,0.02);
    }

    .footer-autor {
        text-align: center;
        padding: 1.4rem 0;
        margin-top: 3rem;
        border-top: 1px solid #E2E8F0;
        color: #64748B;
        font-size: 0.85rem;
        background: #FFFFFF;
        border-radius: 10px;
        box-shadow: 0px -2px 10px rgba(0,0,0,0.02);
    }
    .footer-autor strong {
        color: #0033A0;
    }

    .paciente-sticky-header {
        position: sticky;
        top: 0;
        z-index: 999;
        background: linear-gradient(135deg, #002270 0%, #0033A0 100%);
        color: #FFFFFF;
        padding: 12px 20px;
        border-radius: 10px;
        box-shadow: 0px 4px 15px rgba(0, 34, 112, 0.2);
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 3px solid #00AEC7;
    }
    .paciente-sticky-header span {
        font-size: 0.88rem;
        color: #E2E8F0;
    }
    .paciente-sticky-header strong {
        color: #FFFFFF;
    }

    .card-tto-activo {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #0033A0;
        border-radius: 10px;
        padding: 15px 20px;
        box-shadow: 0px 3px 10px rgba(0,0,0,0.02);
        margin-bottom: 15px;
        transition: all 0.2s ease-in-out;
    }
    .card-tto-activo:hover {
        box-shadow: 0px 5px 15px rgba(0, 51, 160, 0.06);
        border-color: #CBD5E1;
    }
    .tag-turno {
        display: inline-block;
        background: #F8FAFC;
        color: #0F172A;
        border: 1px solid #CBD5E1;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.85rem;
        margin-right: 8px;
    }

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

# ----------------- CARGA DE CATÁLOGOS DESDE EXCEL (SIN CACHÉ) -----------------
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
            tipo_documento TEXT DEFAULT 'CC',
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

    c.execute("""
        CREATE TABLE IF NOT EXISTS entrega_turnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT NOT NULL,
            zona TEXT NOT NULL,
            prioridad TEXT DEFAULT 'Media',
            detalle_novedad TEXT NOT NULL,
            estado TEXT DEFAULT 'Activa',
            registrado_por TEXT NOT NULL,
            fecha_creacion TEXT NOT NULL,
            fecha_resolucion TEXT,
            resuelto_por TEXT,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

    # TABLA DE AYUDAS DIAGNÓSTICAS V6 (INCLUYE ORIGEN Y TIPO DE DOCUMENTO)
    c.execute("""
        CREATE TABLE IF NOT EXISTS ayudas_diagnosticas_v6 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT NOT NULL,
            tipo_documento TEXT DEFAULT 'CC',
            piso TEXT NOT NULL,
            zona TEXT NOT NULL,
            ayuda_solicitada TEXT NOT NULL,
            peso TEXT DEFAULT '',
            requiere_oxigeno TEXT DEFAULT 'NO',
            cantidad_oxigeno TEXT DEFAULT '',
            tipo_transporte TEXT DEFAULT 'No requiere transporte',
            origen_solicitud TEXT DEFAULT 'Línea de Ayudas Diagnósticas',
            prestador_pgp TEXT DEFAULT '',
            observaciones TEXT DEFAULT '',
            estado TEXT DEFAULT 'Pendiente',
            registrado_por TEXT NOT NULL,
            fecha_registro TEXT NOT NULL,
            fecha_cita TEXT DEFAULT '',
            hora_cita TEXT DEFAULT '',
            lugar_cita TEXT DEFAULT '',
            paciente_informado TEXT DEFAULT 'NO',
            nombre_acompanante TEXT DEFAULT '',
            contacto_acompanante TEXT DEFAULT '',
            fecha_gestion TEXT,
            resuelto_por TEXT DEFAULT ''
        )
    """)

    c.execute("PRAGMA table_info(pacientes)")
    cols_p = [col[1] for col in c.fetchall()]
    if "tipo_documento" not in cols_p:
        c.execute("ALTER TABLE pacientes ADD COLUMN tipo_documento TEXT DEFAULT 'CC'")
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
                    INSERT OR REPLACE INTO pacientes (documento, tipo_documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso)
                    VALUES (?, 'CC', ?, ?, ?, ?, 'Zona 1', ?, ?, COALESCE((SELECT aislamiento FROM pacientes WHERE documento=?), 'NO'), COALESCE((SELECT tipo_aislamiento FROM pacientes WHERE documento=?), 'Ninguno'), COALESCE((SELECT observaciones_clinicas FROM pacientes WHERE documento=?), ''), COALESCE((SELECT alerta_permanente FROM pacientes WHERE documento=?), ''), ?, 'Carga CSV Inicial', ?)
                    """,
                    (doc, nombre_completo, plan, programa, piso, municipio, barrio, doc, doc, doc, doc, estado_p, fec_adm[:10]),
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

def evaluar_alerta_fin(fecha_fin_str, frecuencia_horas=8, estado="Activo", descanalizado="NO", novedades="", alerta_revisada="NO", via_administracion="IV"):
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

        es_invasivo = str(via_administracion).upper() in ["IV", "BE", "BE ANALGESIA", "PUSH"]

        if horas_restantes <= 0:
            dias_vencido = abs(horas_restantes) / 24
            if dias_vencido >= 3 and not es_invasivo:
                return "⚠️ SEGUIMIENTO >3 DÍAS (PENDIENTE EGRESO)"
            return "⚠️ SIN TTO ACTIVO (PENDIENTE EGRESO)"
        elif horas_restantes <= frec_val:
            if es_invasivo:
                return "⚠️ FALTA 1 DOSIS (DESCANALIZAR)"
            else:
                return "🔔 FIN DE CICLO (Evaluar Egreso)"
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
    if "SIN TTO ACTIVO" in alerta or "FALTA 1 DOSIS" in alerta or "SEGUIMIENTO" in alerta:
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
    if via in ["PUSH", "IV"]:
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
        worksheet = workbook.active

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
            elif via in ["PUSH", "IV"]:
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
        <div style="background-color: #FFFFFF; padding: 16px 20px; border-radius: 12px; margin-bottom: 18px; text-align: center; box-shadow: 0px 4px 14px rgba(0,0,0,0.12);">
            <img src="{LOGO_SURA_URL}" style="width: 100%; max-width: 150px; height: auto; display: block; margin: 0 auto;">
        </div>
    """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; font-size: 0.75rem; color: #00AEC7 !important; font-weight: 700; letter-spacing: 1px; margin-bottom: 20px;'>SALUD EN CASA (SENC) · BASE AGUDOS</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    
    menu = st.radio(
        "MENÚ OPERATIVO",
        [
            "Gestor Pacientes",
            "Entrega de turno",
            "Novedades programacion",
            "Ayudas diagnosticas",
            "Censo y Planilla en Vivo",
            "Analitica y graficos",
            "Informes",
        ],
    )
    st.markdown("---")

# ----------------- ENCABEZADO SUPERIOR INSTITUCIONAL -----------------
st.markdown(
    f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Base y Control Agudos</h2>
            <p class="hospital-sub">Programa Salud en Casa · SURA</p>
        </div>
        <div>
            <img src="{LOGO_SURA_URL}" style="height: 42px; width: auto;">
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

# ----------------- SECCIÓN 1: GESTOR PACIENTES (CON VALIDACIÓN DE TIPO DE DOCUMENTO) -----------------
if menu == "Gestor Pacientes":
    st.subheader("🔍 Localizador de Pacientes")
    
    col_busq_t, col_busq1, col_busq2, col_busq3 = st.columns([1, 2.2, 1, 1])
    
    tipos_doc_lista = ["CC", "TI", "CE", "PPT", "Pasaporte", "RC"]
    tipo_doc_busq = col_busq_t.selectbox("Tipo Doc", tipos_doc_lista, index=0, key="tipo_doc_busq_sel")

    if "input_doc_val" not in st.session_state:
        st.session_state.input_doc_val = ""

    doc_busqueda = col_busq1.text_input("Documento de Identidad:", value=st.session_state.input_doc_val).strip()

    with col_busq2:
        st.markdown("<div style='margin-top: 27px;'></div>", unsafe_allow_html=True)
        if st.button("🧹 Limpiar Cédula", use_container_width=True):
            st.session_state.input_doc_val = ""
            st.session_state.doc_activo_actual = ""
            st.session_state.id_tto_editando = None
            st.session_state.id_tto_cambio = None
            st.session_state.id_tto_ver = None
            st.session_state.id_tto_suspender = None
            for k_s in list(st.session_state.keys()):
                if "med_select" in k_s or "prof_formular" in k_s or "widget_hora_inicio" in k_s or "chk_cuid_" in k_s or "reg_tipo" in k_s or "reg_det" in k_s:
                    del st.session_state[k_s]
            st.rerun()

    with col_busq3:
        st.markdown("<div style='margin-top: 27px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Nueva Consulta", use_container_width=True):
            st.session_state.input_doc_val = ""
            st.session_state.doc_activo_actual = ""
            st.session_state.id_tto_editando = None
            st.session_state.id_tto_cambio = None
            st.session_state.id_tto_ver = None
            st.session_state.id_tto_suspender = None
            for k_s in list(st.session_state.keys()):
                if "med_select" in k_s or "prof_formular" in k_s or "widget_hora_inicio" in k_s or "chk_cuid_" in k_s or "reg_tipo" in k_s or "reg_det" in k_s:
                    del st.session_state[k_s]
            st.rerun()

    if doc_busqueda != st.session_state.get("doc_activo_actual", ""):
        st.session_state.id_tto_editando = None
        st.session_state.id_tto_cambio = None
        st.session_state.id_tto_ver = None
        st.session_state.id_tto_suspender = None
        st.session_state.doc_activo_actual = doc_busqueda
        st.session_state.limpiar_form_nuevo = True
        for k_s in list(st.session_state.keys()):
            if "med_select" in k_s or "prof_formular" in k_s or "widget_hora_inicio" in k_s or "chk_cuid_" in k_s:
                del st.session_state[k_s]

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
            """SELECT documento, tipo_documento, nombre, plan, programa, piso, zona, municipio, barrio, 
                      aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, 
                      usuario_registro, fecha_ingreso FROM pacientes WHERE documento = ?""",
            (doc_busqueda,),
        )

        p_doc = doc_busqueda
        p_t_doc = tipo_doc_busq
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
        p_fec_ingreso = None

        if paciente_res:
            (
                p_doc, p_t_doc_db, p_nom, p_plan, p_prog, p_piso, p_zona, p_muni,
                p_barrio, p_aisla, p_tipo_aisla, p_obs, p_alerta_fija, p_user, p_fec_ingreso,
            ) = paciente_res[0]

            # VALIDACIÓN ESTRICTA DE TIPO DE DOCUMENTO
            if p_t_doc_db and p_t_doc_db.upper() != tipo_doc_busq.upper():
                st.error(f"🚨 **Error de tipo de documento:** El paciente con documento `{doc_busqueda}` está registrado en el sistema bajo el tipo **{p_t_doc_db}**, no **{tipo_doc_busq}**.")
                st.stop()

            p_obs = str(p_obs) if p_obs else ""
            p_alerta_fija = str(p_alerta_fija) if p_alerta_fija else ""
            p_user = str(p_user) if p_user else ""

            dias_estancia_calc = 0
            if p_fec_ingreso:
                try:
                    dt_ing = datetime.fromisoformat(str(p_fec_ingreso)[:10])
                    dias_estancia_calc = (datetime.now() - dt_ing).days
                    if dias_estancia_calc < 0: dias_estancia_calc = 0
                except Exception:
                    pass

            alerta_badge_txt = f" | ⚠️ ALERTA: {p_alerta_fija.strip()}" if p_alerta_fija.strip() else ""
            st.markdown(
                f"""
                <div class="paciente-sticky-header">
                    <div>
                        <strong style="font-size: 1.05rem;">👤 {p_nom}</strong> 
                        <span style="margin-left: 8px;">({p_t_doc_db} {p_doc})</span>
                    </div>
                    <div>
                        <span><strong>Plan:</strong> {p_plan} | <strong>Mpio:</strong> {p_muni} ({p_barrio}) | <strong>Estancia:</strong> {dias_estancia_calc} Días{alerta_badge_txt}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.expander("👤 Ficha Demográfica, Observaciones y Alertas Fijas", expanded=not bool(paciente_res)):
            col1, col2, col3 = st.columns(3)

            if paciente_res:
                t_doc_sel = col1.selectbox("Tipo Documento", tipos_doc_lista, index=tipos_doc_lista.index(p_t_doc_db) if p_t_doc_db in tipos_doc_lista else 0)
                nombre = col2.text_input("Nombre y Apellido", value=p_nom)
                plan = col3.selectbox("Plan", ["POS", "Póliza", "ARL"], index=(["POS", "Póliza", "ARL"].index(p_plan) if p_plan in ["POS", "Póliza", "ARL"] else 0))
                
                col_p1, col_p2 = st.columns(2)
                programa = col_p1.selectbox("Programa", programas_disponibles, index=0)

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
                            """UPDATE pacientes SET tipo_documento=?, nombre=?, plan=?, programa=?, piso=?, zona=?, 
                                       municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, 
                                       observaciones_clinicas=?, alerta_permanente=?, estado_paciente='Activo', usuario_registro=? WHERE documento=?""",
                            (t_doc_sel, nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_edita.strip(), doc_busqueda),
                            fetch=False,
                        )
                        st.success("Ficha guardada y actualizada correctamente.")
                        st.rerun()
                    else:
                        st.error("Debes ingresar tu nombre como responsable.")
            else:
                st.info("Paciente nuevo. Formulario en blanco:")
                t_doc_sel = col1.selectbox("Tipo Documento", tipos_doc_lista, index=0)
                nombre = col2.text_input("Nombre y Apellido", value="")
                plan = col3.selectbox("Plan", ["POS", "Póliza", "ARL"], index=0)

                programa = col1.selectbox("Programa", programas_disponibles, index=0)

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
                            """INSERT INTO pacientes (documento, tipo_documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso) 
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?, ?)""",
                            (doc_busqueda, t_doc_sel, nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_crea.strip(), fec_hoy),
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

            # ----------------- ACCIONES DE TRATAMIENTO -----------------
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
                
                if st.button("🔙 Regresar al Listado de Tratamientos", key="btn_volver_abajo"):
                    st.session_state.id_tto_ver = None
                    st.rerun()
                st.divider()

            # ----------------- LISTADO DE TRATAMIENTOS ACTIVOS -----------------
            st.markdown("### 💊 Tratamientos Activos del Paciente")
            if ttos_activos:
                for t in ttos_activos:
                    alerta_t = evaluar_alerta_fin(t[9], t[6], t[20], t[21], t[18], t[22], t[4])
                    fin_dt_fmt = datetime.fromisoformat(t[9]).strftime('%d/%m %I:%M %p') if t[9] != "-" else "-"
                    
                    t1_hoy, t2_hoy, t3_hoy = "—", "—", "—"
                    try:
                        frec_t = safe_int(t[6], 8)
                        dt_ini_t = datetime.fromisoformat(t[8])
                        puente_t = safe_int(t[29], 0)
                        hoy_date = datetime.now().date()
                        
                        cursor_t = dt_ini_t
                        eventos_hoy = []
                        if str(t[4]).upper() == "BE":
                            for _ in range(puente_t):
                                if cursor_t.date() == hoy_date:
                                    eventos_hoy.append(cursor_t)
                                cursor_t += timedelta(hours=frec_t)
                            for _ in range(safe_int(t[30], 1) + 10):
                                if cursor_t.date() == hoy_date:
                                    eventos_hoy.append(cursor_t)
                                cursor_t += timedelta(hours=24)
                        elif str(t[4]).upper() == "BE ANALGESIA":
                            for _ in range(6):
                                if cursor_t.date() == hoy_date:
                                    eventos_hoy.append(cursor_t)
                                cursor_t += timedelta(hours=12)
                        elif str(t[4]).upper() in ["TUBERCULOSIS", "MATERNA"]:
                            dias_tot_m = safe_int(t[7], 5)
                            curr_m = dt_ini_t
                            for _ in range(dias_tot_m):
                                if curr_m.date() == hoy_date:
                                    eventos_hoy.append(curr_m)
                                curr_m += timedelta(days=1)
                        elif str(t[4]).upper() == "SC":
                            visita_edu = str(t[22]).strip().upper() if len(t) > 22 else ""
                            if visita_edu == "SOLO EDUCACIÓN":
                                if dt_ini_t.date() == hoy_date:
                                    eventos_hoy.append(dt_ini_t)
                            else:
                                _, formato_c = obtener_horas_ciclo(dt_ini_t, frec_t)
                                curr_sc = dt_ini_t
                                total_d_sc = safe_int(t[7], 5)
                                for _ in range(total_d_sc):
                                    for _ in formato_c:
                                        if curr_sc.date() == hoy_date:
                                            eventos_hoy.append(curr_sc)
                                        curr_sc += timedelta(hours=frec_t)
                        elif str(t[4]).upper() == "NBZ":
                            visita_edu = str(t[22]).strip().upper() if len(t) > 22 else ""
                            if visita_edu == "SOLO EDUCACIÓN":
                                if dt_ini_t.date() == hoy_date:
                                    eventos_hoy.append(dt_ini_t)
                            else:
                                _, formato_c = obtener_horas_ciclo(dt_ini_t, frec_t)
                                curr_nbz = dt_ini_t
                                total_d_nbz = safe_int(t[7], 5)
                                for _ in range(total_d_nbz):
                                    for _ in formato_c:
                                        if curr_nbz.date() == hoy_date:
                                            eventos_hoy.append(curr_nbz)
                                        curr_nbz += timedelta(hours=frec_t)
                        elif str(t[4]).upper() == "NPT":
                            dias_activos_npt = json.loads(t[26]) if t[26] else ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                            nombres_dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                            dia_hoy_str = nombres_dias[hoy_date.weekday()]
                            
                            if dia_hoy_str in dias_activos_npt:
                                eventos_hoy.append(datetime.combine(hoy_date, dt_ini_t.time()))
                            ayer_date = hoy_date - timedelta(days=1)
                            dia_ayer_str = nombres_dias[ayer_date.weekday()]
                            if dia_ayer_str in dias_activos_npt:
                                h_inf = safe_int(t[27], 12)
                                dt_desconex_ayer = datetime.combine(ayer_date, dt_ini_t.time()) + timedelta(hours=h_inf)
                                if dt_desconex_ayer.date() == hoy_date:
                                    eventos_hoy.append(dt_desconex_ayer)
                        else:
                            _, formato_c = obtener_horas_ciclo(dt_inicio_f:=dt_ini_t, frec_t)
                            curr_gen = dt_ini_t
                            for _ in range(safe_int(t[7], 5)):
                                for _ in formato_c:
                                    if curr_gen.date() == hoy_date:
                                        eventos_hoy.append(curr_gen)
                                    curr_gen += timedelta(hours=frec_t)

                        for ev in eventos_hoy:
                            h_ev = ev.hour
                            hora_txt = f"{ev.strftime('%H:%M')}H"
                            if 6 <= h_ev < 14:
                                t1_hoy = hora_txt
                            elif 14 <= h_ev < 22:
                                t2_hoy = hora_txt
                            else:
                                t3_hoy = hora_txt
                                
                        if not eventos_hoy:
                            t1_hoy = limpiar_texto_turno(t[11])
                            t2_hoy = limpiar_texto_turno(t[12])
                            t3_hoy = limpiar_texto_turno(t[13])
                    except Exception:
                        t1_hoy = limpiar_texto_turno(t[11])
                        t2_hoy = limpiar_texto_turno(t[12])
                        t3_hoy = limpiar_texto_turno(t[13])

                    st.markdown(
                        f"""
                        <div class="card-tto-activo">
                            <strong style="font-size: 1.05rem; color: #0033A0;">💊 {t[2]} ({t[3]})</strong> · 
                            <span><strong>Vía:</strong> {t[4]}</span> · 
                            <span><strong>Acceso:</strong> {t[5]}</span> · 
                            <span><strong>Frec:</strong> C/{t[6]}h</span> · 
                            <span><strong>Fin:</strong> {fin_dt_fmt}</span>
                            <div style="margin-top: 10px; margin-bottom: 8px;">
                                <span class="tag-turno">T1: {t1_hoy}</span>
                                <span class="tag-turno">T2: {t2_hoy}</span>
                                <span class="tag-turno">T3: {t3_hoy}</span>
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

            if st.session_state.get('limpiar_form_nuevo', False):
                tto_edit_data = None
                st.session_state.limpiar_form_nuevo = False

            # ----------------- FORMULARIO DE FORMULACIÓN (IV Y PUSH UNIFICADOS) -----------------
            st.markdown("### ➕ Formular Nuevo Tratamiento (Inicia en Dosis 1) o Guardar Cambios")

            vias = ["IV", "VO", "SC", "IM", "BE", "BE analgesia", "NPT", "Tuberculosis", "Materna", "NBZ", "PUSH", "LEV"]
            accesos = ["Periférico", "PICC", "SC", "Ninguno"]

            val_via = tto_edit_data[4] if tto_edit_data else "IV"

            col_v, col_a = st.columns([1, 1])
            t_via = col_v.selectbox("Vía de Administración:", vias, index=vias.index(val_via) if val_via in vias else 0)

            val_acc = tto_edit_data[5] if tto_edit_data else "Periférico"
            if t_via == "NPT":
                t_acceso = "PICC"
                col_a.info("Acceso Vascular: **Fijado en PICC por protocolo NPT**")
            elif t_via == "SC":
                con_cateter_sc = col_a.checkbox("¿La administración SC requiere catéter?", value=True if "SC" in val_acc else False)
                t_acceso = "SC" if con_cateter_sc else "Ninguno"
            elif t_via in ["VO", "IM", "Tuberculosis", "Materna", "NBZ"]:
                t_acceso = "Ninguno"
                col_a.info(f"Acceso Vascular: **No aplica para vía {t_via}**")
            else:
                val_acc_idx = accesos.index(val_acc) if val_acc in accesos else 0
                t_acceso = col_a.selectbox("Acceso Vascular:", accesos, index=val_acc_idx)

            visita_educacion_sc = ""
            npt_dias_semana_json = "[]"
            npt_horas_infusion = 12
            npt_hora_desconexion_str = ""

            col_med, col_dos = st.columns([2.2, 1.2])
            val_med = tto_edit_data[2] if tto_edit_data else ""
            val_dos = tto_edit_data[3] if tto_edit_data else ""

            t_nombre = ""
            bloquear_name = False

            if t_via == "BE analgesia":
                t_nombre = "BOMBA DE ANALGESIA / VIGILANCIA"
                bloquear_name = True
                val_dos = "1 Bomba"
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
                val_dos = "1 Dosis Diaria"
            elif t_via == "LEV":
                t_nombre = "LÍQUIDOS ENDOVENOSOS (LEV)"
                bloquear_name = True
                if not val_dos: val_dos = "1 Bolsa / 1000cc"
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
                if t_via == "BE":
                    match_d = re.search(r'([\d\.,]+\s*(?:mg|gr|g))', t_nombre, re.IGNORECASE)
                    t_dosis = match_d.group(0) if match_d else (val_dos if val_dos else "Automática")
                    st.text_input("Dosis (Extraída Automáticamente):", value=t_dosis, disabled=True)
                elif t_via in ["BE analgesia", "Tuberculosis", "Materna", "NPT", "LEV"]:
                    t_dosis = st.text_input("Dosis / Volumen:", value=val_dos, disabled=True if t_via in ["NPT", "LEV"] else False)
                else:
                    t_dosis = st.text_input("Dosis (Ej: 1g, 500mg, 1 Bolsa, 50 MG):", value=val_dos)

            col_f1, col_f2 = st.columns(2)
            val_frec = safe_int(tto_edit_data[6], 8) if tto_edit_data else 8
            val_dias = safe_int(tto_edit_data[7], 5) if tto_edit_data else 5

            if t_via == "BE":
                frecuencia = col_f1.selectbox("Frecuencia de Unidosis de Puente (Horas):", [6, 8], index=0 if val_frec == 6 else 1)
                dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)
            elif t_via == "BE analgesia":
                frecuencia = 12
                dias = 3
                col_f1.info("Frecuencia: **Cada 12 horas (Fijo)**")
                col_f2.info("Días Ordenados: **3 Días (Fijo)**")
            elif t_via in ["Tuberculosis", "Materna"]:
                frecuencia = 24
                dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)
                col_f1.info(f"Frecuencia: **Cada 24 horas (Diaria para {t_via})**")
            elif t_via == "NPT":
                frecuencia = 24
                dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)
                col_f1.info("Frecuencia NPT: **Protocolo diario/semanal**")
            elif t_via == "LEV":
                frecuencia = col_f1.selectbox("Duración de Infusión LEV (Horas):", [12, 24], index=1 if val_frec == 24 else 0)
                dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)
            else:
                frecuencias = [4, 6, 8, 12, 24]
                def_frec_idx = frecuencias.index(val_frec) if val_frec in frecuencias else 2
                frecuencia = col_f1.selectbox("Frecuencia (Horas):", frecuencias, index=def_frec_idx)
                dias = col_f2.number_input("Días Ordenados por Médico:", min_value=1, max_value=999, value=val_dias)

            tipo_visita_sc_nbz = "Visitar por Horario"
            if t_via in ["SC", "NBZ"]:
                st.markdown(f"#### 🩺 Modalidad de Intervención ({t_via})")
                val_edu_prev = tto_edit_data[22] if (tto_edit_data and len(tto_edit_data) > 22) else "Visitar por Horario"
                idx_edu = 1 if val_edu_prev == "Solo Educación" else 0
                tipo_visita_sc_nbz = st.radio(
                    f"Seleccione la modalidad para {t_via}:",
                    ["Visitar por Horario", "Solo Educación"],
                    index=idx_edu,
                    horizontal=True
                )
                visita_educacion_sc = tipo_visita_sc_nbz

            mapa_admin = {}
            if t_via in ["IV", "PUSH", "SC", "NBZ", "LEV"] and tipo_visita_sc_nbz != "Solo Educación":
                st.markdown("#### 👥 Distribución de Administración (SENC vs Cuidador)")
                st.info("Por defecto todas las dosis son administradas por SENC. Si alguna dosis la administra el cuidador en casa, márcala abajo:")
                
                dt_temp_ini = datetime.combine(datetime.now().date(), datetime.now().time())
                _, formato_horas_temp = obtener_horas_ciclo(dt_temp_ini, frecuencia if t_via != "LEV" else 24)
                
                try:
                    mapa_admin = json.loads(tto_edit_data[21]) if (tto_edit_data and tto_edit_data[21]) else {}
                except Exception:
                    mapa_admin = {}

                cols_c = st.columns(min(len(formato_horas_temp), 4))
                for idx_h, h_str in enumerate(formato_horas_temp):
                    col_actual = cols_c[idx_h % len(cols_c)]
                    actual_resp = mapa_admin.get(h_str, "SENC")
                    chk_cuidador = col_actual.checkbox(f"Dosis {h_str}: Cuidador", value=(actual_resp == "Cuidador"), key=f"chk_cuid_{h_str}")
                    mapa_admin[h_str] = "Cuidador" if chk_cuidador else "SENC"

            if t_via == "NPT":
                st.markdown("#### 🥣 Protocolo NPT (Días de Infusión y Horarios)")
                col_npt1, col_npt2 = st.columns(2)
                
                dias_def_npt = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                if tto_edit_data and len(tto_edit_data) > 26 and tto_edit_data[26]:
                    try:
                        dias_def_npt = json.loads(tto_edit_data[26])
                    except Exception:
                        pass
                
                dias_semana_npt = col_npt1.multiselect("Días programados para infusión NPT en la semana:", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], default=dias_def_npt)
                npt_dias_semana_json = json.dumps(dias_semana_npt)
                
                horas_inf_opts = [12, 14, 16, 18, 20, 24]
                h_inf_def = safe_int(tto_edit_data[27], 12) if (tto_edit_data and len(tto_edit_data) > 27) else 12
                npt_horas_infusion = col_npt2.selectbox("Duración de infusión (Horas):", horas_inf_opts, index=horas_inf_opts.index(h_inf_def) if h_inf_def in horas_inf_opts else 0)

            val_ini = datetime.fromisoformat(tto_edit_data[8]) if tto_edit_data else datetime.now()
            
            if "hora_inicio_persistente" not in st.session_state:
                st.session_state.hora_inicio_persistente = val_ini.time()

            col_h1, col_h2 = st.columns(2)
            fecha_inicio = col_h1.date_input("Fecha Inicio", value=val_ini.date())
            hora_inicio = col_h2.time_input("Hora Inicio", value=st.session_state.hora_inicio_persistente, key="widget_hora_inicio")
            st.session_state.hora_inicio_persistente = hora_inicio

            dt_inicio = datetime.combine(fecha_inicio, hora_inicio)

            if t_via == "NPT":
                dt_desconex = dt_inicio + timedelta(hours=npt_horas_infusion)
                npt_hora_desconexion_str = dt_desconex.strftime("%H:%M")
                st.info(f"🥣 **Horario NPT Calculado:** Instalación a las {dt_inicio.strftime('%I:%M %p')} -> Retiro / Desconexión a las {dt_desconex.strftime('%I:%M %p')}")

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

            requiere_retiro_cateter = t_via in ["IV", "BE", "BE analgesia", "PUSH", "LEV"] or (t_via == "SC" and t_acceso == "SC")

            # ----------------- MOTOR DE CRONOGRAMA CLÍNICO INTELIGENTE -----------------
            if t_via in ["IV", "PUSH"]:
                (
                    fecha_fin_calculada,
                    t1_calc,
                    t2_calc,
                    t3_calc,
                    _, _, _, _
                ) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
                
                fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                
                es_medicamento_2h = any(m in t_nombre.upper() for m in ["VANCOMICINA", "HIERRO", "POLIMIXINA"])
                
                _, formato_c = obtener_horas_ciclo(dt_inicio, frecuencia)
                curr_dt_c = dt_inicio
                for d_idx in range(dias):
                    for h_str in formato_c:
                        resp = mapa_admin.get(h_str, "SENC")
                        cronograma_visitas.append({
                            "Secuencia": f"Dosis {h_str} (Día {d_idx+1})",
                            "Fecha y Hora Programada": curr_dt_c.strftime('%d/%m/%Y %I:%M %p'),
                            "Acción Clínica / Intervención en Domicilio": f"Administración vía {t_via} a cargo de: {resp}"
                        })
                        if es_medicamento_2h:
                            dt_retiro_2h = curr_dt_c + timedelta(hours=2)
                            cronograma_visitas.append({
                                "Secuencia": f"Retiro / Control 2h ({h_str})",
                                "Fecha y Hora Programada": dt_retiro_2h.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": "⚠️ Retiro de infusión / Flush a las 2 horas post-instalación"
                            })
                        curr_dt_c += timedelta(hours=frecuencia)

            elif t_via == "SC":
                if tipo_visita_sc_nbz == "Solo Educación":
                    fecha_fin_calculada = dt_inicio
                    fecha_retiro_cateter_calc = None
                    t1_calc, t2_calc, t3_calc = f"EDUCACIÓN ({dt_inicio.strftime('%H:%M')})", "—", "—"
                    cronograma_visitas.append({
                        "Secuencia": "Visita Única de Educación SC",
                        "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": "Visita única domiciliaria para instrucción y educación en vía Subcutánea"
                    })
                else:
                    (
                        fecha_fin_calculada,
                        t1_calc,
                        t2_calc,
                        t3_calc,
                        _, _, _, _
                    ) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
                    
                    fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                    _, formato_c = obtener_horas_ciclo(dt_inicio, frecuencia)
                    curr_dt_c = dt_inicio
                    for d_idx in range(dias):
                        for h_str in formato_c:
                            resp = mapa_admin.get(h_str, "SENC")
                            cronograma_visitas.append({
                                "Secuencia": f"Dosis {h_str} (Día {d_idx+1})",
                                "Fecha y Hora Programada": curr_dt_c.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": f"Administración vía SC a cargo de: {resp}"
                            })
                            curr_dt_c += timedelta(hours=frecuencia)

            elif t_via == "BE":
                dt_cursor = dt_inicio
                es_nocturno = (dt_inicio.hour >= 23 or dt_inicio.hour < 5) and unidosis_puente > 0
                
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

            elif t_via == "BE analgesia":
                dt_cursor_ba = dt_inicio
                total_visitas_ba = 6
                for v_idx in range(1, total_visitas_ba + 1):
                    es_ultima_ba = (v_idx == total_visitas_ba)
                    txt_ba = f"Visita de revisión y control BE Analgesia #{v_idx} (C/12h)"
                    if es_ultima_ba:
                        txt_ba += " + ⚠️ RETIRO DE BE ANALGESIA Y CATÉTER + EGRESO"
                    cronograma_visitas.append({
                        "Secuencia": f"Control BE Analgesia #{v_idx}",
                        "Fecha y Hora Programada": dt_cursor_ba.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": txt_ba
                    })
                    dt_cursor_ba += timedelta(hours=12)
                
                fecha_fin_calculada = dt_cursor_ba - timedelta(hours=12)
                fecha_retiro_cateter_calc = dt_cursor_ba - timedelta(hours=12)
                t1_calc, t2_calc, t3_calc = f"CONTROL ({dt_inicio.strftime('%H:%M')})", "REVISIÓN C/12H", "—"

            elif t_via in ["Tuberculosis", "Materna"]:
                fecha_fin_calculada = dt_inicio + timedelta(days=dias - 1)
                fecha_retiro_cateter_calc = None
                t1_calc, t2_calc, t3_calc = f"VISITA DIARIA ({dt_inicio.strftime('%H:%M')})", "—", "—"
                
                dt_m_curr = dt_inicio
                for m_d in range(1, dias + 1):
                    cronograma_visitas.append({
                        "Secuencia": f"Visita {t_via} Día {m_d}",
                        "Fecha y Hora Programada": dt_m_curr.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Visita de seguimiento y control programa {t_via} (Día {m_d} de {dias})"
                    })
                    dt_m_curr += timedelta(hours=24)

            elif t_via == "NBZ":
                if tipo_visita_sc_nbz == "Solo Educación":
                    fecha_fin_calculada = dt_inicio
                    fecha_retiro_cateter_calc = None
                    t1_calc, t2_calc, t3_calc = f"EDUCACIÓN ({dt_inicio.strftime('%H:%M')})", "—", "—"
                    cronograma_visitas.append({
                        "Secuencia": "Visita Única de Educación NBZ",
                        "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": "Visita única domiciliaria para instrucción y educación en técnica de Nebulización"
                    })
                else:
                    (
                        fecha_fin_calculada,
                        t1_calc,
                        t2_calc,
                        t3_calc,
                        _, _, _, _
                    ) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
                    
                    fecha_retiro_cateter_calc = None
                    _, formato_c = obtener_horas_ciclo(dt_inicio, frecuencia)
                    curr_nbz = dt_inicio
                    for d_idx in range(dias):
                        for h_str in formato_c:
                            resp = mapa_admin.get(h_str, "SENC")
                            cronograma_visitas.append({
                                "Secuencia": f"Nebulización {h_str} (Día {d_idx+1})",
                                "Fecha y Hora Programada": curr_nbz.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": f"Administración de Nebulización - C/{frecuencia}h a cargo de: {resp}"
                            })
                            curr_nbz += timedelta(hours=frecuencia)

            elif t_via == "NPT":
                fecha_fin_calculada = dt_inicio + timedelta(days=dias)
                fecha_retiro_cateter_calc = None
                t1_calc, t2_calc, t3_calc = f"INSTALAR ({dt_inicio.strftime('%H:%M')})", f"RETIRAR ({npt_hora_desconexion_str})", "—"
                
                dias_activos_npt = json.loads(tto_edit_data[26] if tto_edit_data and len(tto_edit_data)>26 and tto_edit_data[26] else npt_dias_semana_json)
                nombres_dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                
                curr_npt = dt_inicio
                for d_i in range(dias):
                    dia_str = nombres_dias[curr_npt.weekday()]
                    if dia_str in dias_activos_npt:
                        dt_retir = curr_npt + timedelta(hours=npt_horas_infusion)
                        cronograma_visitas.append({
                            "Secuencia": f"NPT Día {d_i+1} ({dia_str})",
                            "Fecha y Hora Programada": curr_npt.strftime('%d/%m/%Y %I:%M %p'),
                            "Acción Clínica / Intervención en Domicilio": f"Instalación NPT y retiro programado a las {dt_retir.strftime('%I:%M %p')}"
                        })
                    else:
                        ayer_dt = curr_npt - timedelta(days=1)
                        dia_ayer_str = nombres_dias[ayer_dt.weekday()]
                        if dia_ayer_str in dias_activos_npt:
                            dt_retir_ayer = curr_npt.replace(hour=int(npt_hora_desconexion_str.split(":")[0]), minute=int(npt_hora_desconexion_str.split(":")[1]))
                            cronograma_visitas.append({
                                "Secuencia": f"NPT Retiro ({dia_str})",
                                "Fecha y Hora Programada": dt_retir_ayer.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": "⚠️ Retiro / Desconexión de NPT (Día sin instalación nueva)"
                            })
                    curr_npt += timedelta(days=1)

            elif t_via == "LEV":
                fecha_fin_calculada = dt_inicio + timedelta(days=dias)
                fecha_retiro_cateter_calc = fecha_fin_calculada
                t1_calc, t2_calc, t3_calc = f"INFUSIÓN LEV ({dt_inicio.strftime('%H:%M')})", "—", "—"
                
                curr_lev = dt_inicio
                for d_lv in range(dias):
                    dt_ret_lev = curr_lev + timedelta(hours=frecuencia)
                    cronograma_visitas.append({
                        "Secuencia": f"LEV Día {d_lv+1}",
                        "Fecha y Hora Programada": curr_lev.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Instalación de Líquidos Endovenosos (Duración {frecuencia}h) - Retiro / Cambio a las {dt_ret_lev.strftime('%I:%M %p')}"
                    })
                    curr_lev += timedelta(hours=24)
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
                                    t1_calc, t2_calc, t3_calc, json.dumps(mapa_admin),
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
                                    t1_calc, t2_calc, t3_calc, json.dumps(mapa_admin),
                                    visita_educacion_sc, dias, dosis_equiv_ajuste, tipo_unidad_ajuste, motivo_ajuste,
                                    npt_dias_semana_json, npt_horas_infusion, npt_hora_desconexion_str,
                                    unidosis_puente, bombas_calc, dosis_final_cierre,
                                    nov_txt, usuario_responsable.strip(), fecha_reg_str,
                                ),
                                fetch=False,
                            )
                            st.success("✅ Nuevo tratamiento formulado exitosamente.")
                            st.session_state.limpiar_form_nuevo = True

                        st.rerun()

            with col_btn_cancelar:
                if st.button("❌ Cancelar / Limpiar"):
                    st.session_state.id_tto_editando = None
                    st.rerun()

# ----------------- SECCIÓN 2: ENTREGA DE TURNO -----------------
elif menu == "Entrega de turno":
    st.subheader("📋 Módulo de Entrega de Turnos Operativos")
    st.info("💡 Gestiona y consulta las novedades de turno filtradas por zona, con clasificación de prioridad y autocompletado desde el Censo.")

    col_z_f1, col_z_f2 = st.columns(2)
    zona_filtro_turno = col_z_f1.selectbox("Filtrar por Zona de Turno:", ["Todas", "Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"])
    ver_historial_turnos = col_z_f2.checkbox("📂 Ver Historial de Turnos Resueltos (Auditoría)")

    with st.expander("➕ Registrar Novedad o Minuta para el Turno", expanded=False):
        ced_turno_in = st.text_input("Cédula del Paciente (Autocompleta del Censo):").strip()
        
        p_nombre_t = ""
        p_piso_t = "Norte"
        p_zona_t = "Zona 1"
        
        if ced_turno_in:
            res_pt = run_query("SELECT nombre, piso, zona FROM pacientes WHERE documento = ?", (ced_turno_in,))
            if res_pt:
                p_nombre_t, p_piso_t, p_zona_t = res_pt[0]
                st.success(f"👤 Paciente: **{p_nombre_t}** | Piso: {p_piso_t} | Zona sugerida: {p_zona_t}")
            else:
                st.warning("⚠️ Cédula no registrada en el censo. Se registrará la novedad con el documento ingresado.")
        
        c_t_col1, c_t_col2, c_t_col3 = st.columns(3)
        zona_asignada = c_t_col1.selectbox("Zona Asignada:", ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"], index=0)
        prioridad_turno = c_t_col2.selectbox("Prioridad:", ["Alta 🔴", "Media 🟡", "Baja 🔵"], index=1)
        gestor_turno = render_selectbox("Gestor que reporta:", cat_gestores, "", key="gestor_turno_sel")

        detalle_turno_txt = st.text_area("Detalle de la novedad de turno a tener en cuenta:")

        if st.button("💾 Guardar Novedad en Minuta de Turno"):
            if not ced_turno_in or not detalle_turno_txt.strip() or not gestor_turno.strip():
                st.error("Ingrese la cédula, el detalle de la novedad y seleccione el gestor.")
            else:
                fec_reg_t = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                run_query(
                    """INSERT INTO entrega_turnos (documento, zona, prioridad, detalle_novedad, estado, registrado_por, fecha_creacion)
                       VALUES (?, ?, ?, ?, 'Activa', ?, ?)""",
                    (ced_turno_in, zona_asignada, prioridad_turno, detalle_turno_txt.strip(), gestor_turno.strip(), fec_reg_t),
                    fetch=False
                )
                st.success("✅ Novedad de turno guardada con éxito.")
                st.rerun()

    st.markdown("---")
    
    if ver_historial_turnos:
        st.markdown("#### 📂 Historial de Novedades de Turno (Resueltas)")
        q_hist = "SELECT * FROM entrega_turnos WHERE estado = 'Resuelta' ORDER BY id DESC"
        df_hist = pd.read_sql_query(q_hist, sqlite3.connect(DB_FILE))
        if not df_hist.empty:
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
        else:
            st.info("No hay historial de turnos resueltos.")
    else:
        st.markdown("#### 📌 Novedades Activas de Turno en Curso")
        q_act = "SELECT * FROM entrega_turnos WHERE estado = 'Activa'"
        if zona_filtro_turno != "Todas":
            q_act += f" AND zona = '{zona_filtro_turno}'"
        q_act += " ORDER BY id DESC"
        
        df_act = pd.read_sql_query(q_act, sqlite3.connect(DB_FILE))
        
        if not df_act.empty:
            for _, r_t in df_act.iterrows():
                t_id = r_t['id']
                t_doc = r_t['documento']
                t_zona = r_t['zona']
                t_prio = r_t['prioridad']
                t_det = r_t['detalle_novedad']
                t_reg = r_t['registrado_por']
                t_fec = r_t['fecha_creacion']
                
                p_nom_busq = t_doc
                res_nom = run_query("SELECT nombre FROM pacientes WHERE documento = ?", (t_doc,))
                if res_nom:
                    p_nom_busq = f"{res_nom[0][0]} (CC {t_doc})"

                with st.expander(f"{t_prio} | [{t_zona}] · {p_nom_busq} - {t_fec}", expanded=True):
                    st.write(f"**Detalle:** {t_det}")
                    st.write(f"*Reportó: {t_reg}*")
                    
                    c_res1, c_res2 = st.columns([2, 1])
                    nombre_gestor_res = render_selectbox("Gestor que resuelve:", cat_gestores, "", key=f"res_gestor_{t_id}")
                    if c_res2.button(f"✅ Marcar como Resuelta (#{t_id})", key=f"btn_res_t_{t_id}"):
                        if not nombre_gestor_res.strip():
                            st.error("Seleccione el gestor que resuelve.")
                        else:
                            fec_res_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            run_query(
                                "UPDATE entrega_turnos SET estado = 'Resuelta', fecha_resolucion = ?, resuelto_por = ? WHERE id = ?",
                                (fec_res_str, nombre_gestor_res.strip(), t_id),
                                fetch=False
                            )
                            st.success("✅ Novedad marcada como resuelta y archivada del turno actual.")
                            st.rerun()
        else:
            st.info(f"No hay novedades activas de turno para la selección actual ({zona_filtro_turno}).")

# ----------------- SECCIÓN 3: NOVEDADES PROGRAMACION -----------------
elif menu == "Novedades programacion":
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

            if diff_horas < 7:
                return "🔴 CRÍTICA / EN TURNO", diff_horas
            elif 7 <= diff_horas <= 15:
                return "🟡 ALERTA PRÓXIMA (VENCE EN 8-15H)", diff_horas
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
        c_m2.metric("🔴 Críticas / En Turno (<7h)", f"{total_criticas}")
        c_m3.metric("🟡 Próximas (8 - 15h)", f"{total_proximas}")
        c_m4.metric("🔵 Programadas (>15h)", f"{total_prog}")

        st.markdown("---")

        f1, f2, f3 = st.columns([1.5, 1.5, 2])
        pisos_disponibles = sorted(list(df_novs['Piso'].dropna().unique()))
        piso_filtro = f1.multiselect("Filtrar por tu Piso:", pisos_disponibles, default=pisos_disponibles)

        zonas_disponibles = sorted(list(df_novs['Zona'].dropna().unique()))
        zona_filtro = f2.multiselect("Filtrar por Zona:", zonas_disponibles, default=zonas_disponibles)

        semaforo_filtro = f3.multiselect(
            "Filtrar por Prioridad:",
            ["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (VENCE EN 8-15H)", "🔵 PROGRAMADA"],
            default=["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (VENCE EN 8-15H)", "🔵 PROGRAMADA"]
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
        doc_nov_ingresada = st.text_input("Digite la Cédula del Paciente:", key="input_cedula_nov").strip()
        if doc_nov_ingresada:
            p_data = run_query("SELECT documento, nombre, piso, zona FROM pacientes WHERE documento = ?", (doc_nov_ingresada,))
            if p_data:
                p_doc, p_nom, p_piso, p_zona = p_data[0]
                
                tto_actual_res = run_query("SELECT tratamiento, dosis, via_administracion, fecha_fin FROM tratamientos WHERE documento = ? AND estado = 'Activo'", (p_doc,))
                
                st.markdown(
                    f"""
                    <div style="background-color: #EBF3FC; border-left: 5px solid #0033A0; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px;">
                        <strong style="color: #0033A0; font-size: 1rem;">👤 Paciente: {p_nom}</strong><br>
                        <span style="font-size: 0.88rem; color: #334155;"><strong>Piso:</strong> {p_piso} | <strong>Zona:</strong> {p_zona}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if tto_actual_res:
                    t_act_nom, t_act_dos, t_act_via, t_act_fin = tto_actual_res[0]
                    fin_fmt_n = datetime.fromisoformat(t_act_fin).strftime('%d/%m/%Y %I:%M %p') if t_act_fin != "-" else "-"
                    st.markdown(
                        f"""
                        <div style="background-color: #F8FAFC; border: 1px solid #CBD5E1; padding: 10px 14px; border-radius: 8px; margin-bottom: 12px; font-size: 0.85rem;">
                            💊 <strong>Tratamiento Activo:</strong> {t_act_nom} ({t_act_dos}) · <strong>Vía:</strong> {t_act_via} · <strong>Fin:</strong> {fin_fmt_n}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    st.info("ℹ️ El paciente no tiene tratamientos activos en este momento.")

                c_reg1, c_reg2 = st.columns(2)
                tipo_nov_reg = c_reg1.selectbox("Tipo de Novedad:", ["Cambio de Horario", "Visita Fallida", "Descanalización / Salida", "Modificación Orden Médica", "Paciente Ausente", "Otra"], key="reg_tipo")
                
                nivel_nov_reg = c_reg2.selectbox("Nivel / Tipo de Acción:", ["Nivel 1: Informativa (Requiere Gestión)", "Nivel 2: Visita Puntual (Gestión Automática)", "Nivel 3: Cambio Permanente / Crítica (Gestión Automática)"], key="reg_nivel")

                c_fec1, c_fec2 = st.columns(2)
                fec_aplica_reg = c_fec1.date_input("Fecha en que Aplica:", value=datetime.now().date(), key="reg_fec")
                hor_aplica_reg = c_fec2.time_input("Hora en que Aplica:", value=datetime.now().time(), key="reg_hor")

                det_reg = st.text_area("Detalle de la Novedad (Escriba con claridad lo sucedido):", key="reg_det")
                prof_reg = render_selectbox("Profesional que Registra:", cat_gestores, "", key="reg_prof")

                if st.button("💾 Guardar Novedad en Bandeja"):
                    if not prof_reg.strip() or not det_reg.strip():
                        st.error("Diligencie todos los campos obligatorios.")
                    else:
                        nivel_txt_puro = nivel_nov_reg.split(":")[0].strip()
                        estado_inicial_nov = 'Pendiente' if "Nivel 1" in nivel_nov_reg else 'Gestionada'
                        fec_gestion_aut = datetime.now().strftime("%d/%m/%Y %I:%M %p") if estado_inicial_nov == 'Gestionada' else None

                        run_query(
                            """INSERT INTO registro_novedades (documento, tipo_novedad, nivel_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion, fecha_gestion, responsable_gestion)
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (p_doc, tipo_nov_reg, nivel_txt_puro, fec_aplica_reg.strftime("%d/%m/%Y"), hor_aplica_reg.strftime("%I:%M %p"), det_reg, prof_reg.strip(), estado_inicial_nov, datetime.now().isoformat(), fec_gestion_aut, prof_reg.strip() if fec_gestion_aut else None),
                            fetch=False,
                        )
                        
                        if "Nivel 3" in nivel_nov_reg:
                            run_query("UPDATE tratamientos SET novedades = novedades || ' // [CAMBIO PERMANENTE AUTOMÁTICO el ' || ? || ']' WHERE documento = ? AND estado = 'Activo'", (datetime.now().strftime("%d/%m/%Y"), p_doc), fetch=False)

                        run_query("UPDATE tratamientos SET alerta_revisada = 'NO' WHERE documento = ? AND estado = 'Activo'", (p_doc,), fetch=False)
                        st.success("✅ Novedad registrada y procesada exitosamente.")
                        st.rerun()
            else:
                st.warning("⚠️ Cédula no encontrada en la base de pacientes activos.")

# ----------------- SECCIÓN 4: AYUDAS DIAGNOSTICAS (DINÁMICA, LIMPIA Y CON ORDEN MÉDICA) -----------------
elif menu == "Ayudas diagnosticas":
    st.subheader("🧬 Gestión de Ayudas Diagnósticas, Peso y Oxígeno")
    st.markdown("---")

    with st.expander("➕ Solicitar Nueva Ayuda Diagnóstica o Insumo", expanded=True):
        st.markdown("#### 👤 Identificación del Paciente")
        
        c_tdoc_dx, c_doc_dx = st.columns([1, 2.5])
        tipos_doc_lista = ["CC", "TI", "CE", "PPT", "Pasaporte", "RC"]
        t_doc_dx = c_tdoc_dx.selectbox("Tipo Doc", tipos_doc_lista, index=0, key="tdoc_dx_ing")
        ced_dx_in = c_doc_dx.text_input("Digite el Documento del Paciente:", key="ced_dx_input_val").strip()
        
        paciente_existe_dx = False
        p_nombre_dx = ""
        p_piso_dx = "Norte"
        p_zona_dx = "Zona 1"
        
        if ced_dx_in:
            res_dx_p = run_query("SELECT nombre, tipo_documento, piso, zona FROM pacientes WHERE documento = ?", (ced_dx_in,))
            if res_dx_p:
                p_nombre_dx, p_tdoc_db, p_piso_dx, p_zona_dx = res_dx_p[0]
                if p_tdoc_db and p_tdoc_db.upper() != t_doc_dx.upper():
                    st.error(f"🚨 **Error de tipo de documento:** El paciente con documento `{ced_dx_in}` está registrado bajo el tipo **{p_tdoc_db}**, no **{t_doc_dx}**.")
                else:
                    paciente_existe_dx = True
                    st.success(f"✅ Paciente Activo Censo: **{p_nombre_dx}** ({p_tdoc_db}) | Piso automático: **{p_piso_dx}** | Zona automática: **{p_zona_dx}**")
            else:
                st.error(f"🚨 **¡ALERTA!** El paciente con documento `{ced_dx_in}` no existe en la base activa.")
                st.warning("Puedes registrar la ayuda, pero se recomienda crearlo primero en el módulo de 'Gestor Pacientes'.")

        # SOLO SE DESPLIEGA LA SOLICITUD SI SE HA DIGITADO DOCUMENTO VÁLIDO
        if ced_dx_in:
            st.markdown("---")
            st.markdown("#### 🏥 Solicitud de Ayuda Diagnóstica")
            c_d1, c_d2 = st.columns(2)
            piso_dx_sel = c_d1.selectbox("Piso del Paciente:", ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"], index=(["Norte", "Sur", "Larga Estancia", "AMI", "PAS"].index(p_piso_dx) if p_piso_dx in ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"] else 0))
            zona_dx_sel = c_d2.selectbox("Zona del Paciente:", ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"], index=(["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"].index(p_zona_dx) if p_zona_dx in ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"] else 0))

            ayuda_sol_txt = st.text_input("Ayuda Diagnóstica / Examen Solicitado (Ej: Ecografía Doppler, TAC, Laboratorios):")

            c_orig1, c_orig2 = st.columns(2)
            origen_sol_sel = c_orig1.selectbox("Origen de Solicitud:", ["Línea de Ayudas Diagnósticas", "PGP"])
            
            prestador_pgp_txt = ""
            if origen_sol_sel == "PGP":
                prestador_pgp_txt = c_orig2.text_input("Indique el Prestador PGP (Obligatorio):")
            else:
                c_orig2.info("Origen Línea de Ayudas SENC")

            st.markdown("---")
            st.markdown("#### ⚖️ Requerimientos Clínicos y Logísticos")
            c_d3, c_d4, c_d5 = st.columns(3)
            peso_pac_txt = c_d3.text_input("Peso Actual (Kg):", value="")
            req_o2_sel = c_d4.selectbox("¿Requiere Oxígeno?", ["NO", "SI"], index=0)
            
            cant_o2_txt = ""
            if req_o2_sel == "SI":
                cant_o2_txt = c_d5.text_input("Cantidad / Flujo (Ej: 2 Litros / Minuto):")
            else:
                c_d5.info("Oxígeno no requerido")

            transporte_dx_sel = st.selectbox(
                "Requerimiento de Transporte para Cita:",
                ["No requiere transporte", "TAB", "TAM", "MVR"]
            )

            obs_dx_txt = st.text_area("Observaciones de Cita / Lugar / Detalles de Transporte:")
            gestor_dx_reg = render_selectbox("Registrado por (Gestor):", cat_gestores, "", key="gestor_dx_select")

            if st.button("💾 Guardar Solicitud de Ayuda Diagnóstica"):
                if not ced_dx_in or not ayuda_sol_txt.strip() or not gestor_dx_reg.strip():
                    st.error("Ingrese la cédula, la ayuda diagnóstica solicitada y seleccione el gestor.")
                elif origen_sol_sel == "PGP" and not prestador_pgp_txt.strip():
                    st.error("Si el origen es PGP, debe indicar el nombre del prestador.")
                else:
                    fec_reg_dx = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                    run_query(
                        """INSERT INTO ayudas_diagnosticas_v6 (documento, tipo_documento, piso, zona, ayuda_solicitada, peso, requiere_oxigeno, cantidad_oxigeno, tipo_transporte, origen_solicitud, prestador_pgp, observaciones, estado, registrado_por, fecha_registro)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente', ?, ?)""",
                        (ced_dx_in, t_doc_dx, piso_dx_sel, zona_dx_sel, ayuda_sol_txt.strip(), peso_pac_txt.strip(), req_o2_sel, cant_o2_txt.strip(), transporte_dx_sel, origen_sol_sel, prestador_pgp_txt.strip(), obs_dx_txt.strip(), gestor_dx_reg.strip(), fec_reg_dx),
                        fetch=False
                    )
                    st.success("✅ Ayuda diagnóstica guardada correctamente. Formulario limpiado.")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### ⏳ Panel de Ayudas Diagnósticas (Pendientes y Resueltas)")
    
    col_fil_dx1, col_fil_dx2 = st.columns(2)
    pisos_disponibles_dx = ["Todos"] + ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
    piso_dx_filtro = col_fil_dx1.selectbox("Filtrar por Piso:", pisos_disponibles_dx)
    ver_completadas_dx = col_fil_dx2.checkbox("📂 Ver Ayudas Diagnósticas Gestionadas / Completadas", key="chk_comp_dx")

    q_dx = "SELECT * FROM ayudas_diagnosticas_v6 WHERE 1=1"
    if not ver_completadas_dx:
        q_dx += " AND estado = 'Pendiente'"
    else:
        q_dx += " AND estado = 'Completada'"
    
    if piso_dx_filtro != "Todos":
        q_dx += f" AND piso = '{piso_dx_filtro}'"
    
    q_dx += " ORDER BY id DESC"
    
    df_dx = pd.read_sql_query(q_dx, sqlite3.connect(DB_FILE))
    
    if not df_dx.empty:
        for _, r_dx in df_dx.iterrows():
            dx_id = r_dx['id']
            dx_doc = r_dx['documento']
            dx_tdoc = r_dx.get('tipo_documento', 'CC')
            dx_piso = r_dx.get('piso', 'Norte')
            dx_zona = r_dx.get('zona', 'Zona 1')
            dx_ayuda = r_dx['ayuda_solicitada']
            dx_peso = r_dx['peso']
            dx_req_o2 = r_dx['requiere_oxigeno']
            dx_cant_o2 = r_dx['cantidad_oxigeno']
            dx_transporte = r_dx.get('tipo_transporte', 'No requiere transporte')
            dx_origen = r_dx.get('origen_solicitud', 'Línea de Ayudas Diagnósticas')
            dx_prestador = r_dx.get('prestador_pgp', '')
            dx_obs = r_dx['observaciones']
            dx_est = r_dx['estado']
            dx_reg = r_dx['registrado_por']
            dx_fec = r_dx['fecha_registro']
            dx_fec_cita = r_dx.get('fecha_cita', '')
            dx_hora_cita = r_dx.get('hora_cita', '')
            dx_lugar_cita = r_dx.get('lugar_cita', '')
            dx_informado = r_dx.get('paciente_informado', 'NO')
            dx_nom_acomp = r_dx.get('nombre_acompanante', '')
            dx_cont_acomp = r_dx.get('contacto_acompanante', '')
            dx_resuelto = r_dx.get('resuelto_por', '')

            p_nom_dx_v = dx_doc
            res_nom_dx = run_query("SELECT nombre FROM pacientes WHERE documento = ?", (dx_doc,))
            if res_nom_dx:
                p_nom_dx_v = res_nom_dx[0][0]

            badge_o2 = f" | Oxígeno: {dx_cant_o2}" if dx_req_o2 == "SI" else " | Sin Oxígeno"
            badge_peso = f" | Peso: {dx_peso}kg" if dx_peso else ""
            badge_trans = f" | Transporte: {dx_transporte}" if dx_transporte != "No requiere transporte" else " | Sin Transporte"
            badge_orig = f" | Origen: {dx_origen}" + (f" ({dx_prestador})" if dx_prestador else "")

            with st.expander(f"🧬 [Piso {dx_piso} - {dx_zona}] · {p_nom_dx_v} ({dx_tdoc} {dx_doc}) - Examen: {dx_ayuda} ({dx_est})", expanded=(dx_est == 'Pendiente')):
                st.markdown(f"**Ayuda Solicitada:** {dx_ayuda}{badge_orig}")
                st.markdown(f"**Requerimientos:** {badge_trans}{badge_o2}{badge_peso}")
                if dx_obs:
                    st.markdown(f"**Observaciones:** {dx_obs}")
                st.markdown(f"*Registró: {dx_reg} el {dx_fec}*")

                if dx_est == 'Completada':
                    st.markdown(f"---\n✅ **DATOS DE CIERRE:** Fecha Cita: {dx_fec_cita} {dx_hora_cita} | Lugar: {dx_lugar_cita} | Paciente Informado: **{dx_informado}**" + (f" | Acompañante: {dx_nom_acomp} ({dx_cont_acomp})" if dx_informado=="SI" else "") + f" | Diligenció: {dx_resuelto}")

                st.markdown("---")
                st.markdown("#### 📝 Formulario de Solución / Cita Asignada y Generación de Transporte")
                
                c_c1, c_c2 = st.columns(2)
                fec_cita_val = c_c1.date_input("Fecha de la Cita (Obligatorio):", value=datetime.now().date(), key=f"fec_cita_{dx_id}")
                hora_cita_val = c_c2.time_input("Hora de la Cita (Obligatorio):", value=datetime.now().time(), key=f"hora_cita_{dx_id}")
                
                lugar_cita_val = st.text_input("Lugar de la Cita (Obligatorio):", placeholder="Ej: Clínica Las Vegas Cons. 454", key=f"lugar_cita_{dx_id}", value=dx_lugar_cita)

                c_c3, c_c4 = st.columns(2)
                info_pac_val = c_c3.selectbox("¿Paciente ya fue informado?", ["NO", "SI"], index=1 if dx_informado=="SI" else 0, key=f"info_pac_{dx_id}")
                
                nombre_acomp_val, contacto_acomp_val = dx_nom_acomp, dx_cont_acomp
                if info_pac_val == "SI":
                    st.markdown("##### 👥 Datos del Acompañante")
                    ac1, ac2 = st.columns(2)
                    nombre_acomp_val = ac1.text_input("Nombre del Acompañante:", value=dx_nom_acomp, key=f"nom_acomp_{dx_id}")
                    contacto_acomp_val = ac2.text_input("Contacto del Acompañante:", value=dx_cont_acomp, key=f"cont_acomp_{dx_id}")

                gestor_res_val = render_selectbox("Diligenciado por (Gestor):", cat_gestores, dx_resuelto, key=f"gestor_res_{dx_id}")

                st.markdown("---")
                btn_col_res, btn_col_trans = st.columns(2)

                with btn_col_res:
                    if st.button(f"✅ Marcar como Resuelta (#{dx_id})", key=f"btn_res_dx_{dx_id}"):
                        if not str(fec_cita_val).strip() or not lugar_cita_val.strip() or not gestor_res_val.strip():
                            st.error("Debe diligenciar la fecha, hora, lugar de la cita y el gestor.")
                        else:
                            fec_ges_dx = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            f_cita_str = fec_cita_val.strftime("%d/%m/%Y")
                            h_cita_str = hora_cita_val.strftime("%I:%M %p")
                            run_query(
                                """UPDATE ayudas_diagnosticas_v6 
                                   SET estado = 'Completada', fecha_cita = ?, hora_cita = ?, lugar_cita = ?, paciente_informado = ?, nombre_acompanante = ?, contacto_acompanante = ?, resuelto_por = ?, fecha_gestion = ? 
                                   WHERE id = ?""",
                                (f_cita_str, h_cita_str, lugar_cita_val.strip(), info_pac_val, nombre_acomp_val.strip(), contacto_acomp_val.strip(), gestor_res_val.strip(), fec_ges_dx, dx_id),
                                fetch=False
                            )
                            st.success("✅ Ayuda diagnóstica solucionada y guardada con éxito.")
                            st.rerun()

                with btn_col_trans:
                    if st.button(f"🚗 Generar Solicitud de Transporte (#{dx_id})", key=f"btn_gen_trans_{dx_id}"):
                        if not str(fec_cita_val).strip() or not lugar_cita_val.strip():
                            st.error("⚠️ Para generar la solicitud de transporte, debe ingresar primero la Fecha, Hora y el Lugar de la cita.")
                        else:
                            f_cita_str = fec_cita_val.strftime("%d/%m/%Y")
                            h_cita_str = hora_cita_val.strftime("%I:%M %p")
                            
                            # TEXTO ESTILO ORDEN MÉDICA CON CÉDULA ABAJO Y SIN OBSERVACIONES
                            texto_correo_transporte = f"""me colaboras generando la siguiente orden de ayudas dx {dx_ayuda} para el paciente en mencion.

Paciente: {p_nom_dx_v}
CC: {dx_doc}
Piso: {dx_piso} | Zona: {dx_zona}
Ayuda Diagnóstica: {dx_ayuda}
Fecha y Hora de la Cita: {f_cita_str} {h_cita_str}
Lugar de la Cita: {lugar_cita_val.strip()}
Tipo de Transporte Requerido: {dx_transporte}
Oxígeno: {dx_req_o2} ({dx_cant_o2 if dx_req_o2=='SI' else 'N/A'})
Peso: {dx_peso if dx_peso else 'No registrado'} kg
Paciente Informado: {info_pac_val}"""

                            if info_pac_val == "SI" and nombre_acomp_val.strip():
                                texto_correo_transporte += f"\nAcompañante: {nombre_acomp_val.strip()} (Contacto: {contacto_acomp_val.strip()})"

                            texto_correo_transporte += f"""
--------------------------------------------------
Diligenciado por: {gestor_res_val if gestor_res_val else dx_reg}
"""
                            st.markdown("##### 📋 Texto generado listo para copiar y pegar en el correo:")
                            st.code(texto_correo_transporte, language="text")
                            st.info("💡 Copia el texto superior y pégalo directamente en tu cliente de correo corporativo.")
    else:
        st.info("No hay ayudas diagnósticas bajo los filtros seleccionados.")

# ----------------- SECCIÓN 5: CENSO Y PLANILLA EN VIVO -----------------
elif menu == "Censo y Planilla en Vivo":
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
                row["Vía"],
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
        total_descanalizar = len(df_filtrado[df_filtrado["Estado Tratamiento"].str.contains("SIN TTO ACTIVO|SEGUIMIENTO") & (df_filtrado["Estado"] == "Activo")])

        m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
        m1.markdown(f"<div class='metric-card'><div class='metric-val'>{total_filtrados}</div><div class='metric-lbl'>ACTIVOS</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-card'><div class='metric-val'>{total_picc}</div><div class='metric-lbl'>🩸 PICC</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-card'><div class='metric-val'>{total_npt}</div><div class='metric-lbl'>🥣 NPT</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-card'><div class='metric-val'>{total_be}</div><div class='metric-lbl'>💉 BE / DOLOR</div></div>", unsafe_allow_html=True)
        m5.markdown(f"<div class='metric-card'><div class='metric-val'>{total_aislados}</div><div class='metric-lbl'>☣️ AISLADOS</div></div>", unsafe_allow_html=True)
        m6.markdown(f"<div class='metric-card'><div class='metric-val'>{total_nov_act}</div><div class='metric-lbl'>📢 NOVEDADES</div></div>", unsafe_allow_html=True)
        m7.markdown(f"<div class='metric-card'><div class='metric-val'>{total_descanalizar}</div><div class='metric-lbl'>⚠️ PEND. EGRESO</div></div>", unsafe_allow_html=True)

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
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_PUSH};'>🟦 Vía PUSH / IV</div>", unsafe_allow_html=True)
        with c_c4:
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_AISLA};'>🟣 Paciente Aislado</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='conv-box' style='background-color: {COLOR_FALTA1};'>⚠️ Sin TTO / Pendiente Egreso</div>", unsafe_allow_html=True)

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

# ----------------- SECCIÓN 6: ANALITICA Y GRAFICOS -----------------
elif menu == "Analitica y graficos":
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

# ----------------- SECCIÓN 7: INFORMES -----------------
elif menu == "Informes":
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

# ----------------- PIE DE PÁGINA INSTITUCIONAL CON TU AUTORÍA -----------------
st.markdown(
    """
    <div class="footer-autor">
        🏥 <strong>Sistema de Gestión y Control de Tratamientos Domiciliarios (SENC)</strong><br>
        Diseñado y desarrollado por <strong>Camilo Andrés Medina</strong> · Salud en Casa SURA Colombia
    </div>
""",
    unsafe_allow_html=True,
)
ñ
