import hashlib
import io
import json
import logging
import math
import os
import re
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ----------------- CONFIGURACIÓN Y CONSTANTES -----------------
DB_FILE = "control_pacientes.db"
LOGO_SURA_URL = "https://upload.wikimedia.org/wikipedia/commons/6/61/Seguros_SURA_Logo.svg"
LOGO_SURA_FAVICON = LOGO_SURA_URL

TIPOS_DOC = ["CC", "TI", "CE", "PPT", "Pasaporte", "RC"]
PLANES = ["POS", "Póliza", "ARL"]
PROGRAMAS = ["Agudos"]
PISOS = ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
ZONAS = ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8", "PAS", "AMI"]
MUNICIPIOS = [
    "Medellín", "Bello", "Itagüí", "Envigado", "Sabaneta", "Caldas",
    "La Estrella", "Copacabana", "Girardota", "Barbosa", "Rionegro"
]
AISLAMIENTOS_TIPOS = ["Contacto", "Gotas", "Aerosoles", "Protector (Invertido)"]
VIAS = ["IV", "VO", "SC", "IM", "BE", "BE analgesia", "NPT", "Tuberculosis", "Materna", "NBZ", "PUSH", "LEV"]
ACCESOS = ["Periférico", "PICC", "SC", "Ninguno"]
FRECUENCIAS = [4, 6, 8, 12, 24]

# Convenciones de color clínico estrictas
COLOR_PICC = "#FFD1D1"
COLOR_NPT = "#FFF3CD"
COLOR_BE = "#D1E7DD"
COLOR_ANALGESIA = "#E2D9F3"
COLOR_TB = "#FFE5D0"
COLOR_PUSH = "#CFF4FC"
COLOR_AISLA = "#E8D7F1"
COLOR_FALTA1 = "#FFE0B2"

MENU_ITEMS = [
    "Gestor Pacientes",
    "Entrega de turno",
    "Novedades programacion",
    "Ayudas diagnosticas",
    "Censo y Planilla en Vivo",
    "Analitica y graficos",
    "Informes",
]

st.set_page_config(
    page_title="SURA - Base Agudos",
    page_icon=LOGO_SURA_FAVICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

logger = logging.getLogger("senc_app")

# ----------------- ESTILOS CLÍNICOS NITIDOS (ALTO CONTRASTE) -----------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
    
    /* Forzado universal para que SIEMPRE sea visible en modo claro y oscuro */
    .stApp {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #051022 0%, #0A1C38 45%, #07152B 100%) !important;
        padding-top: 1.5rem !important;
        border-right: 1px solid rgba(0, 212, 255, 0.15) !important;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] .stRadio label {
        background: rgba(255, 255, 255, 0.04) !important;
        padding: 10px 16px !important;
        border-radius: 10px !important;
        margin-bottom: 6px !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        font-weight: 500 !important;
    }
    [data-testid="stSidebar"] .stRadio label:hover {
        background: linear-gradient(90deg, rgba(0, 212, 255, 0.2) 0%, rgba(0, 51, 160, 0.3) 100%) !important;
        border-color: #00D4FF !important;
    }

    /* LABELS Y TEXTOS: SIEMPRE NEGROS Y LEGIBLES */
    label, [data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] span {
        color: #0F172A !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
    }

    /* INPUTS, TEXTAREAS Y SELECTORES */
    div[data-baseweb="input"] input, 
    div[data-baseweb="textarea"] textarea,
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stDateInput"] input,
    [data-testid="stTimeInput"] input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] * {
        color: #0F172A !important;
    }

    /* EXPANDERS BLANCOS Y NÍTIDOS */
    details[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.02) !important;
        margin-bottom: 12px !important;
    }
    details[data-testid="stExpander"] summary {
        color: #0033A0 !important;
        font-weight: 700 !important;
    }
    details[data-testid="stExpander"] div[data-testid="stExpanderDetails"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }

    /* ENCABEZADOS Y TARJETAS */
    .hospital-header {
        background: #FFFFFF;
        border-radius: 16px;
        padding: 1.2rem 2rem;
        box-shadow: 0px 8px 24px rgba(0, 51, 160, 0.05);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1.5rem;
        border-left: 6px solid #0033A0;
        border: 1px solid #E2E8F0;
    }
    .hospital-title {
        font-size: 1.5rem;
        font-weight: 800;
        color: #0033A0;
        margin: 0;
    }
    .hospital-sub {
        font-size: 0.82rem;
        color: #00AEC7;
        margin: 2px 0 0 0;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    .metric-card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 1.1rem 0.8rem;
        box-shadow: 0px 4px 16px rgba(0, 26, 77, 0.04);
        border: 1px solid #E2E8F0;
        text-align: center;
        margin-bottom: 0.8rem;
    }
    .metric-val {
        font-size: 1.8rem;
        font-weight: 800;
        color: #0033A0;
    }
    .metric-lbl {
        font-size: 0.72rem;
        color: #64748B;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-top: 4px;
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #0033A0 0%, #001F6B 100%) !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.4rem !important;
        border: 1px solid rgba(0, 212, 255, 0.2) !important;
        box-shadow: 0px 4px 14px rgba(0, 51, 160, 0.18) !important;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(135deg, #00AEC7 0%, #0033A0 100%) !important;
        box-shadow: 0px 6px 18px rgba(0, 212, 255, 0.35) !important;
    }

    .paciente-sticky-header {
        position: sticky;
        top: 0;
        z-index: 999;
        background: linear-gradient(135deg, #05142E 0%, #002270 100%);
        color: #FFFFFF;
        padding: 12px 22px;
        border-radius: 12px;
        box-shadow: 0px 8px 24px rgba(0, 26, 77, 0.3);
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 3px solid #00D4FF;
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
        border-radius: 12px;
        padding: 16px 22px;
        box-shadow: 0px 4px 14px rgba(0, 0, 0, 0.03);
        margin-bottom: 14px;
    }
    .tag-turno {
        display: inline-block;
        background: #F1F5F9;
        color: #0F172A;
        border: 1px solid #CBD5E1;
        padding: 4px 10px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 0.82rem;
        margin-right: 6px;
        font-family: 'JetBrains Mono', monospace;
    }

    .footer-autor {
        text-align: center;
        padding: 1.5rem 0;
        margin-top: 3.5rem;
        border-top: 1px solid #E2E8F0;
        color: #64748B;
        font-size: 0.85rem;
        background: #FFFFFF;
        border-radius: 12px;
    }
    .footer-autor strong {
        color: #0033A0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------- BASE DE DATOS Y CONTEXT MANAGER -----------------
@contextmanager
def get_db():
    conn = sqlite3.connect(DB_FILE, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    try:
        yield conn
        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f"Error en BD: {e}", exc_info=True)
        raise
    finally:
        conn.close()

def run_query(query: str, params: tuple = (), fetch: bool = True):
    try:
        with get_db() as conn:
            c = conn.cursor()
            c.execute(query, params)
            if fetch:
                return [dict(r) for r in c.fetchall()]
            return None
    except Exception as e:
        logger.error(f"Error en query '{query}': {e}")
        return [] if fetch else None

def log_auditoria(usuario: str, accion: str, tabla: str, registro_id: str, detalle: str = ""):
    try:
        run_query(
            """INSERT INTO auditoria (usuario, accion, tabla_afectada, registro_id, detalle, fecha)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (usuario or "SISTEMA", accion, tabla, str(registro_id), detalle, datetime.now().isoformat()),
            fetch=False,
        )
    except Exception as e:
        logger.warning(f"Error al registrar auditoría: {e}")

def init_db():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                nombre_completo TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                rol TEXT DEFAULT 'usuario',
                estado TEXT DEFAULT 'pendiente',
                fecha_registro TEXT NOT NULL,
                aprobado_por TEXT,
                fecha_aprobacion TEXT
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS auditoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT NOT NULL,
                accion TEXT NOT NULL,
                tabla_afectada TEXT NOT NULL,
                registro_id TEXT NOT NULL,
                detalle TEXT DEFAULT '',
                fecha TEXT NOT NULL
            )
        """)
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
                t1 TEXT, t2 TEXT, t3 TEXT,
                t1_futuro TEXT DEFAULT '', t2_futuro TEXT DEFAULT '', t3_futuro TEXT DEFAULT '',
                fecha_aplica_futuro TEXT DEFAULT '',
                puntual_fecha TEXT DEFAULT '', puntual_turno TEXT DEFAULT '', puntual_hora TEXT DEFAULT '',
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
        c.execute("""
            CREATE TABLE IF NOT EXISTS ayudas_diagnosticas (
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

        def asegurar_columnas(tabla, columnas_dict):
            c.execute(f"PRAGMA table_info({tabla})")
            cols = [col[1] for col in c.fetchall()]
            for col_nombre, col_tipo in columnas_dict.items():
                if col_nombre not in cols:
                    try:
                        c.execute(f"ALTER TABLE {tabla} ADD COLUMN {col_nombre} {col_tipo}")
                    except Exception:
                        pass

        asegurar_columnas("pacientes", {
            "tipo_documento": "TEXT DEFAULT 'CC'",
            "observaciones_clinicas": "TEXT DEFAULT ''",
            "alerta_permanente": "TEXT DEFAULT ''",
            "fecha_ingreso": "TEXT"
        })
        asegurar_columnas("ayudas_diagnosticas", {
            "tipo_documento": "TEXT DEFAULT 'CC'",
            "piso": "TEXT DEFAULT 'Norte'",
            "zona": "TEXT DEFAULT 'Zona 1'",
            "peso": "TEXT DEFAULT ''",
            "requiere_oxigeno": "TEXT DEFAULT 'NO'",
            "cantidad_oxigeno": "TEXT DEFAULT ''",
            "tipo_transporte": "TEXT DEFAULT 'No requiere transporte'",
            "origen_solicitud": "TEXT DEFAULT 'Línea de Ayudas Diagnósticas'",
            "prestador_pgp": "TEXT DEFAULT ''",
            "fecha_cita": "TEXT DEFAULT ''",
            "hora_cita": "TEXT DEFAULT ''",
            "lugar_cita": "TEXT DEFAULT ''",
            "paciente_informado": "TEXT DEFAULT 'NO'",
            "nombre_acompanante": "TEXT DEFAULT ''",
            "contacto_acompanante": "TEXT DEFAULT ''",
            "resuelto_por": "TEXT DEFAULT ''"
        })

init_db()

# ----------------- UTILIDADES Y CATÁLOGOS -----------------
def safe_int(valor, default=0):
    try:
        if valor is None or str(valor).strip() == "":
            return default
        return int(float(valor))
    except Exception:
        return default

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

@st.cache_data(ttl=3600)
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
        except Exception:
            pass
    return ["", fallback_val]

def render_selectbox(label, options, current_val, key=None, disabled=False):
    safe_options = list(options)
    if current_val and current_val not in safe_options:
        safe_options.insert(0, current_val)
    idx = safe_options.index(current_val) if current_val in safe_options else 0
    if key:
        return st.selectbox(label, safe_options, index=idx, key=key, disabled=disabled)
    return st.selectbox(label, safe_options, index=idx, disabled=disabled)

cat_meds_agudos = load_excel_catalog("BASE MEDICAMENTOS AGUDOS.xlsx", "MEDICAMENTO GENÉRICO")
cat_meds_be = load_excel_catalog("BASE MEDICAMENTOS BE AGUDOS.xlsx", "MEDICAMENTO BE GENÉRICO")
cat_gestores = load_excel_catalog("GESTORES DE AGUDOS.xlsx", "GESTOR GENÉRICO")

# ----------------- CÁLCULOS CLÍNICOS EXACTOS (100% PRESERVADOS) -----------------
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

def get_color_fila_hex(row):
    via = str(row.get("Vía", row.get("via_administracion", ""))).upper()
    acceso = str(row.get("Tipo Acceso", row.get("tipo_acceso", ""))).upper()
    alerta = str(row.get("Estado Tratamiento", ""))
    aisla = str(row.get("Aislamiento", row.get("aislamiento", ""))).upper()

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

# ----------------- EXPORTACIÓN EXCEL PROFESIONAL -----------------
def exportar_excel_con_colores(df, sheet_name="Planilla_SURA"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = workbook[sheet_name]

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

def exportar_excel_multi_hoja(dict_dfs):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dict_dfs.items():
            if not df.empty:
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                ws = writer.sheets[sheet_name[:31]]
                h_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                h_fill = PatternFill(start_color="0033A0", end_color="0033A0", fill_type="solid")
                for col in range(1, len(df.columns) + 1):
                    c = ws.cell(row=1, column=col)
                    c.font = h_font
                    c.fill = h_fill
    return output.getvalue()

def generar_pie_chart_svg(datos, etiquetas, colores, titulo):
    total = sum(datos)
    if total == 0:
        return f"<div style='text-align: center; color: #6C757D; padding: 25px;'><strong>{titulo}</strong><br>Sin datos en este rango</div>"

    cx, cy, r = 100, 100, 80
    current_angle = 0
    slices_svg = ""
    legend_html = "<div style='margin-top: 14px; font-size: 0.82rem;'>"

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
        legend_html += f"<span style='display:inline-block; margin-right: 14px; margin-bottom: 6px;'><span style='display:inline-block; width:11px; height:11px; background-color:{color}; border-radius:50%; margin-right: 5px; vertical-align: middle;'></span><strong>{label}:</strong> {val} ({pct:.1f}%)</span>"

    legend_html += "</div>"

    return f"""
    <div style="background: white; border-radius: 14px; padding: 18px; border: 1px solid #E2E8F0; text-align: center; box-shadow: 0px 4px 14px rgba(0,26,77,0.03);">
        <strong style="font-size: 1rem; color: #0033A0;">{titulo}</strong><br><br>
        <svg width="200" height="200" viewBox="0 0 200 200" style="display: block; margin: 0 auto;">
            {slices_svg}
        </svg>
        {legend_html}
    </div>
    """

# ----------------- AUTENTICACIÓN ROBUSTA (PBKDF2) -----------------
def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 100_000)
    return f"{salt}:{key.hex()}"

def verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt, expected_hex = stored_hash.split(":")
        actual_key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 100_000)
        return secrets.compare_digest(actual_key.hex(), expected_hex)
    except Exception:
        return False

def check_session() -> bool:
    return st.session_state.get("auth_logged_in", False)

def get_current_user() -> dict:
    if not check_session():
        return {}
    return {
        "username": st.session_state.get("auth_username", ""),
        "nombre_completo": st.session_state.get("auth_nombre", ""),
        "rol": st.session_state.get("auth_rol", "usuario")
    }

def is_admin() -> bool:
    return st.session_state.get("auth_rol", "") == "admin"

def logout():
    st.session_state["auth_logged_in"] = False
    st.session_state["auth_username"] = ""
    st.session_state["auth_nombre"] = ""
    st.session_state["auth_rol"] = ""
    st.rerun()

def login_page():
    admin_check = run_query("SELECT COUNT(*) as cant FROM usuarios WHERE rol = 'admin'")
    num_admins = admin_check[0]["cant"] if admin_check else 0

    st.markdown("""
        <div style="text-align: center; margin-top: 2rem; margin-bottom: 2rem;">
            <div style="display: inline-block; padding: 12px 24px; background: rgba(0, 212, 255, 0.08); border: 1px solid rgba(0, 212, 255, 0.3); border-radius: 50px; margin-bottom: 15px;">
                <span style="color: #00D4FF; font-weight: 700; font-size: 0.85rem; letter-spacing: 1.5px; text-transform: uppercase;">
                    🛡️ Acceso Seguro Red Asistencial · SENC
                </span>
            </div>
            <h1 style="color: #0033A0; font-weight: 800; font-size: 2.2rem;">
                Base y Control de Agudos
            </h1>
            <p style="color: #64748B; font-size: 1rem; max-width: 500px; margin: 0 auto;">
                Gestión Clínica Hospitalaria & Tratamientos Domiciliarios Especializados
            </p>
        </div>
    """, unsafe_allow_html=True)

    _, col_l2, _ = st.columns([1, 1.8, 1])
    with col_l2:
        if num_admins == 0:
            st.warning("⚠️ **Inicialización:** No se ha detectado administrador. Registra la cuenta del Administrador Principal.")
            with st.form("form_primer_admin"):
                st.subheader("👑 Registro Administrador Principal")
                nom_adm = st.text_input("Nombre Completo:")
                usr_adm = st.text_input("Usuario de Red:")
                pwd_adm = st.text_input("Contraseña:", type="password")
                pwd_adm2 = st.text_input("Confirmar Contraseña:", type="password")
                submit_adm = st.form_submit_button("Crear Administrador Principal")

                if submit_adm:
                    if not nom_adm.strip() or not usr_adm.strip() or not pwd_adm.strip():
                        st.error("Completa todos los campos.")
                    elif pwd_adm != pwd_adm2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        phash = hash_password(pwd_adm)
                        ahora = datetime.now().isoformat()
                        run_query(
                            """INSERT INTO usuarios (username, nombre_completo, password_hash, rol, estado, fecha_registro, aprobado_por, fecha_aprobacion)
                               VALUES (?, ?, ?, 'admin', 'aprobado', ?, 'SISTEMA_AUTO', ?)""",
                            (usr_adm.strip().lower(), nom_adm.strip(), phash, ahora, ahora),
                            fetch=False
                        )
                        log_auditoria(usr_adm.strip(), "CREAR_PRIMER_ADMIN", "usuarios", usr_adm.strip(), "Admin inicial registrado")
                        st.success("✅ Administrador creado correctamente. Ya puedes iniciar sesión.")
                        st.rerun()
            return

        tab_ingreso, tab_registro = st.tabs(["🔐 Iniciar Sesión", "📝 Solicitar Registro"])

        with tab_ingreso:
            with st.form("form_login"):
                u_in = st.text_input("Usuario:")
                p_in = st.text_input("Contraseña:", type="password")
                btn_login = st.form_submit_button("Ingresar", use_container_width=True)

                if btn_login:
                    if not u_in.strip() or not p_in.strip():
                        st.error("Ingresa usuario y contraseña.")
                    else:
                        res_u = run_query("SELECT * FROM usuarios WHERE username = ?", (u_in.strip().lower(),))
                        if not res_u:
                            st.error("Usuario no encontrado.")
                        else:
                            u_data = res_u[0]
                            if u_data["estado"] == "pendiente":
                                st.warning("⏳ Tu usuario está en cola. Un administrador debe aprobar tu acceso.")
                            elif u_data["estado"] == "rechazado":
                                st.error("❌ Solicitud rechazada. Contacta a tu gestor de servicio.")
                            elif verify_password(p_in, u_data["password_hash"]):
                                st.session_state["auth_logged_in"] = True
                                st.session_state["auth_username"] = u_data["username"]
                                st.session_state["auth_nombre"] = u_data["nombre_completo"]
                                st.session_state["auth_rol"] = u_data["rol"]
                                log_auditoria(u_data["username"], "LOGIN", "usuarios", str(u_data["id"]), "Inicio exitoso")
                                st.rerun()
                            else:
                                st.error("Contraseña incorrecta.")

        with tab_registro:
            st.info("💡 Tu cuenta quedará en estado pendiente hasta la aprobación del administrador de red.")
            with st.form("form_solicitud"):
                reg_nom = st.text_input("Nombre Completo:")
                reg_usr = st.text_input("Usuario deseado:")
                reg_pwd = st.text_input("Contraseña:", type="password")
                reg_pwd2 = st.text_input("Confirmar Contraseña:", type="password")
                btn_solicitar = st.form_submit_button("Solicitar Acceso", use_container_width=True)

                if btn_solicitar:
                    if not reg_nom.strip() or not reg_usr.strip() or not reg_pwd.strip():
                        st.error("Diligencia todos los campos.")
                    elif reg_pwd != reg_pwd2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        existe = run_query("SELECT id FROM usuarios WHERE username = ?", (reg_usr.strip().lower(),))
                        if existe:
                            st.error("El usuario ya existe.")
                        else:
                            phash = hash_password(reg_pwd)
                            ahora = datetime.now().isoformat()
                            run_query(
                                """INSERT INTO usuarios (username, nombre_completo, password_hash, rol, estado, fecha_registro)
                                   VALUES (?, ?, ?, 'usuario', 'pendiente', ?)""",
                                (reg_usr.strip().lower(), reg_nom.strip(), phash, ahora),
                                fetch=False
                            )
                            log_auditoria(reg_usr.strip(), "SOLICITUD_ACCESO", "usuarios", reg_usr.strip(), f"Solicitud {reg_nom.strip()}")
                            st.success("✅ Solicitud enviada exitosamente. Solicita la aprobación al administrador.")

def render_admin_panel():
    st.subheader("🛡️ Panel Administrativo — Control de Accesos, Delegación y Roles")
    pendientes = run_query("SELECT id, username, nombre_completo, fecha_registro FROM usuarios WHERE estado = 'pendiente' ORDER BY id ASC")
    
    st.markdown("#### ⏳ Solicitudes Pendientes de Aprobación")
    if pendientes:
        for p in pendientes:
            c_p1, c_p2, c_p3 = st.columns([3, 1, 1])
            c_p1.markdown(f"**{p['nombre_completo']}** (`{p['username']}`) — *Solicitud: {p['fecha_registro'][:16]}*")
            if c_p2.button("✅ Aprobar", key=f"apr_{p['id']}"):
                ahora = datetime.now().isoformat()
                admin_actual = st.session_state.get("auth_username", "admin")
                run_query(
                    "UPDATE usuarios SET estado = 'aprobado', aprobado_por = ?, fecha_aprobacion = ? WHERE id = ?",
                    (admin_actual, ahora, p['id']),
                    fetch=False
                )
                log_auditoria(admin_actual, "APROBAR_USUARIO", "usuarios", str(p['id']), f"Aprobó a {p['username']}")
                st.success(f"Usuario {p['username']} aprobado.")
                st.rerun()

            if c_p3.button("❌ Rechazar", key=f"rec_{p['id']}"):
                admin_actual = st.session_state.get("auth_username", "admin")
                run_query("UPDATE usuarios SET estado = 'rechazado' WHERE id = ?", (p['id'],), fetch=False)
                log_auditoria(admin_actual, "RECHAZAR_USUARIO", "usuarios", str(p['id']), f"Rechazó a {p['username']}")
                st.warning(f"Usuario {p['username']} rechazado.")
                st.rerun()
    else:
        st.success("No hay solicitudes de acceso pendientes.")

    st.markdown("---")
    st.markdown("#### 👥 Gestión de Usuarios y Delegación de Administradores")
    usuarios_act = run_query("SELECT id, username, nombre_completo, rol, estado, fecha_registro, aprobado_por FROM usuarios WHERE estado != 'pendiente' ORDER BY rol DESC, nombre_completo ASC")
    
    if usuarios_act:
        for u in usuarios_act:
            cu1, cu2, cu3 = st.columns([3, 1.5, 1.5])
            badge_r = "👑 Administrador" if u["rol"] == "admin" else "👤 Asistencial"
            badge_e = "🟢 Aprobado" if u["estado"] == "aprobado" else "🔴 Rechazado"
            cu1.markdown(f"**{u['nombre_completo']}** (`{u['username']}`) · {badge_r} · {badge_e}")
            
            # Botón de delegación de rol
            if u["username"] != st.session_state.get("auth_username", ""):
                if u["rol"] == "usuario":
                    if cu2.button("⭐ Promover a Admin", key=f"promo_{u['id']}"):
                        run_query("UPDATE usuarios SET rol = 'admin' WHERE id = ?", (u['id'],), fetch=False)
                        log_auditoria(st.session_state.get("auth_username", ""), "PROMOVER_ADMIN", "usuarios", str(u['id']), f"Ascendido a admin: {u['username']}")
                        st.success(f"Se ha otorgado rol de Administrador a {u['username']}.")
                        st.rerun()
                else:
                    if cu2.button("⬇️ Degradar a Asistencial", key=f"degrad_{u['id']}"):
                        run_query("UPDATE usuarios SET rol = 'usuario' WHERE id = ?", (u['id'],), fetch=False)
                        log_auditoria(st.session_state.get("auth_username", ""), "DEGRADAR_ADMIN", "usuarios", str(u['id']), f"Degradado a usuario: {u['username']}")
                        st.warning(f"Se cambió a Asistencial a {u['username']}.")
                        st.rerun()
                
                # Desactivar / Activar
                if u["estado"] == "aprobado":
                    if cu3.button("🚫 Desactivar", key=f"desact_{u['id']}"):
                        run_query("UPDATE usuarios SET estado = 'rechazado' WHERE id = ?", (u['id'],), fetch=False)
                        log_auditoria(st.session_state.get("auth_username", ""), "DESACTIVAR_USUARIO", "usuarios", str(u['id']), f"Desactivó a {u['username']}")
                        st.rerun()
                else:
                    if cu3.button("🟢 Reactivar", key=f"react_{u['id']}"):
                        run_query("UPDATE usuarios SET estado = 'aprobado' WHERE id = ?", (u['id'],), fetch=False)
                        log_auditoria(st.session_state.get("auth_username", ""), "REACTIVAR_USUARIO", "usuarios", str(u['id']), f"Reactivó a {u['username']}")
                        st.rerun()

# ----------------- FLUJO PRINCIPAL -----------------
if not check_session():
    login_page()
    st.stop()

current_user = get_current_user()

with st.sidebar:
    st.markdown(
        f"""
        <div style="background-color: #FFFFFF; padding: 14px 18px; border-radius: 14px; margin-bottom: 16px; text-align: center;">
            <img src="{LOGO_SURA_URL}" style="width: 100%; max-width: 140px; height: auto; display: block; margin: 0 auto;">
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; font-size: 0.72rem; color: #00D4FF !important; font-weight: 800; letter-spacing: 1.2px; margin-bottom: 14px;'>SALUD EN CASA (SENC) · BASE AGUDOS</p>",
        unsafe_allow_html=True,
    )

    badge_rol = "👑 Administrador" if is_admin() else "👤 Asistencial"
    st.markdown(
        f"""
        <div style="background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(0, 212, 255, 0.2); border-radius: 10px; padding: 10px 14px; margin-bottom: 16px;">
            <div style="font-size: 0.85rem; font-weight: 700; color: #FFFFFF;">{current_user.get('nombre_completo', 'Usuario')}</div>
            <div style="font-size: 0.72rem; color: #00D4FF; font-weight: 600;">{badge_rol} · @{current_user.get('username')}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("---")
    
    opciones_menu = list(MENU_ITEMS)
    if is_admin():
        opciones_menu.append("Panel Administrador")

    menu = st.radio("MENÚ OPERATIVO", opciones_menu, key="menu_radio_sel")
    st.markdown("---")

    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        logout()

# Encabezado institucional
st.markdown(
    f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Base y Control Agudos</h2>
            <p class="hospital-sub">Programa Salud en Casa · SURA</p>
        </div>
        <div>
            <img src="{LOGO_SURA_URL}" style="height: 44px; width: auto;">
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

# ==================== 1. GESTOR PACIENTES ====================
if menu == "Gestor Pacientes":
    st.subheader("🔍 Localizador de Pacientes")
    col_busq_t, col_busq1, col_busq2, col_busq3 = st.columns([1, 2.2, 1, 1])
    tipo_doc_busq = col_busq_t.selectbox("Tipo Doc", TIPOS_DOC, index=0, key="tipo_doc_busq_sel")

    if "input_doc_busqueda" not in st.session_state:
        st.session_state.input_doc_busqueda = ""

    doc_busqueda = col_busq1.text_input("Documento de Identidad:", key="input_doc_busqueda").strip()

    # FUNCIÓN DE LIMPIEZA TOTAL DE FORMULARIO
    def reset_formulario_paciente():
        st.session_state.input_doc_busqueda = ""
        st.session_state.doc_activo_actual = ""
        st.session_state.id_tto_editando = None
        st.session_state.id_tto_cambio = None
        st.session_state.id_tto_ver = None
        st.session_state.id_tto_suspender = None
        st.session_state.limpiar_form_nuevo = True
        for k in list(st.session_state.keys()):
            if any(x in k for x in ["med_select", "prof_formular", "widget_hora_inicio", "chk_cuid_"]):
                del st.session_state[k]

    with col_busq2:
        st.markdown("<div style='margin-top: 27px;'></div>", unsafe_allow_html=True)
        if st.button("🧹 Limpiar Cédula", use_container_width=True):
            reset_formulario_paciente()
            st.rerun()

    with col_busq3:
        st.markdown("<div style='margin-top: 27px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Nueva Consulta", use_container_width=True):
            reset_formulario_paciente()
            st.rerun()

    if doc_busqueda != st.session_state.get("doc_activo_actual", ""):
        st.session_state.id_tto_editando = None
        st.session_state.id_tto_cambio = None
        st.session_state.id_tto_ver = None
        st.session_state.id_tto_suspender = None
        st.session_state.doc_activo_actual = doc_busqueda
        st.session_state.limpiar_form_nuevo = True

    if doc_busqueda:
        paciente_res = run_query(
            """SELECT documento, tipo_documento, nombre, plan, programa, piso, zona, municipio, barrio, 
                      aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, 
                      usuario_registro, fecha_ingreso FROM pacientes WHERE documento = ?""",
            (doc_busqueda,),
        )

        p_doc, p_t_doc_db, p_nom, p_plan, p_prog, p_piso, p_zona = doc_busqueda, tipo_doc_busq, "", "POS", "Agudos", "Norte", "Zona 1"
        p_muni, p_barrio, p_aisla, p_tipo_aisla, p_obs, p_alerta_fija, p_user, p_fec_ingreso = "Medellín", "", "NO", "Ninguno", "", "", current_user.get("nombre_completo", ""), None

        if paciente_res:
            p_data = paciente_res[0]
            p_doc = p_data["documento"]
            p_t_doc_db = p_data.get("tipo_documento", "CC")
            p_nom = p_data["nombre"]
            p_plan = p_data["plan"]
            p_prog = p_data["programa"]
            p_piso = p_data["piso"]
            p_zona = p_data["zona"]
            p_muni = p_data["municipio"]
            p_barrio = p_data["barrio"]
            p_aisla = p_data["aislamiento"]
            p_tipo_aisla = p_data["tipo_aislamiento"]
            p_obs = p_data["observaciones_clinicas"] or ""
            p_alerta_fija = p_data["alerta_permanente"] or ""
            p_user = p_data["usuario_registro"] or ""
            p_fec_ingreso = p_data["fecha_ingreso"]

            if p_t_doc_db and p_t_doc_db.upper() != tipo_doc_busq.upper():
                st.error(f"🚨 **Inconsistencia de Tipo Documental:** Registrado como **{p_t_doc_db}**, no **{tipo_doc_busq}**.")
                st.stop()

            dias_estancia_calc = 0
            if p_fec_ingreso:
                try:
                    dt_ing = datetime.fromisoformat(str(p_fec_ingreso)[:10])
                    dias_estancia_calc = max(0, (datetime.now() - dt_ing).days)
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
                t_doc_sel = col1.selectbox("Tipo Documento", TIPOS_DOC, index=TIPOS_DOC.index(p_t_doc_db) if p_t_doc_db in TIPOS_DOC else 0, key="edit_tdoc")
                nombre = col2.text_input("Nombre y Apellido", value=p_nom, key="edit_nombre")
                plan = col3.selectbox("Plan", PLANES, index=PLANES.index(p_plan) if p_plan in PLANES else 0, key="edit_plan")

                programa = col1.selectbox("Programa", PROGRAMAS, index=0, key="edit_prog")
                col4, col5, col6, col7 = st.columns(4)
                municipio = col4.selectbox("Municipio", MUNICIPIOS, index=MUNICIPIOS.index(p_muni) if p_muni in MUNICIPIOS else 0, key="edit_muni")
                barrio = col5.text_input("Barrio", value=p_barrio, key="edit_barrio")
                piso = col6.selectbox("Piso", PISOS, index=PISOS.index(p_piso) if p_piso in PISOS else 0, key="edit_piso")
                zona = col7.selectbox("Zona", ZONAS, index=ZONAS.index(p_zona) if p_zona in ZONAS else 0, key="edit_zona")

                st.markdown("#### 📝 Observaciones Clínicas y Alerta Fija Permanente")
                col_obs1, col_obs2 = st.columns(2)
                obs_clinica = col_obs1.text_area("Observación Clínica General:", value=p_obs, key="edit_obs")
                alerta_fija = col_obs2.text_area("📌 Alerta Fija Permanente:", value=p_alerta_fija, key="edit_alerta")

                st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
                c_ais1, c_ais2 = st.columns(2)
                tiene_aislamiento = c_ais1.radio("¿Tiene aislamiento?", ["NO", "SI"], index=0 if p_aisla == "NO" else 1, horizontal=True, key="edit_aisla_radio")
                tipo_aisla_val = "Ninguno"
                if tiene_aislamiento == "SI":
                    idx_t = AISLAMIENTOS_TIPOS.index(p_tipo_aisla) if p_tipo_aisla in AISLAMIENTOS_TIPOS else 0
                    tipo_aisla_val = c_ais2.selectbox("Tipo de Aislamiento:", AISLAMIENTOS_TIPOS, index=idx_t, key="edit_tipo_aisla")

                usuario_edita = render_selectbox("Profesional responsable:", cat_gestores, current_user.get("nombre_completo", p_user), key="edit_user_resp")

                if st.button("Actualizar Ficha Paciente"):
                    if usuario_edita.strip():
                        run_query(
                            """UPDATE pacientes SET tipo_documento=?, nombre=?, plan=?, programa=?, piso=?, zona=?, 
                                       municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, 
                                       observaciones_clinicas=?, alerta_permanente=?, estado_paciente='Activo', usuario_registro=? WHERE documento=?""",
                            (t_doc_sel, nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_edita.strip(), doc_busqueda),
                            fetch=False,
                        )
                        log_auditoria(current_user.get("username", ""), "ACTUALIZAR_PACIENTE", "pacientes", doc_busqueda, f"Actualizó a {nombre}")
                        st.success("Ficha guardada y actualizada correctamente.")
                        st.rerun()
                    else:
                        st.error("Debes ingresar tu nombre como responsable.")
            else:
                st.info("Paciente nuevo. Diligencie los datos de ingreso:")
                t_doc_sel = col1.selectbox("Tipo Documento", TIPOS_DOC, index=0, key="new_tdoc")
                nombre = col2.text_input("Nombre y Apellido", value="", key="new_nombre")
                plan = col3.selectbox("Plan", PLANES, index=0, key="new_plan")

                programa = col1.selectbox("Programa", PROGRAMAS, index=0, key="new_prog")
                col4, col5, col6, col7 = st.columns(4)
                municipio = col4.selectbox("Municipio", MUNICIPIOS, index=0, key="new_muni")
                barrio = col5.text_input("Barrio", value="", key="new_barrio")
                piso = col6.selectbox("Piso", PISOS, index=0, key="new_piso")
                zona = col7.selectbox("Zona", ZONAS, index=0, key="new_zona")

                st.markdown("#### 📝 Observaciones Clínicas y Alerta Fija Permanente")
                col_obs1, col_obs2 = st.columns(2)
                obs_clinica = col_obs1.text_area("Observación Clínica General:", value="", key="new_obs")
                alerta_fija = col_obs2.text_area("📌 Alerta Fija Permanente:", value="", key="new_alerta")

                st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
                c_ais1, c_ais2 = st.columns(2)
                tiene_aislamiento = c_ais1.radio("¿Tiene aislamiento?", ["NO", "SI"], index=0, horizontal=True, key="new_aisla_radio")
                tipo_aisla_val = "Ninguno"
                if tiene_aislamiento == "SI":
                    tipo_aisla_val = c_ais2.selectbox("Tipo de Aislamiento:", AISLAMIENTOS_TIPOS, key="new_tipo_aisla")

                usuario_crea = render_selectbox("Profesional que registra:", cat_gestores, current_user.get("nombre_completo", ""), key="new_user_reg")

                if st.button("Admitir Paciente a la Base"):
                    if nombre and usuario_crea.strip():
                        fec_hoy = datetime.now().date().isoformat()
                        run_query(
                            """INSERT INTO pacientes (documento, tipo_documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso) 
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?, ?)""",
                            (doc_busqueda, t_doc_sel, nombre, plan, programa, piso, zona, municipio, barrio, tiene_aislamiento, tipo_aisla_val, obs_clinica.strip(), alerta_fija.strip(), usuario_crea.strip(), fec_hoy),
                            fetch=False,
                        )
                        log_auditoria(current_user.get("username", ""), "ADMITIR_PACIENTE", "pacientes", doc_busqueda, f"Admitió a {nombre}")
                        st.success("Paciente admitido exitosamente.")
                        st.rerun()
                    else:
                        st.error("Diligencie el nombre completo del paciente y el profesional responsable.")

        if paciente_res:
            st.divider()
            ttos_activos = run_query(
                "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo' ORDER BY id DESC",
                (doc_busqueda,),
            )

            # Acciones de tratamiento
            if st.session_state.id_tto_cambio:
                tc_res = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_cambio,))
                if tc_res:
                    tc = tc_res[0]
                    st.error(f"🔄 **CAMBIO DE TRATAMIENTO MÉDICO**\n\nMedicamento: **{tc['tratamiento']}** ({tc['dosis']}) - Vía: {tc['via_administracion']}")
                    motivo_cambio_med = st.text_input("Motivo médico del cambio (Obligatorio):", value="Rotación antibiótica / Cambio de orden médica")

                    col_cbtn1, col_cbtn2 = st.columns([2, 1])
                    with col_cbtn1:
                        if st.button("Confirmar Cierre y Formular Nuevo"):
                            fecha_cambio_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            nov_cierre = f"[CAMBIO DE TTO el {fecha_cambio_str}]: Se suspende {tc['tratamiento']}. Motivo: {motivo_cambio_med}"
                            run_query("UPDATE tratamientos SET estado = 'Cambiado por Orden Médica', novedades = novedades || ' // ' || ? WHERE id = ?", (nov_cierre, tc["id"]), fetch=False)
                            log_auditoria(current_user.get("username", ""), "CAMBIO_TTO", "tratamientos", str(tc["id"]), nov_cierre)
                            st.session_state.id_tto_cambio = None
                            st.success(f"✅ {tc['tratamiento']} cerrado. Formule el nuevo medicamento abajo.")
                            st.rerun()
                    with col_cbtn2:
                        if st.button("Cancelar Cambio"):
                            st.session_state.id_tto_cambio = None
                            st.rerun()
                    st.divider()

            if st.session_state.id_tto_suspender:
                ts_res = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_suspender,))
                if ts_res:
                    ts = ts_res[0]
                    st.error(f"⛔ **SUSPENDER TRATAMIENTO DEFINITIVAMENTE**\n\nMedicamento: **{ts['tratamiento']}** ({ts['dosis']}) - Vía: {ts['via_administracion']}")
                    col_s1, col_s2 = st.columns(2)
                    motivo_suspension = col_s1.text_input("Motivo de la suspensión médica:", value="Suspensión médica definitiva")
                    prof_susp = render_selectbox("Profesional responsable:", cat_gestores, current_user.get("nombre_completo", ""), key="prof_susp")

                    col_sbtn1, col_sbtn2 = st.columns([2, 1])
                    with col_sbtn1:
                        if st.button("Confirmar Suspensión"):
                            if not prof_susp.strip():
                                st.error("Ingrese profesional responsable.")
                            else:
                                fecha_susp_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                                nov_susp = f"[SUSPENDIDO el {fecha_susp_str} por {prof_susp.strip()}]: {motivo_suspension}"
                                run_query("UPDATE tratamientos SET estado = 'Suspendido Médicamente', novedades = novedades || ' // ' || ?, usuario_modificacion = ? WHERE id = ?", (nov_susp, prof_susp.strip(), ts["id"]), fetch=False)
                                log_auditoria(current_user.get("username", ""), "SUSPENDER_TTO", "tratamientos", str(ts["id"]), nov_susp)
                                st.session_state.id_tto_suspender = None
                                st.success(f"✅ {ts['tratamiento']} suspendido permanentemente.")
                                st.rerun()
                    with col_sbtn2:
                        if st.button("Cancelar"):
                            st.session_state.id_tto_suspender = None
                            st.rerun()
                    st.divider()

            if st.session_state.id_tto_ver:
                tv_res = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_ver,))
                if tv_res:
                    tv = tv_res[0]
                    if st.button("🔙 Regresar"):
                        st.session_state.id_tto_ver = None
                        st.rerun()
                    st.info(f"👁️ **DETALLE COMPLETO TRATAMIENTO #{tv['id']}**")
                    st.write(f"**Medicamento:** {tv['tratamiento']} | **Dosis:** {tv['dosis']} | **Vía:** {tv['via_administracion']} ({tv['tipo_acceso']})")
                    st.write(f"**Frecuencia:** Cada {tv['frecuencia_horas']}h por {tv['dias']} días")
                    st.write(f"**Fin:** {tv['fecha_fin']} | **Retiro Catéter:** {tv['fecha_retiro_cateter'] or 'N/A'}")
                    st.write(f"**Historial de Novedades:** {tv['novedades']}")
                    st.divider()

            st.markdown("### 💊 Tratamientos Activos")
            if ttos_activos:
                for t in ttos_activos:
                    alerta_t = evaluar_alerta_fin(t["fecha_fin"], t["frecuencia_horas"], t["estado"], t["descanalizado"], t["novedades"], t["alerta_revisada"], t["via_administracion"])
                    fin_dt_fmt = datetime.fromisoformat(t["fecha_fin"]).strftime('%d/%m %I:%M %p') if t["fecha_fin"] != "-" else "-"
                    
                    t1_hoy, t2_hoy, t3_hoy = limpiar_texto_turno(t["t1"]), limpiar_texto_turno(t["t2"]), limpiar_texto_turno(t["t3"])

                    st.markdown(
                        f"""
                        <div class="card-tto-activo">
                            <strong style="font-size: 1.05rem; color: #0033A0;">💊 {t['tratamiento']} ({t['dosis']})</strong> · 
                            <span><strong>Vía:</strong> {t['via_administracion']}</span> · 
                            <span><strong>Acceso:</strong> {t['tipo_acceso']}</span> · 
                            <span><strong>Frec:</strong> C/{t['frecuencia_horas']}h</span> · 
                            <span><strong>Fin:</strong> {fin_dt_fmt}</span>
                            <div style="margin-top: 10px; margin-bottom: 8px;">
                                <span class="tag-turno">T1: {t1_hoy}</span>
                                <span class="tag-turno">T2: {t2_hoy}</span>
                                <span class="tag-turno">T3: {t3_hoy}</span>
                            </div>
                            <span style="font-size: 0.85rem; font-weight: 700; color: #475569;">Estado:</span> 
                            <span style="font-size: 0.85rem; font-weight: 700; color: #0033A0;">{alerta_t}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    c_act1, c_act2, c_act3, c_act4 = st.columns(4)
                    if c_act1.button("✏️ Modificar", key=f"btn_edit_{t['id']}"):
                        st.session_state.id_tto_editando = t["id"]
                        st.session_state.id_tto_cambio = None
                        st.session_state.id_tto_ver = None
                        st.session_state.id_tto_suspender = None
                        st.rerun()

                    if c_act2.button("🔄 Cambio TTO", key=f"btn_cambio_{t['id']}"):
                        st.session_state.id_tto_cambio = t["id"]
                        st.session_state.id_tto_editando = None
                        st.session_state.id_tto_ver = None
                        st.session_state.id_tto_suspender = None
                        st.rerun()

                    if c_act3.button("⛔ Suspender", key=f"btn_susp_{t['id']}"):
                        st.session_state.id_tto_suspender = t["id"]
                        st.session_state.id_tto_editando = None
                        st.session_state.id_tto_cambio = None
                        st.session_state.id_tto_ver = None
                        st.rerun()

                    if c_act4.button("👁️ Ver Historial", key=f"btn_ver_{t['id']}"):
                        st.session_state.id_tto_ver = t["id"]
                        st.rerun()
            else:
                st.info("El paciente no tiene tratamientos activos formulados.")

            tto_edit_data = None
            if st.session_state.id_tto_editando:
                tto_edit_res = run_query("SELECT * FROM tratamientos WHERE id = ?", (st.session_state.id_tto_editando,))
                if tto_edit_res:
                    tto_edit_data = tto_edit_res[0]
                    st.warning(f"✏️ **Modificando Tratamiento Activo ID #{tto_edit_data['id']}:** {tto_edit_data['tratamiento']}")

            if st.session_state.get('limpiar_form_nuevo', False):
                tto_edit_data = None
                st.session_state.limpiar_form_nuevo = False

            # Formulario de formulación
            st.markdown("### ➕ Formular Nuevo Tratamiento o Guardar Cambios")
            val_via = tto_edit_data["via_administracion"] if tto_edit_data else "IV"
            col_v, col_a = st.columns([1, 1])
            t_via = col_v.selectbox("Vía de Administración:", VIAS, index=VIAS.index(val_via) if val_via in VIAS else 0)

            val_acc = tto_edit_data["tipo_acceso"] if tto_edit_data else "Periférico"
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
                val_acc_idx = ACCESOS.index(val_acc) if val_acc in ACCESOS else 0
                t_acceso = col_a.selectbox("Acceso Vascular:", ACCESOS, index=val_acc_idx)

            visita_educacion_sc = ""
            npt_dias_semana_json = "[]"
            npt_horas_infusion = 12
            npt_hora_desconexion_str = ""

            col_med, col_dos = st.columns([2.2, 1.2])
            val_med = tto_edit_data["tratamiento"] if tto_edit_data else ""
            val_dos = tto_edit_data["dosis"] if tto_edit_data else ""

            t_nombre = ""
            bloquear_name = False

            if t_via == "BE analgesia":
                t_nombre, bloquear_name, val_dos = "BOMBA DE ANALGESIA / VIGILANCIA", True, "1 Bomba"
            elif t_via == "NPT":
                t_nombre, bloquear_name = "NUTRICIÓN PARENTERAL TOTAL (NPT)", True
                if not val_dos: val_dos = "1 Bolsa"
            elif t_via == "Materna":
                t_nombre, bloquear_name, val_dos = "PROGRAMA MATERNO / EDUCACIÓN", True, "N/A"
            elif t_via == "Tuberculosis":
                t_nombre, bloquear_name, val_dos = "TRATAMIENTO TUBERCULOSIS (TB)", True, "1 Dosis Diaria"
            elif t_via == "LEV":
                t_nombre, bloquear_name = "LÍQUIDOS ENDOVENOSOS (LEV)", True
                if not val_dos: val_dos = "1 Bolsa / 1000cc"
            elif t_via == "BE":
                lista_meds = cat_meds_be
            else:
                lista_meds = cat_meds_agudos

            with col_med:
                if bloquear_name:
                    t_nombre = st.text_input("Medicamento / Terapia:", value=t_nombre, disabled=True)
                else:
                    t_nombre = render_selectbox("Medicamento / Terapia:", lista_meds, val_med, key="med_select")

            with col_dos:
                if t_via == "BE":
                    match_d = re.search(r'([\d\.,]+\s*(?:mg|gr|g))', t_nombre, re.IGNORECASE)
                    t_dosis = match_d.group(0) if match_d else (val_dos if val_dos else "Automática")
                    st.text_input("Dosis (Automática):", value=t_dosis, disabled=True)
                elif t_via in ["BE analgesia", "Tuberculosis", "Materna", "NPT", "LEV"]:
                    t_dosis = st.text_input("Dosis / Volumen:", value=val_dos, disabled=True if t_via in ["NPT", "LEV"] else False)
                else:
                    t_dosis = st.text_input("Dosis (Ej: 1g, 500mg, 1 Bolsa):", value=val_dos)

            col_f1, col_f2 = st.columns(2)
            val_frec = safe_int(tto_edit_data["frecuencia_horas"], 8) if tto_edit_data else 8
            val_dias = safe_int(tto_edit_data["dias"], 5) if tto_edit_data else 5

            if t_via == "BE":
                frecuencia = col_f1.selectbox("Frecuencia Unidosis Puente (Horas):", [6, 8], index=0 if val_frec == 6 else 1)
                dias = col_f2.number_input("Días Ordenados:", min_value=1, max_value=999, value=val_dias)
            elif t_via == "BE analgesia":
                frecuencia, dias = 12, 3
                col_f1.info("Frecuencia: **Cada 12 horas (Fijo)**")
                col_f2.info("Días Ordenados: **3 Días (Fijo)**")
            elif t_via in ["Tuberculosis", "Materna"]:
                frecuencia = 24
                dias = col_f2.number_input("Días Ordenados:", min_value=1, max_value=999, value=val_dias)
                col_f1.info(f"Frecuencia: **Cada 24 horas (Diaria para {t_via})**")
            elif t_via == "NPT":
                frecuencia = 24
                dias = col_f2.number_input("Días Ordenados:", min_value=1, max_value=999, value=val_dias)
                col_f1.info("Frecuencia NPT: **Protocolo diario/semanal**")
            elif t_via == "LEV":
                frecuencia = col_f1.selectbox("Duración de Infusión LEV (Horas):", [12, 24], index=1 if val_frec == 24 else 0)
                dias = col_f2.number_input("Días Ordenados:", min_value=1, max_value=999, value=val_dias)
            else:
                def_frec_idx = FRECUENCIAS.index(val_frec) if val_frec in FRECUENCIAS else 2
                frecuencia = col_f1.selectbox("Frecuencia (Horas):", FRECUENCIAS, index=def_frec_idx)
                dias = col_f2.number_input("Días Ordenados:", min_value=1, max_value=999, value=val_dias)

            tipo_visita_sc_nbz = "Visitar por Horario"
            if t_via in ["SC", "NBZ"]:
                st.markdown(f"#### 🩺 Modalidad de Intervención ({t_via})")
                val_edu_prev = tto_edit_data.get("visita_educacion", "Visitar por Horario") if tto_edit_data else "Visitar por Horario"
                idx_edu = 1 if val_edu_prev == "Solo Educación" else 0
                tipo_visita_sc_nbz = st.radio(
                    f"Modalidad para {t_via}:",
                    ["Visitar por Horario", "Solo Educación"],
                    index=idx_edu,
                    horizontal=True
                )
                visita_educacion_sc = tipo_visita_sc_nbz

            mapa_admin = {}
            if t_via in ["IV", "PUSH", "SC", "NBZ", "LEV"] and tipo_visita_sc_nbz != "Solo Educación":
                st.markdown("#### 👥 Distribución de Administración (SENC vs Cuidador)")
                dt_temp_ini = datetime.combine(datetime.now().date(), datetime.now().time())
                _, formato_horas_temp = obtener_horas_ciclo(dt_temp_ini, frecuencia if t_via != "LEV" else 24)
                
                try:
                    mapa_admin = json.loads(tto_edit_data.get("distribucion_admin", "{}")) if tto_edit_data else {}
                except Exception:
                    mapa_admin = {}

                cols_c = st.columns(min(len(formato_horas_temp), 4))
                for idx_h, h_str in enumerate(formato_horas_temp):
                    col_actual = cols_c[idx_h % len(cols_c)]
                    actual_resp = mapa_admin.get(h_str, "SENC")
                    chk_cuidador = col_actual.checkbox(f"Dosis {h_str}: Cuidador", value=(actual_resp == "Cuidador"), key=f"chk_cuid_{h_str}")
                    mapa_admin[h_str] = "Cuidador" if chk_cuidador else "SENC"

            if t_via == "NPT":
                st.markdown("#### 🥣 Protocolo NPT")
                col_npt1, col_npt2 = st.columns(2)
                dias_def_npt = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                if tto_edit_data and tto_edit_data.get("npt_dias_semana"):
                    try:
                        dias_def_npt = json.loads(tto_edit_data["npt_dias_semana"])
                    except Exception:
                        pass
                
                dias_semana_npt = col_npt1.multiselect("Días programados en semana:", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], default=dias_def_npt)
                npt_dias_semana_json = json.dumps(dias_semana_npt)
                horas_inf_opts = [12, 14, 16, 18, 20, 24]
                h_inf_def = safe_int(tto_edit_data.get("npt_horas_infusion", 12), 12) if tto_edit_data else 12
                npt_horas_infusion = col_npt2.selectbox("Duración de infusión (Horas):", horas_inf_opts, index=horas_inf_opts.index(h_inf_def) if h_inf_def in horas_inf_opts else 0)

            val_ini = datetime.fromisoformat(tto_edit_data["fecha_inicio"]) if tto_edit_data else datetime.now()
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
                st.info(f"🥣 **Horario NPT Calculado:** Instalación a las {dt_inicio.strftime('%I:%M %p')} -> Retiro a las {dt_desconex.strftime('%I:%M %p')}")

            unidosis_puente, bombas_calc, dosis_final_cierre = 0, 0, 0
            veredicto_cierre_be = ""
            cronograma_visitas = []

            if t_via == "BE":
                st.markdown(f"#### 🧪 Protocolo Clínico BE C/{frecuencia}h")
                col_be1, col_be2 = st.columns(2)
                val_puente_prev = safe_int(tto_edit_data.get("be_unidosis_puente", 0), 0) if tto_edit_data else 0
                unidosis_puente = col_be1.number_input("Unidosis IV de puente previo (nocturna 23:00 - 05:00):", min_value=0, max_value=50, value=val_puente_prev)

                total_unidosis_ord, bombas_calc, dosis_final_cierre, _ = calcular_esquema_be_exacto(dias, frecuencia, unidosis_puente)
                veredicto_cierre_be = "✅ El ciclo CONCLUYE CON BE COMPLETA." if dosis_final_cierre == 0 else f"⚠️ Concluye en UNIDOSIS DE CIERRE ({dosis_final_cierre} dosis finales)."

                col_be2.markdown(
                    f"""
                    <div style='background-color: #D1E7DD; color: #000000; padding: 12px; border-radius: 8px; font-size: 0.85rem;'>
                        <strong>📦 BALANCE BE EXACTO:</strong><br>
                        • Días: <strong>{dias} Días</strong> = <strong>{bombas_calc} Bombas Completas</strong>.<br>
                        • Veredicto: <strong>{veredicto_cierre_be}</strong>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            col_aj1, col_aj2, col_aj3 = st.columns([1, 1, 2])
            tipo_unidad_ajuste = col_aj1.selectbox("Unidad de Ajuste:", ["Unidosis", "BE Completa"], index=0 if "BE" not in t_via else 1)
            val_ajuste_prev = safe_int(tto_edit_data.get("dosis_perdidas", 0), 0) if tto_edit_data else 0
            cant_ajuste = col_aj2.number_input("Cantidad (+ / -):", min_value=-99, max_value=99, value=val_ajuste_prev)

            dosis_equiv_ajuste = cant_ajuste
            if tipo_unidad_ajuste == "BE Completa":
                dosis_por_be_adj = 4 if frecuencia == 6 else 3
                dosis_equiv_ajuste = cant_ajuste * dosis_por_be_adj

            motivo_ajuste = col_aj3.selectbox("Motivo del Ajuste:", ["Ninguno / Exacto", "Aumento ordenado", "Extensión de días", "Paciente Ausente", "Suspensión anticipada"])
            requiere_retiro_cateter = t_via in ["IV", "BE", "BE analgesia", "PUSH", "LEV"] or (t_via == "SC" and t_acceso == "SC")

            # Cálculo de Cronograma y Fechas Clínicas
            if t_via in ["IV", "PUSH"]:
                (fecha_fin_calculada, t1_calc, t2_calc, t3_calc, _, _, _, _) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
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
                                "Secuencia": f"Retiro 2h ({h_str})",
                                "Fecha y Hora Programada": dt_retiro_2h.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": "⚠️ Retiro de infusión / Flush a las 2 horas"
                            })
                        curr_dt_c += timedelta(hours=frecuencia)

            elif t_via == "SC":
                if tipo_visita_sc_nbz == "Solo Educación":
                    fecha_fin_calculada, fecha_retiro_cateter_calc = dt_inicio, None
                    t1_calc, t2_calc, t3_calc = f"EDUCACIÓN ({dt_inicio.strftime('%H:%M')})", "—", "—"
                    cronograma_visitas.append({
                        "Secuencia": "Visita Única Educación SC",
                        "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": "Instrucción y educación en vía Subcutánea"
                    })
                else:
                    (fecha_fin_calculada, t1_calc, t2_calc, t3_calc, _, _, _, _) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
                    fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                    _, formato_c = obtener_horas_ciclo(dt_inicio, frecuencia)
                    curr_dt_c = dt_inicio
                    for d_idx in range(dias):
                        for h_str in formato_c:
                            resp = mapa_admin.get(h_str, "SENC")
                            cronograma_visitas.append({
                                "Secuencia": f"Dosis {h_str} (Día {d_idx+1})",
                                "Fecha y Hora Programada": curr_dt_c.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": f"Administración vía SC por: {resp}"
                            })
                            curr_dt_c += timedelta(hours=frecuencia)

            elif t_via == "BE":
                dt_cursor = dt_inicio
                es_nocturno = (dt_inicio.hour >= 23 or dt_inicio.hour < 5) and unidosis_puente > 0
                for p_idx in range(1, unidosis_puente + 1):
                    cronograma_visitas.append({
                        "Secuencia": f"Puente #{p_idx}",
                        "Fecha y Hora Programada": dt_cursor.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Unidosis Puente IV (C/{frecuencia}h)"
                    })
                    dt_cursor += timedelta(hours=frecuencia)

                dt_bomba_inicio = dt_cursor if es_nocturno else dt_inicio
                dias_be_efectivos = max(1, dias + (cant_ajuste if tipo_unidad_ajuste == "BE Completa" else 0))
                _, bombas_calc_ajustadas, dosis_final_cierre_ajustada, _ = calcular_esquema_be_exacto(dias_be_efectivos, frecuencia, unidosis_puente)

                horario_visita_be = f"{dt_bomba_inicio.strftime('%H:%M')}H (RECAMBIO BE C/24H)"
                t1_calc = horario_visita_be if 6 <= dt_bomba_inicio.hour < 14 else "—"
                t2_calc = horario_visita_be if 14 <= dt_bomba_inicio.hour < 22 else "—"
                t3_calc = horario_visita_be if (dt_bomba_inicio.hour >= 22 or dt_bomba_inicio.hour < 6) else "—"

                dt_b_iter = dt_bomba_inicio
                for b_num in range(1, bombas_calc_ajustadas + 1):
                    cronograma_visitas.append({
                        "Secuencia": f"Bomba BE #{b_num}",
                        "Fecha y Hora Programada": dt_b_iter.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": f"Instalación / Recambio Bomba BE #{b_num}"
                    })
                    dt_b_iter += timedelta(hours=24)

                fecha_fin_calculada = dt_b_iter - timedelta(hours=24)
                if dosis_final_cierre_ajustada > 0:
                    dt_cierre = dt_b_iter - timedelta(hours=24)
                    for c_num in range(1, dosis_final_cierre_ajustada + 1):
                        es_ultima = (c_num == dosis_final_cierre_ajustada)
                        txt_acc = f"Unidosis Cierre #{c_num} (8h)" + (" + ⚠️ RETIRO CATÉTER" if es_ultima else "")
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
                for v_idx in range(1, 7):
                    es_ultima_ba = (v_idx == 6)
                    txt_ba = f"Control BE Analgesia #{v_idx}" + (" + ⚠️ RETIRO CATÉTER Y EGRESO" if es_ultima_ba else "")
                    cronograma_visitas.append({
                        "Secuencia": f"Control BE Analgesia #{v_idx}",
                        "Fecha y Hora Programada": dt_cursor_ba.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": txt_ba
                    })
                    dt_cursor_ba += timedelta(hours=12)
                fecha_fin_calculada = dt_cursor_ba - timedelta(hours=12)
                fecha_retiro_cateter_calc = fecha_fin_calculada
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
                        "Acción Clínica / Intervención en Domicilio": f"Seguimiento {t_via} (Día {m_d} de {dias})"
                    })
                    dt_m_curr += timedelta(hours=24)

            elif t_via == "NBZ":
                if tipo_visita_sc_nbz == "Solo Educación":
                    fecha_fin_calculada, fecha_retiro_cateter_calc = dt_inicio, None
                    t1_calc, t2_calc, t3_calc = f"EDUCACIÓN ({dt_inicio.strftime('%H:%M')})", "—", "—"
                    cronograma_visitas.append({
                        "Secuencia": "Visita Única Educación NBZ",
                        "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                        "Acción Clínica / Intervención en Domicilio": "Educación en técnica de Nebulización"
                    })
                else:
                    (fecha_fin_calculada, t1_calc, t2_calc, t3_calc, _, _, _, _) = calcular_tratamiento_mixto(dt_inicio, dias, frecuencia, mapa_admin, dosis_equiv_ajuste)
                    fecha_retiro_cateter_calc = None
                    _, formato_c = obtener_horas_ciclo(dt_inicio, frecuencia)
                    curr_nbz = dt_inicio
                    for d_idx in range(dias):
                        for h_str in formato_c:
                            resp = mapa_admin.get(h_str, "SENC")
                            cronograma_visitas.append({
                                "Secuencia": f"Nebulización {h_str} (Día {d_idx+1})",
                                "Fecha y Hora Programada": curr_nbz.strftime('%d/%m/%Y %I:%M %p'),
                                "Acción Clínica / Intervención en Domicilio": f"Nebulización C/{frecuencia}h por: {resp}"
                            })
                            curr_nbz += timedelta(hours=frecuencia)

            elif t_via == "NPT":
                fecha_fin_calculada, fecha_retiro_cateter_calc = dt_inicio + timedelta(days=dias), None
                t1_calc, t2_calc, t3_calc = f"INSTALAR ({dt_inicio.strftime('%H:%M')})", f"RETIRAR ({npt_hora_desconexion_str})", "—"
                dias_activos_npt = json.loads(tto_edit_data.get("npt_dias_semana") if (tto_edit_data and tto_edit_data.get("npt_dias_semana")) else npt_dias_semana_json)
                nombres_dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                curr_npt = dt_inicio
                for d_i in range(dias):
                    dia_str = nombres_dias[curr_npt.weekday()]
                    if dia_str in dias_activos_npt:
                        dt_retir = curr_npt + timedelta(hours=npt_horas_infusion)
                        cronograma_visitas.append({
                            "Secuencia": f"NPT Día {d_i+1} ({dia_str})",
                            "Fecha y Hora Programada": curr_npt.strftime('%d/%m/%Y %I:%M %p'),
                            "Acción Clínica / Intervención en Domicilio": f"Instalación NPT -> Retiro a las {dt_retir.strftime('%I:%M %p')}"
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
                        "Acción Clínica / Intervención en Domicilio": f"Instalación LEV ({frecuencia}h) -> Cambio a las {dt_ret_lev.strftime('%I:%M %p')}"
                    })
                    curr_lev += timedelta(hours=24)
            else:
                fecha_fin_calculada = dt_inicio + timedelta(days=dias)
                fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
                t1_calc, t2_calc, t3_calc = "—", "—", "—"
                cronograma_visitas.append({
                    "Secuencia": "Tratamiento Estándar",
                    "Fecha y Hora Programada": dt_inicio.strftime('%d/%m/%Y %I:%M %p'),
                    "Acción Clínica / Intervención en Domicilio": f"Inicio vía {t_via} por {dias} días"
                })

            usuario_responsable = render_selectbox("✍️ Profesional responsable:", cat_gestores, current_user.get("nombre_completo", ""), key="prof_formular")

            st.markdown("#### ⚖️ Resumen de Fechas Clínicas")
            mb1, mb2, mb3, mb4 = st.columns(4)
            mb1.metric("Días Ordenados", f"{dias} Días")
            mb2.metric("Ajuste", f"{cant_ajuste} ({tipo_unidad_ajuste})" if cant_ajuste != 0 else "0")
            mb3.metric("Fecha Fin", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))
            mb4.metric("Retiro Catéter", fecha_retiro_cateter_calc.strftime("%d/%m %I:%M %p") if fecha_retiro_cateter_calc else "N/A")

            st.markdown("#### 📅 Cronograma de Visitas")
            df_cronograma = pd.DataFrame(cronograma_visitas)
            st.dataframe(df_cronograma, use_container_width=True, hide_index=True)

            col_btn_guardar, col_btn_cancelar = st.columns([3, 1])

            with col_btn_guardar:
                btn_label = "💾 Actualizar Cambios" if tto_edit_data else "💾 Guardar Formulación"
                if st.button(btn_label):
                    if not usuario_responsable.strip():
                        st.error("❌ Indique el profesional responsable.")
                    elif not t_nombre.strip():
                        st.error("❌ Ingrese el medicamento.")
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
                                    nov_txt, usuario_responsable.strip(), tto_edit_data["id"],
                                ),
                                fetch=False,
                            )
                            log_auditoria(current_user.get("username", ""), "EDITAR_TTO", "tratamientos", str(tto_edit_data["id"]), nov_txt)
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
                            log_auditoria(current_user.get("username", ""), "CREAR_TTO", "tratamientos", doc_busqueda, nov_txt)
                            st.success("✅ Nuevo tratamiento formulado exitosamente.")
                            st.session_state.limpiar_form_nuevo = True

                        st.rerun()

            with col_btn_cancelar:
                if st.button("❌ Cancelar"):
                    st.session_state.id_tto_editando = None
                    st.rerun()

# ==================== 2. ENTREGA DE TURNO ====================
elif menu == "Entrega de turno":
    st.subheader("📋 Módulo de Entrega de Turnos Operativos")
    col_z_f1, col_z_f2 = st.columns(2)
    zona_filtro_turno = col_z_f1.selectbox("Filtrar por Zona:", ["Todas"] + ZONAS)
    ver_historial_turnos = col_z_f2.checkbox("📂 Ver Historial de Turnos Resueltos")

    with st.expander("➕ Registrar Novedad para el Turno", expanded=False):
        ced_turno_in = st.text_input("Cédula del Paciente:").strip()
        p_piso_t, p_zona_t = "Norte", "Zona 1"
        if ced_turno_in:
            res_pt = run_query("SELECT nombre, piso, zona FROM pacientes WHERE documento = ?", (ced_turno_in,))
            if res_pt:
                pt_data = res_pt[0]
                p_piso_t, p_zona_t = pt_data["piso"], pt_data["zona"]
                st.success(f"👤 Paciente: **{pt_data['nombre']}** | Piso: {p_piso_t} | Zona sugerida: {p_zona_t}")
            else:
                st.warning("⚠️ Cédula no registrada en el censo.")

        c_t1, c_t2, c_t3 = st.columns(3)
        zona_asignada = c_t1.selectbox("Zona Asignada:", ZONAS, index=ZONAS.index(p_zona_t) if p_zona_t in ZONAS else 0)
        prioridad_turno = c_t2.selectbox("Prioridad:", ["Alta 🔴", "Media 🟡", "Baja 🔵"], index=1)
        gestor_turno = render_selectbox("Gestor que reporta:", cat_gestores, current_user.get("nombre_completo", ""), key="gestor_turno_sel")
        detalle_turno_txt = st.text_area("Detalle de la novedad de turno:")

        if st.button("💾 Guardar en Minuta"):
            if not ced_turno_in or not detalle_turno_txt.strip() or not gestor_turno.strip():
                st.error("Diligencie cédula, detalle y gestor.")
            else:
                fec_reg_t = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                run_query(
                    """INSERT INTO entrega_turnos (documento, zona, prioridad, detalle_novedad, estado, registrado_por, fecha_creacion)
                       VALUES (?, ?, ?, ?, 'Activa', ?, ?)""",
                    (ced_turno_in, zona_asignada, prioridad_turno, detalle_turno_txt.strip(), gestor_turno.strip(), fec_reg_t),
                    fetch=False
                )
                log_auditoria(current_user.get("username", ""), "CREAR_ENTREGA_TURNO", "entrega_turnos", ced_turno_in, f"{zona_asignada} - {prioridad_turno}")
                st.success("✅ Novedad de turno guardada.")
                st.rerun()

    st.markdown("---")
    if ver_historial_turnos:
        st.markdown("#### 📂 Historial de Novedades Resueltas")
        data_hist = run_query("SELECT * FROM entrega_turnos WHERE estado = 'Resuelta' ORDER BY id DESC")
        df_hist = pd.DataFrame(data_hist) if data_hist else pd.DataFrame()
        if not df_hist.empty:
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
        else:
            st.info("No hay historial de turnos resueltos.")
    else:
        st.markdown("#### 📌 Novedades Activas de Turno")
        q_act = "SELECT * FROM entrega_turnos WHERE estado = 'Activa'"
        params_act = []
        if zona_filtro_turno != "Todas":
            q_act += " AND zona = ?"
            params_act.append(zona_filtro_turno)
        q_act += " ORDER BY id DESC"
        data_act = run_query(q_act, tuple(params_act))
        df_act = pd.DataFrame(data_act) if data_act else pd.DataFrame()

        if not df_act.empty:
            for _, r_t in df_act.iterrows():
                t_id, t_doc, t_zona, t_prio, t_det, t_reg, t_fec = r_t['id'], r_t['documento'], r_t['zona'], r_t['prioridad'], r_t['detalle_novedad'], r_t['registrado_por'], r_t['fecha_creacion']
                p_nom_busq = t_doc
                res_nom = run_query("SELECT nombre FROM pacientes WHERE documento = ?", (t_doc,))
                if res_nom:
                    p_nom_busq = f"{res_nom[0]['nombre']} (CC {t_doc})"

                with st.expander(f"{t_prio} | [{t_zona}] · {p_nom_busq} - {t_fec}", expanded=True):
                    st.write(f"**Detalle:** {t_det}")
                    st.write(f"*Reportó: {t_reg}*")
                    c_res1, c_res2 = st.columns([2, 1])
                    nombre_gestor_res = render_selectbox("Gestor que resuelve:", cat_gestores, current_user.get("nombre_completo", ""), key=f"res_gestor_{t_id}")
                    if c_res2.button(f"✅ Resolver (#{t_id})", key=f"btn_res_t_{t_id}"):
                        if not nombre_gestor_res.strip():
                            st.error("Seleccione gestor.")
                        else:
                            fec_res_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            run_query(
                                "UPDATE entrega_turnos SET estado = 'Resuelta', fecha_resolucion = ?, resuelto_por = ? WHERE id = ?",
                                (fec_res_str, nombre_gestor_res.strip(), t_id),
                                fetch=False
                            )
                            log_auditoria(current_user.get("username", ""), "RESOLVER_ENTREGA_TURNO", "entrega_turnos", str(t_id), f"Por {nombre_gestor_res.strip()}")
                            st.success("✅ Novedad resuelta.")
                            st.rerun()
        else:
            st.info(f"No hay novedades activas para {zona_filtro_turno}.")

# ==================== 3. NOVEDADES PROGRAMACIÓN ====================
elif menu == "Novedades programacion":
    st.subheader("📢 Gestión Operativa de Novedades Clínicas")
    query_novs_operativas = """
        SELECT 
            rn.id AS 'ID', rn.documento AS 'Cédula', p.nombre AS 'Nombre y Apellido',
            p.piso AS 'Piso', p.zona AS 'Zona', rn.tipo_novedad AS 'Tipo Novedad',
            rn.nivel_novedad AS 'Nivel / Tipo de Acción', rn.fecha_aplicacion AS 'Fecha Aplica',
            rn.hora_aplicacion AS 'Hora Aplica', rn.detalle AS 'Detalle Novedad',
            rn.responsable AS 'Registrado Por'
        FROM registro_novedades rn
        JOIN pacientes p ON rn.documento = p.documento
        WHERE rn.estado_novedad = 'Pendiente'
        ORDER BY rn.id DESC
    """
    data_novs = run_query(query_novs_operativas)
    df_novs = pd.DataFrame(data_novs) if data_novs else pd.DataFrame()
    ahora = datetime.now()

    def clasificar_urgencia(row):
        f_str, h_str = str(row['Fecha Aplica']).strip(), str(row['Hora Aplica']).strip()
        try:
            if len(f_str) == 10 and "/" in f_str:
                d, m, y = f_str.split("/")
                dt_aplica = datetime(int(y), int(m), int(d))
            else:
                dt_aplica = datetime.fromisoformat(f_str[:10])

            if ":" in h_str:
                partes = h_str.split(":")
                hh, mm = int(partes[0].strip()[-2:]), int(partes[1].strip()[:2])
                if "PM" in h_str.upper() and hh < 12: hh += 12
                if "AM" in h_str.upper() and hh == 12: hh = 0
                dt_aplica = dt_aplica.replace(hour=hh, minute=mm)
            else:
                dt_aplica = dt_aplica.replace(hour=8, minute=0)

            diff_horas = (dt_aplica - ahora).total_seconds() / 3600.0
            if diff_horas < 7: return "🔴 CRÍTICA / EN TURNO", diff_horas
            elif 7 <= diff_horas <= 15: return "🟡 ALERTA PRÓXIMA (8-15H)", diff_horas
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
        c_m2.metric("🔴 Críticas (<7h)", f"{total_criticas}")
        c_m3.metric("🟡 Próximas (8-15h)", f"{total_proximas}")
        c_m4.metric("🔵 Programadas (>15h)", f"{total_prog}")

        st.markdown("---")
        f1, f2, f3 = st.columns([1.5, 1.5, 2])
        pisos_disp = sorted(list(df_novs['Piso'].dropna().unique()))
        zonas_disp = sorted(list(df_novs['Zona'].dropna().unique()))
        piso_filtro = f1.multiselect("Piso:", pisos_disp, default=pisos_disp)
        zona_filtro = f2.multiselect("Zona:", zonas_disp, default=zonas_disp)
        semaforo_filtro = f3.multiselect(
            "Prioridad:",
            ["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (8-15H)", "🔵 PROGRAMADA"],
            default=["🔴 CRÍTICA / EN TURNO", "🟡 ALERTA PRÓXIMA (8-15H)", "🔵 PROGRAMADA"]
        )

        df_filtradas = df_novs[
            (df_novs['Piso'].isin(piso_filtro)) &
            (df_novs['Zona'].isin(zona_filtro)) &
            (df_novs['Semaforo'].isin(semaforo_filtro))
        ]

        st.markdown(f"#### 🛎️ Novedades a Gestionar ({len(df_filtradas)})")
        for _, nov in df_filtradas.iterrows():
            n_id, n_doc, n_nom, n_piso, n_zona = nov['ID'], nov['Cédula'], nov['Nombre y Apellido'], nov['Piso'], nov['Zona']
            n_tipo, n_sem, n_fec, n_hor, n_det, n_resp = nov['Tipo Novedad'], nov['Semaforo'], nov['Fecha Aplica'], nov['Hora Aplica'], nov['Detalle Novedad'], nov['Registrado Por']

            with st.expander(f"{n_sem} | #{n_id} · {n_nom} (CC {n_doc}) - {n_piso}/{n_zona}", expanded=("🔴" in n_sem or "🟡" in n_sem)):
                st.markdown(f"**Categoría:** {n_tipo} | **Fecha:** {n_fec} {n_hor} | *Registró: {n_resp}*")
                st.info(f"📝 {n_det}")

                c_g1, c_g2 = st.columns([2, 2])
                prof_ges = render_selectbox("Gestor:", cat_gestores, current_user.get("nombre_completo", ""), key=f"prof_gestion_{n_id}")
                acc_tomada = c_g2.text_input("Nota:", value="Novedad atendida y programada en ruta.", key=f"acc_{n_id}")

                if st.button(f"✅ Resolver (#{n_id})", key=f"btn_res_directo_{n_id}"):
                    if not prof_ges.strip():
                        st.error("Seleccione gestor.")
                    else:
                        fec_ahora_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                        run_query(
                            """UPDATE registro_novedades SET estado_novedad = 'Gestionada', 
                                       fecha_gestion = ?, responsable_gestion = ? WHERE id = ?""",
                            (fec_ahora_str, prof_ges.strip(), n_id),
                            fetch=False,
                        )
                        restantes = run_query("SELECT COUNT(*) as cant FROM registro_novedades WHERE documento = ? AND estado_novedad = 'Pendiente'", (n_doc,))
                        if restantes and restantes[0]["cant"] == 0:
                            run_query(
                                """UPDATE tratamientos SET alerta_revisada = 'SI', 
                                           novedades = novedades || ' // [GESTIONADA ' || ? || ' por ' || ? || ']' WHERE documento = ? AND estado = 'Activo'""",
                                (fec_ahora_str, prof_ges.strip(), n_doc),
                                fetch=False,
                            )
                        log_auditoria(current_user.get("username", ""), "RESOLVER_NOVEDAD", "registro_novedades", str(n_id), f"{prof_ges.strip()} - {acc_tomada}")
                        st.success(f"✅ Novedad #{n_id} gestionada.")
                        st.rerun()
    else:
        st.success("🎉 No hay novedades pendientes en este momento.")

    st.markdown("---")
    with st.expander("➕ Registrar Nueva Novedad a Paciente"):
        doc_nov_ingresada = st.text_input("Cédula Paciente:", key="input_cedula_nov").strip()
        if doc_nov_ingresada:
            p_data = run_query("SELECT documento, nombre, piso, zona FROM pacientes WHERE documento = ?", (doc_nov_ingresada,))
            if p_data:
                p_item = p_data[0]
                p_doc, p_nom, p_piso, p_zona = p_item["documento"], p_item["nombre"], p_item["piso"], p_item["zona"]
                tto_act = run_query("SELECT tratamiento, dosis, via_administracion, fecha_fin FROM tratamientos WHERE documento = ? AND estado = 'Activo'", (p_doc,))
                
                st.success(f"👤 Paciente: **{p_nom}** | Piso: {p_piso} | Zona: {p_zona}")
                if tto_act:
                    st.write(f"💊 Tratamiento Activo: {tto_act[0]['tratamiento']} ({tto_act[0]['dosis']}) - Fin: {tto_act[0]['fecha_fin']}")

                c_r1, c_r2 = st.columns(2)
                tipo_nov_reg = c_r1.selectbox("Tipo:", ["Cambio de Horario", "Visita Fallida", "Descanalización / Salida", "Modificación Orden Médica", "Paciente Ausente", "Otra"], key="reg_tipo")
                nivel_nov_reg = c_r2.selectbox("Nivel:", ["Nivel 1: Informativa", "Nivel 2: Visita Puntual", "Nivel 3: Cambio Permanente / Crítica"], key="reg_nivel")

                c_f1, c_f2 = st.columns(2)
                fec_aplica_reg = c_f1.date_input("Fecha Aplica:", value=datetime.now().date(), key="reg_fec")
                hor_aplica_reg = c_f2.time_input("Hora Aplica:", value=datetime.now().time(), key="reg_hor")

                det_reg = st.text_area("Detalle de la novedad:", key="reg_det")
                prof_reg = render_selectbox("Profesional que registra:", cat_gestores, current_user.get("nombre_completo", ""), key="reg_prof")

                if st.button("💾 Guardar Novedad"):
                    if not prof_reg.strip() or not det_reg.strip():
                        st.error("Diligencie todos los campos.")
                    else:
                        nivel_txt_puro = nivel_nov_reg.split(":")[0].strip()
                        estado_ini = 'Pendiente' if "Nivel 1" in nivel_nov_reg else 'Gestionada'
                        fec_ges_aut = datetime.now().strftime("%d/%m/%Y %I:%M %p") if estado_ini == 'Gestionada' else None

                        run_query(
                            """INSERT INTO registro_novedades (documento, tipo_novedad, nivel_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion, fecha_gestion, responsable_gestion)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (p_doc, tipo_nov_reg, nivel_txt_puro, fec_aplica_reg.strftime("%d/%m/%Y"), hor_aplica_reg.strftime("%I:%M %p"), det_reg, prof_reg.strip(), estado_ini, datetime.now().isoformat(), fec_ges_aut, prof_reg.strip() if fec_ges_aut else None),
                            fetch=False,
                        )
                        if "Nivel 3" in nivel_nov_reg:
                            run_query("UPDATE tratamientos SET novedades = novedades || ' // [CAMBIO PERMANENTE el ' || ? || ']' WHERE documento = ? AND estado = 'Activo'", (datetime.now().strftime("%d/%m/%Y"), p_doc), fetch=False)

                        run_query("UPDATE tratamientos SET alerta_revisada = 'NO' WHERE documento = ? AND estado = 'Activo'", (p_doc,), fetch=False)
                        log_auditoria(current_user.get("username", ""), "CREAR_NOVEDAD", "registro_novedades", p_doc, f"{tipo_nov_reg} - {nivel_txt_puro}")
                        st.success("✅ Novedad registrada exitosamente.")
                        st.rerun()

# ==================== 4. AYUDAS DIAGNÓSTICAS ====================
elif menu == "Ayudas diagnosticas":
    st.subheader("🧬 Gestión de Ayudas Diagnósticas, Peso y Oxígeno")
    with st.expander("➕ Solicitar Nueva Ayuda Diagnóstica", expanded=True):
        c_td, c_d = st.columns([1, 2.5])
        t_doc_dx = c_td.selectbox("Tipo Doc", TIPOS_DOC, index=0, key="tdoc_dx_ing")
        ced_dx_in = c_d.text_input("Documento del Paciente:", key="ced_dx_input_val").strip()
        p_piso_dx, p_zona_dx = "Norte", "Zona 1"

        if ced_dx_in:
            res_dx_p = run_query("SELECT nombre, tipo_documento, piso, zona FROM pacientes WHERE documento = ?", (ced_dx_in,))
            if res_dx_p:
                p_item = res_dx_p[0]
                p_piso_dx, p_zona_dx = p_item["piso"], p_item["zona"]
                st.success(f"✅ Paciente: **{p_item['nombre']}** ({p_item.get('tipo_documento', 'CC')}) | Piso: {p_piso_dx}")
            else:
                st.warning(f"⚠️ Cédula `{ced_dx_in}` no encontrada.")

        if ced_dx_in:
            piso_dx_sel = st.selectbox("Piso:", PISOS, index=PISOS.index(p_piso_dx) if p_piso_dx in PISOS else 0)
            zona_dx_sel = st.selectbox("Zona:", ZONAS, index=ZONAS.index(p_zona_dx) if p_zona_dx in ZONAS else 0)
            ayuda_sol_txt = st.text_input("Examen Solicitado:", key="ayuda_sol_input_field")
            origen_sol_sel = st.selectbox("Origen:", ["Línea de Ayudas Diagnósticas", "PGP"], key="origen_dx_sel")
            prestador_pgp_txt = st.text_input("Prestador PGP:", key="prest_pgp_field") if origen_sol_sel == "PGP" else ""

            c_d3, c_d4 = st.columns(2)
            peso_pac_txt = c_d3.text_input("Peso (Kg):", value="", key="peso_dx_field")
            req_o2_sel = c_d4.selectbox("¿Requiere Oxígeno?", ["NO", "SI"], index=0, key="req_o2_field")
            cant_o2_txt = st.text_input("Flujo O2:", key="cant_o2_field") if req_o2_sel == "SI" else ""

            transporte_dx_sel = st.selectbox("Transporte:", ["No requiere transporte", "TAB", "TAM", "MVR"], key="trans_dx_field")
            obs_dx_txt = st.text_area("Observaciones:", key="obs_dx_field")
            gestor_dx_reg = render_selectbox("Gestor:", cat_gestores, current_user.get("nombre_completo", ""), key="gestor_dx_select")

            col_btn_g1, col_btn_g2 = st.columns(2)
            with col_btn_g1:
                if st.button("💾 Guardar Solicitud"):
                    if not ced_dx_in or not ayuda_sol_txt.strip() or not gestor_dx_reg.strip():
                        st.error("Ingrese cédula, examen y gestor.")
                    elif origen_sol_sel == "PGP" and not prestador_pgp_txt.strip():
                        st.error("Indique prestador PGP.")
                    else:
                        fec_reg_dx = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                        run_query(
                            """INSERT INTO ayudas_diagnosticas (documento, tipo_documento, piso, zona, ayuda_solicitada, peso, requiere_oxigeno, cantidad_oxigeno, tipo_transporte, origen_solicitud, prestador_pgp, observaciones, estado, registrado_por, fecha_registro)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente', ?, ?)""",
                            (ced_dx_in, t_doc_dx, piso_dx_sel, zona_dx_sel, ayuda_sol_txt.strip(), peso_pac_txt.strip(), req_o2_sel, cant_o2_txt.strip(), transporte_dx_sel, origen_sol_sel, prestador_pgp_txt.strip(), obs_dx_txt.strip(), gestor_dx_reg.strip(), fec_reg_dx),
                            fetch=False
                        )
                        log_auditoria(current_user.get("username", ""), "SOLICITAR_AYUDA_DX", "ayudas_diagnosticas", ced_dx_in, ayuda_sol_txt.strip())
                        st.success("✅ Ayuda diagnóstica registrada.")
                        st.rerun()

            with col_btn_g2:
                if st.button("📋 Generar Formato Cita"):
                    res_p_n = run_query("SELECT nombre FROM pacientes WHERE documento = ?", (ced_dx_in,))
                    nom_p_gen = res_p_n[0]["nombre"] if res_p_n else "Paciente"
                    pgp_est = f"cuenta con PGP ({prestador_pgp_txt})" if origen_sol_sel == "PGP" and prestador_pgp_txt else "no tiene PGP"
                    texto_sol = f"""me colaboras con asignación de cita para el paciente en mencion, se encuentra hospitalizado en Salud en casa.\nPaciente: {nom_p_gen}\nCC {ced_dx_in}\nTipo de ayuda diagnóstica: {ayuda_sol_txt.strip()}\nEstado PGP: {pgp_est}"""
                    st.code(texto_sol, language="text")

    st.markdown("---")
    st.subheader("📊 Solicitudes Pendientes por Piso")
    cols_cont = st.columns(len(PISOS))
    data_dx_all = run_query("SELECT piso, estado FROM ayudas_diagnosticas")
    df_dx_all = pd.DataFrame(data_dx_all) if data_dx_all else pd.DataFrame(columns=["piso", "estado"])

    for idx, p_item in enumerate(PISOS):
        cant_p = len(df_dx_all[(df_dx_all['piso'].str.upper() == p_item.upper()) & (df_dx_all['estado'] == 'Pendiente')]) if not df_dx_all.empty else 0
        cols_cont[idx].markdown(
            f"""
            <div style="background: white; border-radius: 10px; padding: 12px; text-align: center; border: 1px solid #E2E8F0;">
                <span style="font-size: 0.75rem; color: #64748B; font-weight: 700;">{p_item}</span><br>
                <span style="font-size: 1.5rem; font-weight: 800; color: #0033A0;">{cant_p}</span><br>
                <span style="font-size: 0.7rem; color: #D97706;">Pendientes</span>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")
    c_fd1, c_fd2 = st.columns(2)
    piso_dx_filtro = c_fd1.selectbox("Filtrar Piso:", ["Todos"] + PISOS)
    estado_dx_filtro = c_fd2.selectbox("Estado:", ["Pendiente", "Completada", "Todas"])

    q_dx = "SELECT * FROM ayudas_diagnosticas WHERE 1=1"
    params_dx = []
    if estado_dx_filtro == "Pendiente":
        q_dx += " AND estado = 'Pendiente'"
    elif estado_dx_filtro == "Completada":
        q_dx += " AND estado = 'Completada'"
    if piso_dx_filtro != "Todos":
        q_dx += " AND piso = ?"
        params_dx.append(piso_dx_filtro)
    q_dx += " ORDER BY id DESC"

    data_dx = run_query(q_dx, tuple(params_dx))
    df_dx = pd.DataFrame(data_dx) if data_dx else pd.DataFrame()

    if not df_dx.empty:
        for _, r_dx in df_dx.iterrows():
            dx_id, dx_doc, dx_piso, dx_zona, dx_ayuda = r_dx['id'], r_dx['documento'], r_dx.get('piso', 'Norte'), r_dx.get('zona', 'Zona 1'), r_dx['ayuda_solicitada']
            dx_peso, dx_req_o2, dx_cant_o2, dx_transporte = r_dx['peso'], r_dx['requiere_oxigeno'], r_dx['cantidad_oxigeno'], r_dx.get('tipo_transporte', 'No requiere transporte')
            dx_origen, dx_prestador, dx_obs, dx_est, dx_reg, dx_fec = r_dx.get('origen_solicitud', ''), r_dx.get('prestador_pgp', ''), r_dx['observaciones'], r_dx['estado'], r_dx['registrado_por'], r_dx['fecha_registro']
            dx_fec_cita, dx_hora_cita, dx_lugar_cita, dx_informado = r_dx.get('fecha_cita', ''), r_dx.get('hora_cita', ''), r_dx.get('lugar_cita', ''), r_dx.get('paciente_informado', 'NO')
            dx_nom_acomp, dx_cont_acomp, dx_resuelto = r_dx.get('nombre_acompanante', ''), r_dx.get('contacto_acompanante', ''), r_dx.get('resuelto_por', '')

            p_nom_v = dx_doc
            res_nom = run_query("SELECT nombre FROM pacientes WHERE documento = ?", (dx_doc,))
            if res_nom: p_nom_v = res_nom[0]["nombre"]

            badge_est = "⚠️ PENDIENTE" if dx_est == 'Pendiente' else "✅ COMPLETADA"
            with st.expander(f"📌 [{dx_piso} - {dx_zona}] · {p_nom_v} (CC {dx_doc}) - {dx_ayuda} | {badge_est}", expanded=False):
                st.write(f"**Examen:** {dx_ayuda} | Origen: {dx_origen} ({dx_prestador}) | Transporte: {dx_transporte} | O2: {dx_req_o2} {dx_cant_o2} | Peso: {dx_peso}kg")
                if dx_obs: st.write(f"**Obs:** {dx_obs}")
                st.write(f"*Registrado: {dx_reg} el {dx_fec}*")

                c_c1, c_c2 = st.columns(2)
                fec_cita_val = c_c1.date_input("Fecha Cita:", value=datetime.now().date(), key=f"fec_cita_{dx_id}")
                hora_cita_val = c_c2.time_input("Hora Cita:", value=datetime.now().time(), key=f"hora_cita_{dx_id}")
                lugar_cita_val = st.text_input("Lugar Cita:", key=f"lugar_cita_{dx_id}", value=dx_lugar_cita or "")
                c_c3, c_c4 = st.columns(2)
                info_pac_val = c_c3.selectbox("¿Paciente informado?", ["NO", "SI"], index=1 if dx_informado=="SI" else 0, key=f"info_pac_{dx_id}")
                nombre_acomp_val = c_c4.text_input("Acompañante:", value=dx_nom_acomp or "", key=f"nom_acomp_{dx_id}") if info_pac_val == "SI" else ""
                gestor_res_val = render_selectbox("Diligenciado por:", cat_gestores, dx_resuelto or current_user.get("nombre_completo", ""), key=f"gestor_res_{dx_id}")

                b1, b2 = st.columns(2)
                with b1:
                    if st.button(f"✅ Resolver (#{dx_id})", key=f"btn_res_dx_{dx_id}"):
                        if not lugar_cita_val.strip() or not gestor_res_val.strip():
                            st.error("Diligencie lugar y gestor.")
                        else:
                            fec_ges_dx = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                            run_query(
                                """UPDATE ayudas_diagnosticas 
                                   SET estado = 'Completada', fecha_cita = ?, hora_cita = ?, lugar_cita = ?, paciente_informado = ?, nombre_acompanante = ?, resuelto_por = ?, fecha_gestion = ? 
                                   WHERE id = ?""",
                                (fec_cita_val.strftime("%d/%m/%Y"), hora_cita_val.strftime("%I:%M %p"), lugar_cita_val.strip(), info_pac_val, nombre_acomp_val.strip(), gestor_res_val.strip(), fec_ges_dx, dx_id),
                                fetch=False
                            )
                            log_auditoria(current_user.get("username", ""), "RESOLVER_AYUDA_DX", "ayudas_diagnosticas", str(dx_id), f"{lugar_cita_val}")
                            st.success("✅ Ayuda resuelta.")
                            st.rerun()
                with b2:
                    if st.button(f"🚗 Pedir Transporte (#{dx_id})", key=f"btn_trans_{dx_id}"):
                        t_txt = f"""me colaboras asignando transporte al paciente en mencion.\nPaciente: {p_nom_v}\nCC {dx_doc}\nPiso: {dx_piso} | Zona: {dx_zona}\nAyuda: {dx_ayuda}\nCita: {fec_cita_val.strftime('%d/%m/%Y')} {hora_cita_val.strftime('%I:%M %p')}\nLugar: {lugar_cita_val.strip()}\nTransporte: {dx_transporte}\nOxígeno: {dx_req_o2} ({dx_cant_o2})\nPeso: {dx_peso} kg"""
                        st.code(t_txt, language="text")
    else:
        st.info("No hay solicitudes registradas.")

# ==================== 5. CENSO Y PLANILLA EN VIVO ====================
elif menu == "Censo y Planilla en Vivo":
    col_v1, col_v2 = st.columns([3, 1])
    filtro_estado = col_v2.selectbox("Estado:", ["Activos", "Completados / Descanalizados", "Todos"])

    cond_est = "WHERE t.estado = 'Activo'"
    if filtro_estado == "Completados / Descanalizados": cond_est = "WHERE t.estado IN ('Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"
    elif filtro_estado == "Todos": cond_est = "WHERE t.estado IN ('Activo', 'Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"

    query_general = f"""
        SELECT 
            p.documento AS 'Documento', p.nombre AS 'Nombre y Apellido', p.municipio AS 'Municipio',
            p.barrio AS 'Barrio', p.piso AS 'Piso', p.aislamiento AS 'Aislamiento',
            COALESCE(p.alerta_permanente, '') AS 'Alerta_Permanente_Texto',
            COALESCE(t.tratamiento, 'Sin Tratamiento') AS 'Tratamiento', COALESCE(t.dosis, '-') AS 'Dosis',
            COALESCE(t.via_administracion, '-') AS 'Vía', COALESCE(t.tipo_acceso, 'Periférico') AS 'Tipo Acceso',
            COALESCE(t.frecuencia_horas, 0) AS 'Frec (h)', COALESCE(t.dias, 0) AS 'Días',
            COALESCE(t.fecha_inicio, '-') AS 'Fecha Inicio', COALESCE(t.fecha_fin, '-') AS 'Fecha Fin',
            COALESCE(t.fecha_retiro_cateter, '') AS 'Fecha Retiro Cateter',
            COALESCE(t.t1, '-') AS 'T1', COALESCE(t.t2, '-') AS 'T2', COALESCE(t.t3, '-') AS 'T3',
            COALESCE(t.estado, 'Sin TTO') AS 'Estado', COALESCE(t.descanalizado, 'NO') AS 'Descanalizado',
            COALESCE(t.alerta_revisada, 'NO') AS 'Alerta Revisada', COALESCE(t.novedades, '') AS 'NOVEDAD'
        FROM pacientes p
        JOIN tratamientos t ON p.documento = t.documento
        {cond_est}
        ORDER BY p.piso ASC, p.nombre ASC
    """
    data_gen = run_query(query_general)
    df_gen = pd.DataFrame(data_gen) if data_gen else pd.DataFrame()

    if not df_gen.empty:
        df_gen["Estado Tratamiento"] = df_gen.apply(
            lambda r: evaluar_alerta_fin(r["Fecha Fin"], r["Frec (h)"], r["Estado"], r["Descanalizado"], r["NOVEDAD"], r["Alerta Revisada"], r["Vía"])
            if r["Fecha Fin"] != "-" else "Sin tratamiento",
            axis=1,
        )
        df_gen["Alerta Fija"] = df_gen["Alerta_Permanente_Texto"].apply(lambda txt: "SI" if str(txt).strip() else "NO")

        c_f1, c_f2, c_f3, c_f4 = st.columns(4)
        filtro_piso = c_f1.multiselect("Piso", options=df_gen["Piso"].dropna().unique())
        filtro_muni = c_f2.multiselect("Municipio", options=df_gen["Municipio"].dropna().unique())
        filtro_via = c_f3.multiselect("Vía", options=df_gen["Vía"].dropna().unique())
        filtro_alerta = c_f4.multiselect("Alertas", options=df_gen["Estado Tratamiento"].dropna().unique())

        df_fil = df_gen.copy()
        if filtro_piso: df_fil = df_fil[df_fil["Piso"].isin(filtro_piso)]
        if filtro_muni: df_fil = df_fil[df_fil["Municipio"].isin(filtro_muni)]
        if filtro_via: df_fil = df_fil[df_fil["Vía"].isin(filtro_via)]
        if filtro_alerta: df_fil = df_fil[df_fil["Estado Tratamiento"].isin(filtro_alerta)]

        total_act = len(df_fil[df_fil["Estado"] == "Activo"])
        total_picc = len(df_fil[(df_fil["Tipo Acceso"].str.upper() == "PICC") & (df_fil["Estado"] == "Activo")])
        total_npt = len(df_fil[df_fil["Vía"].str.upper().str.contains("NPT") & (df_fil["Estado"] == "Activo")])
        total_be = len(df_fil[df_fil["Vía"].str.upper().str.contains("BE") & (df_fil["Estado"] == "Activo")])
        total_ais = len(df_fil[(df_fil["Aislamiento"] == "SI") & (df_fil["Estado"] == "Activo")])
        total_nov = len(df_fil[df_fil["Estado Tratamiento"].str.contains("NOVEDAD ACTIVA") & (df_fil["Estado"] == "Activo")])
        total_egr = len(df_fil[df_fil["Estado Tratamiento"].str.contains("SIN TTO ACTIVO|SEGUIMIENTO") & (df_fil["Estado"] == "Activo")])

        m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
        m1.markdown(f"<div class='metric-card'><div class='metric-val'>{total_act}</div><div class='metric-lbl'>ACTIVOS</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-card'><div class='metric-val'>{total_picc}</div><div class='metric-lbl'>🩸 PICC</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-card'><div class='metric-val'>{total_npt}</div><div class='metric-lbl'>🥣 NPT</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-card'><div class='metric-val'>{total_be}</div><div class='metric-lbl'>💉 BE / DOLOR</div></div>", unsafe_allow_html=True)
        m5.markdown(f"<div class='metric-card'><div class='metric-val'>{total_ais}</div><div class='metric-lbl'>☣️ AISLADOS</div></div>", unsafe_allow_html=True)
        m6.markdown(f"<div class='metric-card'><div class='metric-val'>{total_nov}</div><div class='metric-lbl'>📢 NOVEDADES</div></div>", unsafe_allow_html=True)
        m7.markdown(f"<div class='metric-card'><div class='metric-val'>{total_egr}</div><div class='metric-lbl'>⚠️ PEND. EGRESO</div></div>", unsafe_allow_html=True)

        html_filas = ""
        for num_idx, (_, r) in enumerate(df_fil.iterrows(), start=1):
            bg_color = get_color_fila_hex(r.to_dict())
            alerta_fija_texto = str(r["Alerta_Permanente_Texto"]).strip()
            celda_alerta = f'<span style="color:#D97706; font-weight:800;">SI ⚠️</span>' if alerta_fija_texto else '<span style="color:#94A3B8;">NO</span>'
            fin_fmt = datetime.fromisoformat(r["Fecha Fin"]).strftime("%d/%m %I:%M %p") if r["Fecha Fin"] != "-" else "-"
            ret_fmt = datetime.fromisoformat(r["Fecha Retiro Cateter"]).strftime("%d/%m %I:%M %p") if str(r["Fecha Retiro Cateter"]).strip() else "N/A"

            html_filas += f"""
            <tr style="background-color: {bg_color};">
                <td><strong>{num_idx}</strong></td>
                <td>{celda_alerta}</td>
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
                <td>{fin_fmt}</td>
                <td style="background-color: #FFF3CD; font-weight: bold;">{ret_fmt}</td>
                <td><code>{limpiar_texto_turno(r['T1'])}</code></td>
                <td><code>{limpiar_texto_turno(r['T2'])}</code></td>
                <td><code>{limpiar_texto_turno(r['T3'])}</code></td>
                <td><strong>{r['Estado Tratamiento']}</strong></td>
            </tr>
            """

        tabla_html = f"""
        <!DOCTYPE html><html><head><style>
        .tc {{ width: 100%; height: 520px; overflow: auto; border-radius: 12px; box-shadow: 0px 4px 18px rgba(0,0,0,0.06); }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; font-family: sans-serif; }}
        th {{ background: linear-gradient(180deg, #0033A0 0%, #002270 100%); color: #FFF; padding: 10px; position: sticky; top: 0; }}
        td {{ padding: 8px 6px; text-align: center; border: 1px solid #E6EAF0; white-space: nowrap; }}
        </style></head><body><div class="tc"><table>
        <thead><tr><th>N°</th><th>Alerta</th><th>Documento</th><th>Nombre</th><th>Mpio</th><th>Barrio</th><th>Piso</th><th>Aisla</th><th>Tratamiento</th><th>Dosis</th><th>Vía</th><th>Acceso</th><th>Frec</th><th>Días</th><th>Fin</th><th>Retiro</th><th>T1</th><th>T2</th><th>T3</th><th>Estado</th></tr></thead>
        <tbody>{html_filas}</tbody>
        </table></div></body></html>
        """
        components.html(tabla_html, height=530, scrolling=True)

        df_desc = df_fil.copy()
        df_desc["Alerta Fija"] = df_desc["Alerta_Permanente_Texto"].apply(lambda t: "SI" if str(t).strip() else "NO")
        cols_ex = ["Alerta Fija", "Documento", "Nombre y Apellido", "Municipio", "Barrio", "Piso", "Aislamiento", "Alerta_Permanente_Texto", "Tratamiento", "Dosis", "Vía", "Tipo Acceso", "Frec (h)", "Días", "Fecha Fin", "Fecha Retiro Cateter", "T1", "T2", "T3", "Estado Tratamiento", "NOVEDAD"]
        df_desc = df_desc[cols_ex].rename(columns={"Alerta_Permanente_Texto": "Detalle Alerta Fija"})
        excel_b = exportar_excel_con_colores(df_desc)

        st.download_button(
            label="📥 Descargar Excel (.xlsx)",
            data=excel_b,
            file_name=f"Planilla_SENC_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No hay registros bajo los filtros.")

# ==================== 6. ANALÍTICA Y GRÁFICOS ====================
elif menu == "Analitica y graficos":
    st.subheader("📈 Tablero Analítico y Gráficos Circulares SENC")
    data_p = run_query("SELECT * FROM pacientes")
    data_t = run_query("SELECT * FROM tratamientos")
    df_p = pd.DataFrame(data_p) if data_p else pd.DataFrame()
    df_t = pd.DataFrame(data_t) if data_t else pd.DataFrame()

    if not df_p.empty:
        col_f1, col_f2 = st.columns(2)
        f_desde = col_f1.date_input("Desde:", value=(datetime.now() - timedelta(days=30)).date())
        f_hasta = col_f2.date_input("Hasta:", value=datetime.now().date())

        df_p["fecha_ingreso_dt"] = pd.to_datetime(df_p["fecha_ingreso"], errors="coerce")
        df_p_rng = df_p[(df_p["fecha_ingreso_dt"].dt.date >= f_desde) & (df_p["fecha_ingreso_dt"].dt.date <= f_hasta)]

        total_ing = len(df_p_rng)
        act_rng = len(df_p_rng[df_p_rng["estado_paciente"] == "Activo"])
        egr_rng = len(df_p_rng[df_p_rng["estado_paciente"] == "Egresado"])

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Ingresos en Periodo", f"{total_ing}")
        col_m2.metric("Pacientes Activos", f"{act_rng}")
        col_m3.metric("Pacientes Egresados", f"{egr_rng}")
        pct_alta = round((egr_rng / total_ing) * 100, 1) if total_ing > 0 else 0
        col_m4.metric("% Tasa de Alta", f"{pct_alta}%")

        st.markdown("---")
        cg1, cg2 = st.columns(2)
        with cg1:
            st.markdown(generar_pie_chart_svg([act_rng, egr_rng], ["Activos", "Egresados"], ["#0033A0", "#6C757D"], "Distribución Estado Pacientes"), unsafe_allow_html=True)
        with cg2:
            if not df_t.empty:
                df_t_act = df_t[df_t["estado"] == "Activo"]
                if not df_t_act.empty:
                    vc = df_t_act["via_administracion"].value_counts()
                    st.markdown(generar_pie_chart_svg(list(vc.values), list(vc.index), ["#0033A0", "#00AEC7", "#28A745", "#FFC107", "#DC3545", "#6F42C1"][:len(vc)], "Distribución por Vía"), unsafe_allow_html=True)
    else:
        st.info("Sin datos para analítica.")

# ==================== 7. INFORMES MODULARES Y CONSOLIDADOS ====================
elif menu == "Informes":
    st.subheader("📑 Centro Consolidado de Informes Clínicos y Estadísticos")
    st.markdown("Selecciona el tipo de informe que deseas auditar y descargar. Puedes ver el balance estadístico en pantalla y exportar a Excel con formato institucional.")

    pestanas_informes = st.tabs([
        "🌐 Consolidado General",
        "💊 Tratamientos y Vías",
        "📋 Entrega de Turnos",
        "📢 Novedades Programación",
        "🧬 Ayudas Diagnósticas",
        "📦 Descarga Todo (Multi-Hoja)"
    ])

    # ---- 1. CONSOLIDADO GENERAL ----
    with pestanas_informes[0]:
        st.markdown("#### 🌐 Informe General: Censo, Movimientos y Estancias")
        q_gen = """
            SELECT 
                p.documento, p.tipo_documento, p.nombre, p.plan, p.programa, p.piso, p.zona,
                p.municipio, p.barrio, p.aislamiento, p.tipo_aislamiento, p.estado_paciente,
                p.fecha_ingreso, p.usuario_registro
            FROM pacientes p
            ORDER BY p.piso, p.zona, p.nombre
        """
        data_inf_gen = run_query(q_gen)
        df_inf_gen = pd.DataFrame(data_inf_gen) if data_inf_gen else pd.DataFrame()

        if not df_inf_gen.empty:
            df_inf_gen["dias_estancia"] = df_inf_gen["fecha_ingreso"].apply(
                lambda f: max(0, (datetime.now() - datetime.fromisoformat(str(f)[:10])).days) if f else 0
            )
            estancia_prom = round(df_inf_gen["dias_estancia"].mean(), 1) if not df_inf_gen.empty else 0

            tot_pac = len(df_inf_gen)
            tot_act = len(df_inf_gen[df_inf_gen["estado_paciente"] == "Activo"])
            tot_ais = len(df_inf_gen[df_inf_gen["aislamiento"] == "SI"])

            c_g1, c_g2, c_g3, c_g4 = st.columns(4)
            c_g1.metric("Total Pacientes Registrados", f"{tot_pac}")
            c_g2.metric("Pacientes Activos", f"{tot_act}")
            c_g3.metric("Pacientes Aislados", f"{tot_ais}")
            c_g4.metric("Estancia Promedio", f"{estancia_prom} días")

            st.dataframe(df_inf_gen, use_container_width=True, hide_index=True)

            excel_gen = exportar_excel_con_colores(df_inf_gen, "Censo_General")
            st.download_button(
                "📥 Descargar Informe Censo General (.xlsx)",
                data=excel_gen,
                file_name=f"Informe_General_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Sin registros de pacientes.")

    # ---- 2. TRATAMIENTOS Y VÍAS ----
    with pestanas_informes[1]:
        st.markdown("#### 💊 Informe Estadístico de Tratamientos y Accesos Vasculares")
        q_ttos = """
            SELECT 
                t.id, t.documento, p.nombre, p.piso, p.zona, t.tratamiento, t.dosis,
                t.via_administracion, t.tipo_acceso, t.frecuencia_horas, t.dias,
                t.fecha_inicio, t.fecha_fin, t.fecha_retiro_cateter, t.estado, t.usuario_modificacion
            FROM tratamientos t
            JOIN pacientes p ON t.documento = p.documento
            ORDER BY t.id DESC
        """
        data_ttos = run_query(q_ttos)
        df_ttos = pd.DataFrame(data_ttos) if data_ttos else pd.DataFrame()

        if not df_ttos.empty:
            t_act = len(df_ttos[df_ttos["estado"] == "Activo"])
            t_be = len(df_ttos[df_ttos["via_administracion"].str.contains("BE", na=False)])
            t_npt = len(df_ttos[df_ttos["via_administracion"].str.contains("NPT", na=False)])
            t_picc = len(df_ttos[df_ttos["tipo_acceso"].str.upper() == "PICC"])

            ct1, ct2, ct3, ct4 = st.columns(4)
            ct1.metric("TTOs Activos", f"{t_act}")
            ct2.metric("Bombas (BE)", f"{t_be}")
            ct3.metric("Nutrición (NPT)", f"{t_npt}")
            ct4.metric("Accesos PICC", f"{t_picc}")

            st.dataframe(df_ttos, use_container_width=True, hide_index=True)

            excel_ttos = exportar_excel_con_colores(df_ttos, "Tratamientos")
            st.download_button(
                "📥 Descargar Informe Tratamientos (.xlsx)",
                data=excel_ttos,
                file_name=f"Informe_Tratamientos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Sin registros de tratamientos.")

    # ---- 3. ENTREGA DE TURNOS ----
    with pestanas_informes[2]:
        st.markdown("#### 📋 Informe de Minuta y Novedades de Turno")
        q_et = """
            SELECT 
                et.id, et.documento, p.nombre, et.zona, et.prioridad, et.detalle_novedad,
                et.estado, et.registrado_por, et.fecha_creacion, et.resuelto_por, et.fecha_resolucion
            FROM entrega_turnos et
            LEFT JOIN pacientes p ON et.documento = p.documento
            ORDER BY et.id DESC
        """
        data_et = run_query(q_et)
        df_et = pd.DataFrame(data_et) if data_et else pd.DataFrame()

        if not df_et.empty:
            tot_turnos = len(df_et)
            act_turnos = len(df_et[df_et["estado"] == "Activa"])
            res_turnos = len(df_et[df_et["estado"] == "Resuelta"])
            pct_res = round((res_turnos / tot_turnos) * 100, 1) if tot_turnos > 0 else 0

            ce1, ce2, ce3, ce4 = st.columns(4)
            ce1.metric("Total Minutas", f"{tot_turnos}")
            ce2.metric("Pendientes / Activas", f"{act_turnos}")
            ce3.metric("Resueltas", f"{res_turnos}")
            ce4.metric("% Resolución", f"{pct_res}%")

            st.dataframe(df_et, use_container_width=True, hide_index=True)

            excel_et = exportar_excel_con_colores(df_et, "Entrega_Turnos")
            st.download_button(
                "📥 Descargar Informe Entrega de Turno (.xlsx)",
                data=excel_et,
                file_name=f"Informe_Turnos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Sin registros en entrega de turnos.")

    # ---- 4. NOVEDADES PROGRAMACIÓN ----
    with pestanas_informes[3]:
        st.markdown("#### 📢 Informe de Novedades de Programación y Rutas")
        q_nov = """
            SELECT 
                rn.id, rn.documento, p.nombre, p.piso, p.zona, rn.tipo_novedad,
                rn.nivel_novedad, rn.fecha_aplicacion, rn.hora_aplicacion, rn.detalle,
                rn.responsable, rn.estado_novedad, rn.responsable_gestion, rn.fecha_gestion
            FROM registro_novedades rn
            LEFT JOIN pacientes p ON rn.documento = p.documento
            ORDER BY rn.id DESC
        """
        data_nov = run_query(q_nov)
        df_nov = pd.DataFrame(data_nov) if data_nov else pd.DataFrame()

        if not df_nov.empty:
            tot_nov = len(df_nov)
            nov_pend = len(df_nov[df_nov["estado_novedad"] == "Pendiente"])
            nov_gest = len(df_nov[df_nov["estado_novedad"] == "Gestionada"])
            pct_ges = round((nov_gest / tot_nov) * 100, 1) if tot_nov > 0 else 0

            cn1, cn2, cn3, cn4 = st.columns(4)
            cn1.metric("Total Novedades", f"{tot_nov}")
            cn2.metric("Pendientes", f"{nov_pend}")
            cn3.metric("Gestionadas", f"{nov_gest}")
            cn4.metric("% Cumplimiento", f"{pct_ges}%")

            st.dataframe(df_nov, use_container_width=True, hide_index=True)

            excel_nov = exportar_excel_con_colores(df_nov, "Novedades")
            st.download_button(
                "📥 Descargar Informe Novedades (.xlsx)",
                data=excel_nov,
                file_name=f"Informe_Novedades_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Sin registros de novedades.")

    # ---- 5. AYUDAS DIAGNÓSTICAS ----
    with pestanas_informes[4]:
        st.markdown("#### 🧬 Informe de Ayudas Diagnósticas, Oxígeno y Transporte")
        q_ad = """
            SELECT 
                ad.id, ad.documento, p.nombre, ad.piso, ad.zona, ad.ayuda_solicitada,
                ad.origen_solicitud, ad.prestador_pgp, ad.peso, ad.requiere_oxigeno,
                ad.cantidad_oxigeno, ad.tipo_transporte, ad.estado, ad.registrado_por,
                ad.fecha_registro, ad.fecha_cita, ad.lugar_cita, ad.resuelto_por
            FROM ayudas_diagnosticas ad
            LEFT JOIN pacientes p ON ad.documento = p.documento
            ORDER BY ad.id DESC
        """
        data_ad = run_query(q_ad)
        df_ad = pd.DataFrame(data_ad) if data_ad else pd.DataFrame()

        if not df_ad.empty:
            tot_ad = len(df_ad)
            pend_ad = len(df_ad[df_ad["estado"] == "Pendiente"])
            comp_ad = len(df_ad[df_ad["estado"] == "Completada"])
            trans_ad = len(df_ad[df_ad["tipo_transporte"] != "No requiere transporte"])

            ca1, ca2, ca3, ca4 = st.columns(4)
            ca1.metric("Total Solicitudes", f"{tot_ad}")
            ca2.metric("Pendientes", f"{pend_ad}")
            ca3.metric("Completadas", f"{comp_ad}")
            ca4.metric("Con Ambulancia/Transporte", f"{trans_ad}")

            st.dataframe(df_ad, use_container_width=True, hide_index=True)

            excel_ad = exportar_excel_con_colores(df_ad, "Ayudas_Diagnosticas")
            st.download_button(
                "📥 Descargar Informe Ayudas Diagnósticas (.xlsx)",
                data=excel_ad,
                file_name=f"Informe_AyudasDx_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Sin registros de ayudas diagnósticas.")

    # ---- 6. CONSOLIDADO MULTI-HOJA ----
    with pestanas_informes[5]:
        st.markdown("#### 📦 Descargar Libro Maestro Consolidado (Todas las Hojas)")
        st.info("Genera un solo archivo de Excel que reúne en pestañas separadas: Censo General, Tratamientos, Turnos, Novedades y Ayudas Diagnósticas.")

        if st.button("📊 Generar Libro Consolidado Maestro"):
            libro_dict = {}
            if 'df_inf_gen' in locals() and not df_inf_gen.empty: libro_dict["Censo_General"] = df_inf_gen
            if 'df_ttos' in locals() and not df_ttos.empty: libro_dict["Tratamientos"] = df_ttos
            if 'df_et' in locals() and not df_et.empty: libro_dict["Entrega_Turnos"] = df_et
            if 'df_nov' in locals() and not df_nov.empty: libro_dict["Novedades"] = df_nov
            if 'df_ad' in locals() and not df_ad.empty: libro_dict["Ayudas_Dx"] = df_ad

            if libro_dict:
                excel_multi = exportar_excel_multi_hoja(libro_dict)
                st.download_button(
                    "⬇️ Descargar Libro Maestro (.xlsx)",
                    data=excel_multi,
                    file_name=f"Consolidado_Maestro_SENC_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No hay suficientes datos cargados para exportar.")

# ==================== 8. PANEL ADMINISTRADOR ====================
elif menu == "Panel Administrador" and is_admin():
    render_admin_panel()

# Pie de página institucional
st.markdown(
    """
    <div class="footer-autor">
        🏥 <strong>Sistema de Gestión y Control de Tratamientos Domiciliarios (SENC)</strong><br>
        Diseñado y desarrollado por <strong>Camilo Andrés Medina</strong> · Salud en Casa SURA Colombia
    </div>
    """,
    unsafe_allow_html=True,
)
