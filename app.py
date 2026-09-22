import io
import json
import os
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
    .badge-alerta-fija {
        background-color: #FFF3CD;
        color: #856404;
        padding: 10px 16px;
        border-radius: 8px;
        font-weight: 700;
        margin-bottom: 15px;
        border-left: 6px solid #FFC107;
        font-size: 0.95rem;
    }
    .turno-box {
        display: inline-block;
        background-color: #E9ECEF;
        padding: 3px 8px;
        border-radius: 4px;
        font-family: monospace;
        font-size: 0.85rem;
        color: #000000;
        margin-right: 5px;
        border: 1px solid #CED4DA;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ----------------- CARGA ROBUSTA DE CATÁLOGOS DESDE EXCEL -----------------
@st.cache_data(show_spinner=False)
def load_excel_catalog(filepath, fallback_val):
    if os.path.exists(filepath):
        try:
            # Forzamos leer solo la primera hoja, la primera columna, y limpiar espacios/nulos
            df = pd.read_excel(filepath, sheet_name=0, header=None)
            if not df.empty:
                items = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                items = sorted(list(set([i for i in items if i and i.lower() != 'nan'])))
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
    c.execute(
        "ALTER TABLE pacientes ADD COLUMN observaciones_clinicas TEXT DEFAULT"
        " ''"
    )
  if "alerta_permanente" not in cols_p:
    c.execute(
        "ALTER TABLE pacientes ADD COLUMN alerta_permanente TEXT DEFAULT ''"
    )
  if "fecha_ingreso" not in cols_p:
    c.execute("ALTER TABLE pacientes ADD COLUMN fecha_ingreso TEXT")

  conn.commit()

  # Carga automática desde CSV Inicial (Si existe)
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
        fec_adm = str(
            row.get("Fecha Admisión", datetime.now().date().isoformat())
        ).strip()

        c.execute(
            """
                    INSERT OR IGNORE INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso)
                    VALUES (?, ?, ?, ?, ?, 'Zona 1', ?, ?, 'NO', 'Ninguno', '', '', ?, 'Carga CSV Inicial', ?)
                """,
            (
                doc,
                nombre_completo,
                plan,
                programa,
                piso,
                municipio,
                barrio,
                estado_p,
                fec_adm[:10],
            ),
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
  horas_ciclo = [
      (dt_inicio + timedelta(hours=i * frecuencia_horas))
      for i in range(dosis_diarias)
  ]
  formato_horas = [
      (f"{dt.hour:02d}H" if dt.hour != 0 else "24H") for dt in horas_ciclo
  ]
  return horas_ciclo, formato_horas


def calcular_tratamiento_mixto(
    dt_inicio, dias, frecuencia_horas, mapa_admin, ajuste_dosis=0
):
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


def calcular_esquema_be_exacto(
    dias_ordenados, frecuencia_dosis, unidosis_puente
):
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


def evaluar_alerta_fin(
    fecha_fin_str,
    frecuencia_horas=8,
    estado="Activo",
    descanalizado="NO",
    novedades="",
    alerta_revisada="NO",
):
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


def exportar_excel_con_colores(df):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Planilla_SURA")
    workbook = writer.book
    worksheet = writer.sheets["Planilla_SURA"]

    fill_picc = PatternFill(
        start_color="FFD1D1", end_color="FFD1D1", fill_type="solid"
    )
    fill_npt = PatternFill(
        start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"
    )
    fill_be = PatternFill(
        start_color="D1E7DD", end_color="D1E7DD", fill_type="solid"
    )
    fill_analgesia = PatternFill(
        start_color="E2D9F3", end_color="E2D9F3", fill_type="solid"
    )
    fill_tb = PatternFill(
        start_color="FFE5D0", end_color="FFE5D0", fill_type="solid"
    )
    fill_push = PatternFill(
        start_color="CFF4FC", end_color="CFF4FC", fill_type="solid"
    )
    fill_aisla = PatternFill(
        start_color="E8D7F1", end_color="E8D7F1", fill_type="solid"
    )

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="0033A0", end_color="0033A0", fill_type="solid"
    )
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
    return (
        f"<div style='text-align: center; color: #6C757D; padding: 20px;'>"
        f"<strong>{titulo}</strong><br>Sin datos en este rango</div>"
    )

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
      slices_svg += (
          f"<circle cx='{cx}' cy='{cy}' r='{r}' fill='{color}' stroke='#fff'"
          " stroke-width='2'/>"
      )
    else:
      start_rad = math.radians(current_angle - 90)
      end_rad = math.radians(current_angle + angle - 90)
      x1 = cx + r * math.cos(start_rad)
      y1 = cy + r * math.sin(start_rad)
      x2 = cx + r * math.cos(end_rad)
      y2 = cy + r * math.sin(end_rad)
      large_arc = 1 if angle > 180 else 0
      d = (
          f"M {cx} {cy} L {x1:.2f} {y1:.2f} A {r} {r} 0 {large_arc} 1"
          f" {x2:.2f} {y2:.2f} Z"
      )
      slices_svg += (
          f"<path d='{d}' fill='{color}' stroke='#FFFFFF' stroke-width='1.5'/>"
      )

    current_angle += angle
    legend_html += (
        f"<span style='display:inline-block; margin-right: 12px; margin-bottom:"
        f" 4px;'><span style='display:inline-block; width:10px; height:10px;"
        f" background-color:{color}; border-radius:50%; margin-right:"
        f" 4px;'></span><strong>{label}:</strong> {val} ({pct:.1f}%)</span>"
    )

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
      "<p style='text-align: center; font-size: 0.85rem; color: #00AEC7"
      " !important; font-weight: 600;'>SALUD EN CASA (SENC) · GESTIÓN"
      " AGUDOS</p>",
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
  doc_busqueda = st.text_input(
      "Documento de Identidad del Paciente:", ""
  ).strip()

  if doc_busqueda != st.session_state.doc_previo_gestion:
    st.session_state.id_tto_editando = None
    st.session_state.id_tto_cambio = None
    st.session_state.id_tto_ver = None
    st.session_state.id_tto_suspender = None
    st.session_state.doc_previo_gestion = doc_busqueda

  programas_disponibles = ["Agudos"]
  municipios_disponibles = [
      "Medellín",
      "Bello",
      "Itagüí",
      "Envigado",
      "Sabaneta",
      "Caldas",
      "La Estrella",
      "Copacabana",
      "Girardota",
      "Barbosa",
      "Rionegro",
  ]
  pisos = ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
  zonas = [
      "Zona 1",
      "Zona 2",
      "Zona 3",
      "Zona 4",
      "Zona 5",
      "Zona 6",
      "Zona 7",
      "Zona 8",
      "PAS",
      "AMI",
  ]
  aislamientos_tipos = [
      "Contacto",
      "Gotas",
      "Aerosoles",
      "Protector (Invertido)",
  ]

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
          p_doc,
          p_nom,
          p_plan,
          p_prog,
          p_piso,
          p_zona,
          p_muni,
          p_barrio,
          p_aisla,
          p_tipo_aisla,
          p_obs,
          p_alerta_fija,
          p_user,
      ) = paciente_res[0]

      p_obs = str(p_obs) if p_obs else ""
      p_alerta_fija = str(p_alerta_fija) if p_alerta_fija else ""
      p_user = str(p_user) if p_user else ""

      if p_alerta_fija.strip():
        st.markdown(
            f"<div class='badge-alerta-fija'>📌 ALERTA FIJA PERMANENTE DE ESTE"
            f" PACIENTE: {p_alerta_fija.strip()}</div>",
            unsafe_allow_html=True,
        )

    with st.expander(
        "👤 Ficha Demográfica, Observaciones y Alertas Fijas",
        expanded=not bool(paciente_res),
    ):
      col1, col2, col3 = st.columns(3)

      if paciente_res:
        nombre = col1.text_input("Nombre y Apellido", value=p_nom)
        plan = col2.selectbox(
            "Plan",
            ["POS", "Póliza", "ARL"],
            index=(
                ["POS", "Póliza", "ARL"].index(p_plan)
                if p_plan in ["POS", "Póliza", "ARL"]
                else 0
            ),
        )
        programa = col3.selectbox("Programa", programas_disponibles, index=0)

        col4, col5, col6, col7 = st.columns(4)
        muni_idx = (
            municipios_disponibles.index(p_muni)
            if p_muni in municipios_disponibles
            else 0
        )
        municipio = col4.selectbox(
            "Municipio", municipios_disponibles, index=muni_idx
        )
        barrio = col5.text_input("Barrio", value=p_barrio)
        piso = col6.selectbox(
            "Piso", pisos, index=pisos.index(p_piso) if p_piso in pisos else 0
        )
        zona = col7.selectbox(
            "Zona", zonas, index=zonas.index(p_zona) if p_zona in zonas else 0
        )

        st.markdown("#### 📝 Observaciones Clínicas y Alerta Fija Permanente")
        col_obs1, col_obs2 = st.columns(2)
        obs_clinica = col_obs1.text_area(
            "Observación Clínica General (En blanco si no aplica):", value=p_obs
        )
        alerta_fija = col_obs2.text_area(
            "📌 Alerta Fija Permanente (Ej: NUNCA MOVER HORARIO, DIALÍTICO,"
            " TRABAJA MAÑANA):",
            value=p_alerta_fija,
        )

        st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
        c_ais1, c_ais2 = st.columns(2)
        tiene_aislamiento = c_ais1.radio(
            "¿Tiene aislamiento? (Obligatorio)",
            ["NO", "SI"],
            index=0 if p_aisla == "NO" else 1,
            horizontal=True,
        )
        tipo_aisla_val = "Ninguno"
        if tiene_aislamiento == "SI":
          idx_t = (
              aislamientos_tipos.index(p_tipo_aisla)
              if p_tipo_aisla in aislamientos_tipos
              else 0
          )
          tipo_aisla_val = c_ais2.selectbox(
              "Tipo de Aislamiento:", aislamientos_tipos, index=idx_t
          )

        usuario_edita = render_selectbox(
            "Tu Nombre / Profesional responsable:", cat_gestores, p_user
        )

        if st.button("Actualizar Ficha Paciente"):
          if usuario_edita.strip():
            run_query(
                """UPDATE pacientes SET nombre=?, plan=?, programa=?, piso=?, zona=?, 
                           municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, 
                           observaciones_clinicas=?, alerta_permanente=?, estado_paciente='Activo', usuario_registro=? WHERE documento=?""",
                (
                    nombre,
                    plan,
                    programa,
                    piso,
                    zona,
                    municipio,
                    barrio,
                    tiene_aislamiento,
                    tipo_aisla_val,
                    obs_clinica.strip(),
                    alerta_fija.strip(),
                    usuario_edita.strip(),
                    doc_busqueda,
                ),
                fetch=False,
            )
            st.success("Ficha guardada y actualizada correctamente.")
            st.rerun()
          else:
            st.error("Debes ingresar tu nombre como responsable.")
      else:
        st.info(
            "Paciente nuevo. Formulario en blanco (sin datos de otro"
            " paciente):"
        )
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
        obs_clinica = col_obs1.text_area(
            "Observación Clínica General (En blanco si no aplica):", value=""
        )
        alerta_fija = col_obs2.text_area(
            "📌 Alerta Fija Permanente (Ej: NUNCA MOVER HORARIO, DIALÍTICO,"
            " TRABAJA MAÑANA):",
            value="",
        )

        st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
        c_ais1, c_ais2 = st.columns(2)
        tiene_aislamiento = c_ais1.radio(
            "¿Tiene aislamiento? (Obligatorio)",
            ["NO", "SI"],
            index=0,
            horizontal=True,
        )
        tipo_aisla_val = "Ninguno"
        if tiene_aislamiento == "SI":
          tipo_aisla_val = c_ais2.selectbox(
              "Tipo de Aislamiento:", aislamientos_tipos
          )

        usuario_crea = render_selectbox(
            "Tu Nombre / Profesional que registra:", cat_gestores, ""
        )

        if st.button("Admitir Paciente a la Base"):
          if nombre and usuario_crea.strip():
            fec_hoy = datetime.now().date().isoformat()
            run_query(
                """INSERT INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, observaciones_clinicas, alerta_permanente, estado_paciente, usuario_registro, fecha_ingreso) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?, ?)""",
                (
                    doc_busqueda,
                    nombre,
                    plan,
                    programa,
                    piso,
                    zona,
                    municipio,
                    barrio,
                    tiene_aislamiento,
                    tipo_aisla_val,
                    obs_clinica.strip(),
                    alerta_fija.strip(),
                    usuario_crea.strip(),
                    fec_hoy,
                ),
                fetch=False,
            )
            st.success("Paciente admitido exitosamente.")
            st.rerun()
          else:
            st.error("Diligencie el nombre del paciente y su nombre.")

    if paciente_res:
      st.divider()

      ttos_activos = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC",
          (doc_busqueda,),
      )

      # ----------------- PANEL DIRECTO DE EGRESO / ALTA -----------------
      if ttos_activos:
        t_primero = ttos_activos[0]
        alerta_egreso = evaluar_alerta_fin(
            t_primero[9],
            t_primero[6],
            t_primero[20],
            t_primero[21],
            t_primero[18],
            t_primero[22],
        )

        with st.container():
          st.markdown(
              "### 🏁 Gestión de Egreso / Alta Médica (Cierre de Caso)"
          )
          if "FALTA 1 DOSIS" in alerta_egreso:
            st.error(
                "⚠️ **PACIENTE PENDIENTE DE RETIRO DE CATÉTER / ÚLTIMA VISITA.**"
                " Puede egresar al paciente una vez realizada la última"
                " aplicación."
            )

          with st.expander("✅ Egresar Paciente / Dar Alta Médica"):
            c_egr1, c_egr2 = st.columns(2)
            prof_egreso = render_selectbox(
                "Profesional que autoriza el alta:",
                cat_gestores,
                "",
                key="prof_alta",
            )
            fecha_alta = c_egr2.date_input(
                "Fecha de Egreso / Alta:", value=datetime.now().date()
            )
            motivo_alta = st.selectbox(
                "Motivo de Egreso:",
                [
                    "Tratamiento Completo Cumplido (Curación / Alta)",
                    "Hospitalización en IPS / Remisión",
                    "Retiro Voluntario / Cambio de Asegurador",
                    "Fallecimiento",
                ],
            )
            retira_cateter = st.radio(
                "¿Se retira catéter / descanalizado?",
                ["SI", "NO (Maneja PICC u otro servicio)"],
                horizontal=True,
            )

            if st.button("Confirmar Egreso y Alta Médica"):
              if not prof_egreso.strip():
                st.error("Ingrese su nombre como responsable del alta.")
              else:
                fec_alta_str = (
                    f"{fecha_alta.strftime('%d/%m/%Y')} por"
                    f" {prof_egreso.strip()}"
                )
                nov_alta_txt = f"[EGRESO / ALTA]: {motivo_alta} ({fec_alta_str})"

                run_query(
                    """UPDATE tratamientos SET estado = 'Completado', descanalizado = ?, 
                               novedades = novedades || ' // ' || ?, fecha_retiro_cateter = ? WHERE documento = ? AND estado = 'Activo'""",
                    (
                        "SI" if "SI" in retira_cateter else "NO",
                        nov_alta_txt,
                        fec_alta_str,
                        doc_busqueda,
                    ),
                    fetch=False,
                )
                run_query(
                    "UPDATE pacientes SET estado_paciente = 'Egresado' WHERE"
                    " documento = ?",
                    (doc_busqueda,),
                    fetch=False,
                )
                st.success("✅ Paciente egresado exitosamente.")
                st.rerun()

      # ----------------- PANELES SECUNDARIOS (CAMBIO, SUSPENDER, VER) -----------------
      if st.session_state.id_tto_cambio:
        tto_a_cambiar = run_query(
            "SELECT * FROM tratamientos WHERE id = ?",
            (st.session_state.id_tto_cambio,),
        )
        if tto_a_cambiar:
          tc = tto_a_cambiar[0]
          st.error(
              f"🔄 **CAMBIO DE TRATAMIENTO MÉDICO (CERRAR ANTERIOR E INICIAR"
              f" NUEVO)**\n\nMedicamento que se suspende: **{tc[2]}** ({tc[3]})"
              f" - Vía: {tc[4]}"
          )

          motivo_cambio_med = st.text_input(
              "Motivo médico del cambio / rotación (Obligatorio):",
              value="Cambio por orden médica / rotación de antibiótico",
          )

          col_cbtn1, col_cbtn2 = st.columns([2, 1])
          with col_cbtn1:
            if st.button("Confirmar Cierre de este TTO y Pasar a Nuevo"):
              fecha_cambio_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
              nov_cierre = (
                  f"[CAMBIO MÉDICO DE TTO el {fecha_cambio_str}]: Se suspende"
                  f" {tc[2]}. Motivo: {motivo_cambio_med}"
              )

              run_query(
                  """UPDATE tratamientos SET estado = 'Cambiado por Orden Médica', 
                             novedades = novedades || ' // ' || ? WHERE id = ?""",
                  (nov_cierre, tc[0]),
                  fetch=False,
              )

              st.session_state.id_tto_cambio = None
              st.success(
                  f"✅ {tc[2]} cerrado exitosamente. Formule el nuevo"
                  " medicamento abajo (inicia desde la dosis 1)."
              )
              st.rerun()

          with col_cbtn2:
            if st.button("Cancelar Cambio"):
              st.session_state.id_tto_cambio = None
              st.rerun()
          st.divider()

      if st.session_state.id_tto_suspender:
        tto_a_suspender = run_query(
            "SELECT * FROM tratamientos WHERE id = ?",
            (st.session_state.id_tto_suspender,),
        )
        if tto_a_suspender:
          ts = tto_a_suspender[0]
          st.error(
              f"⛔ **SUSPENDER TRATAMIENTO (CERRAR DEFINITIVAMENTE SIN"
              f" REEMPLAZAR)**\n\nMedicamento: **{ts[2]}** ({ts[3]}) - Vía:"
              f" {ts[4]}"
          )

          col_s1, col_s2 = st.columns(2)
          motivo_suspension = col_s1.text_input(
              "Motivo de la suspensión definitiva:",
              value="Suspensión por orden médica / RAM / Finalizado anticipado",
          )
          prof_susp = render_selectbox(
              "Profesional responsable de suspender:", cat_gestores, "", key="prof_susp"
          )

          col_sbtn1, col_sbtn2 = st.columns([2, 1])
          with col_sbtn1:
            if st.button("Confirmar Suspensión Definitiva (Dosis a 0)"):
              if not prof_susp.strip():
                st.error("Ingrese responsable.")
              else:
                fecha_susp_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                nov_susp = (
                    f"[SUSPENDIDO el {fecha_susp_str} por {prof_susp.strip()}]:"
                    f" {motivo_suspension}"
                )

                run_query(
                    """UPDATE tratamientos SET estado = 'Suspendido Médicamente', 
                               novedades = novedades || ' // ' || ?, usuario_modificacion = ? WHERE id = ?""",
                    (nov_susp, prof_susp.strip(), ts[0]),
                    fetch=False,
                )

                st.session_state.id_tto_suspender = None
                st.success(f"✅ {ts[2]} suspendido permanentemente.")
                st.rerun()

          with col_sbtn2:
            if st.button("Cancelar Suspensión"):
              st.session_state.id_tto_suspender = None
              st.rerun()
          st.divider()

      if st.session_state.id_tto_ver:
        tto_a_ver = run_query(
            "SELECT * FROM tratamientos WHERE id = ?",
            (st.session_state.id_tto_ver,),
        )
        if tto_a_ver:
          tv = tto_a_ver[0]
          st.info(f"👁️ **DETALLE COMPLETO DEL TRATAMIENTO ID #{tv[0]}**")
          st.write(f"**Medicamento / Terapia:** {tv[2]}")
          st.write(f"**Dosis:** {tv[3]}")
          st.write(f"**Vía y Acceso:** {tv[4]} - {tv[5]}")
          st.write(f"**Frecuencia:** Cada {tv[6]} horas por {tv[7]} días")
          
          fin_formateada_ver = datetime.fromisoformat(tv[9]).strftime("%d/%m %I:%M %p") if tv[9] != "-" else "-"
          retiro_formateado_ver = datetime.fromisoformat(tv[10]).strftime("%d/%m %I:%M %p") if tv[10] else "N/A"
          
          st.write(
              f"**Fechas:** Inicio {datetime.fromisoformat(tv[8]).strftime('%d/%m %I:%M %p')} | Fin {fin_formateada_ver} | Retiro Cat:"
              f" {retiro_formateado_ver}"
          )
          st.write(f"**Novedades y Registro:** {tv[24]}")
          if st.button("Ocultar Detalle"):
            st.session_state.id_tto_ver = None
            st.rerun()
          st.divider()

      # ----------------- LISTADO DE TRATAMIENTOS ACTIVOS -----------------
      st.markdown("### 💊 Tratamientos Activos del Paciente")
      if ttos_activos:
        for t in ttos_activos:
          alerta_t = evaluar_alerta_fin(t[9], t[6], t[20], t[21], t[18], t[22])
          col_info, col_btn = st.columns([7, 3])
          with col_info:
            st.info(
                f"**{t[2]} ({t[3]})** · **Vía:** {t[4]} · **Acceso:** {t[5]} ·"
                f" **Frec:** C/{t[6]}h · **Fin:**"
                f" {datetime.fromisoformat(t[9]).strftime('%d/%m %I:%M %p')}\n\n<span"
                f" class='turno-box'>T1: {t[11]}</span> <span"
                f" class='turno-box'>T2: {t[12]}</span> <span"
                f" class='turno-box'>T3: {t[13]}</span>\n\n**Estado:**"
                f" `{alerta_t}`",
                icon="💊"
            )
          with col_btn:
            c_b1, c_b2 = st.columns(2)
            c_b3, c_b4 = st.columns(2)
            if c_b1.button(f"👁️ Ver", key=f"btn_ver_{t[0]}"):
              st.session_state.id_tto_ver = t[0]
              st.session_state.id_tto_editando = None
              st.session_state.id_tto_cambio = None
              st.session_state.id_tto_suspender = None
              st.rerun()
            if c_b2.button(f"✏️ Mod", key=f"btn_edit_{t[0]}"):
              st.session_state.id_tto_editando = t[0]
              st.session_state.id_tto_cambio = None
              st.session_state.id_tto_ver = None
              st.session_state.id_tto_suspender = None
              st.rerun()
            if c_b3.button(f"🔄 Rotar", key=f"btn_cambio_{t[0]}"):
              st.session_state.id_tto_cambio = t[0]
              st.session_state.id_tto_editando = None
              st.session_state.id_tto_ver = None
              st.session_state.id_tto_suspender = None
              st.rerun()
            if c_b4.button(f"⛔ Susp", key=f"btn_susp_{t[0]}"):
              st.session_state.id_tto_suspender = t[0]
              st.session_state.id_tto_editando = None
              st.session_state.id_tto_cambio = None
              st.session_state.id_tto_ver = None
              st.rerun()
      else:
        st.info("El paciente no tiene tratamientos activos en este momento.")

      tto_edit_data = None
      if st.session_state.id_tto_editando:
        tto_edit_res = run_query(
            "SELECT * FROM tratamientos WHERE id = ?",
            (st.session_state.id_tto_editando,),
        )
        if tto_edit_res:
          tto_edit_data = tto_edit_res[0]
          st.warning(
              f"✏️ **Modificando Tratamiento Activo ID #{tto_edit_data[0]}:**"
              f" {tto_edit_data[2]}"
          )

      # ----------------- FORMULARIO: FORMULACIÓN O EDICIÓN -----------------
      st.markdown(
          "### ➕ Formular Nuevo Tratamiento (Inicia en Dosis 1) o Guardar"
          " Cambios"
      )

      vias = [
          "IV",
          "VO",
          "SC",
          "IM",
          "BE",
          "BE analgesia",
          "NPT",
          "Tuberculosis",
          "Materna",
          "NBZ",
          "PUSH",
      ]
      accesos = ["Periférico", "PICC", "SC", "Ninguno"]

      val_via = (
          tto_edit_data[4]
          if (tto_edit_data and tto_edit_data[4] in vias)
          else "IV"
      )
      
      # SELECCIÓN PRIMERO DE LA VÍA
      col_v, col_m, col_d, col_a = st.columns([1.5, 2.5, 1, 1])
      t_via = col_v.selectbox(
          "Vía de Adm:", vias, index=vias.index(val_via) if val_via in vias else 0
      )

      # LÓGICAS CONDICIONALES DE AUTOCOMPLETADO SEGÚN VÍA
      val_med = tto_edit_data[2] if tto_edit_data else ""
      val_dos = tto_edit_data[3] if tto_edit_data else ""
      val_acc = (
          tto_edit_data[5]
          if (tto_edit_data and tto_edit_data[5] in accesos)
          else "Periférico"
      )

      t_nombre = ""
      bloquear_nombre = False
      
      if t_via == "BE analgesia":
          t_nombre = "BOMBA DE ANALGESIA / VIGILANCIA"
          bloquear_nombre = True
      elif t_via == "NPT":
          t_nombre = "NUTRICIÓN PARENTERAL TOTAL (NPT)"
          bloquear_nombre = True
          if not val_dos: val_dos = "1 Bolsa"
      elif t_via == "Materna":
          t_nombre = "PROGRAMA MATERNO / EDUCACIÓN"
          bloquear_nombre = True
          val_dos = "N/A"
      elif t_via == "Tuberculosis":
          t_nombre = "TRATAMIENTO TUBERCULOSIS (TB)"
          bloquear_nombre = True
      elif t_via == "BE":
          lista_meds = cat_meds_be
      else:
          lista_meds = cat_meds_agudos

      if bloquear_nombre:
          t_nombre = col_m.text_input("Terapia:", value=t_nombre, disabled=True)
      else:
          t_nombre = render_selectbox(
              "Medicamento (Autocompletable):", lista_meds, val_med, key="med_select"
          )

      t_dosis = col_d.text_input("Dosis:", value=val_dos)
      
      # Si es SC, preguntar si es con catéter o no
      if t_via == "SC":
          con_cateter_sc = col_a.checkbox("¿Usa Catéter?", value=True if "SC" in val_acc else False)
          t_acceso = "SC" if con_cateter_sc else "Ninguno"
      else:
          t_acceso = col_a.selectbox(
              "Acceso:",
              accesos,
              index=accesos.index(val_acc) if val_acc in accesos else 0,
          )

      val_frec = safe_int(tto_edit_data[6], 8) if tto_edit_data else 8
      val_dias = safe_int(tto_edit_data[7], 5) if tto_edit_data else 5
      val_ini = (
          datetime.fromisoformat(tto_edit_data[8])
          if tto_edit_data
          else datetime.now()
      )

      # ----------------- CASO 1: NPT Y TUBERCULOSIS (DÍAS ESPECÍFICOS) -----------------
      npt_dias_semana_json = "[]"
      npt_horas_infusion = 12
      npt_hora_desconexion_str = ""
      dias_semana_npt = None

      if t_via in ["NPT", "Tuberculosis"]:
        st.markdown(
            f"#### {'🥣 Protocolo NPT: Programación' if t_via == 'NPT' else '🟧 Protocolo Tuberculosis (Días asignados)'}"
        )
        col_npt1, col_npt2 = st.columns(2)
        dias_def = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"] if t_via == "Tuberculosis" else ["Lunes", "Martes", "Miércoles"]
        
        if tto_edit_data and len(tto_edit_data) > 26 and tto_edit_data[26]:
          try:
            dias_def = json.loads(tto_edit_data[26])
          except Exception:
            pass

        dias_semana_npt = col_npt1.multiselect(
            "Días de atención en la semana (Descansa los demás):",
            [
                "Lunes",
                "Martes",
                "Miércoles",
                "Jueves",
                "Viernes",
                "Sábado",
                "Domingo",
            ],
            default=dias_def,
        )
        npt_dias_semana_json = json.dumps(dias_semana_npt)
        
        if t_via == "NPT":
            horas_opc = [12, 14, 16, 18, 20, 24]
            h_inf_def = (
                safe_int(tto_edit_data[27], 12)
                if (tto_edit_data and len(tto_edit_data) > 27)
                else 12
            )
            npt_horas_infusion = col_npt2.selectbox(
                "Horas de infusión programadas:",
                horas_opc,
                index=horas_opc.index(h_inf_def) if h_inf_def in horas_opc else 0,
            )

      # ----------------- CASO 2: BE ANALGESIA -----------------
      elif t_via == "BE analgesia":
        st.markdown("#### 🟪 Protocolo BE Analgesia (Vigilancia C/12 Horas)")
        st.info(
            "📌 **Vigilancia de Bomba de Dolor:** Se programa visita cada 12"
            " horas durante 3 días. En la última visita del 3er día se retira"
            " dispositivo y catéter."
        )

      # ----------------- CASO 3: SC CON EDUCACIÓN OPCIONAL -----------------
      visita_educacion_sc = ""
      if t_via == "SC":
        st.markdown(
            "#### 💉 Protocolo SC (Autocuidado / Aplicación por Paciente o"
            " Cuidador)"
        )
        requiere_edu = st.checkbox(
            "¿Requiere programar 1 visita de educación / supervisión"
            " opcional?",
            value=False,
        )
        if requiere_edu:
          col_ed1, col_ed2 = st.columns(2)
          fecha_ed = col_ed1.date_input(
              "Fecha Visita Educación:", value=datetime.now().date()
          )
          hora_ed = col_ed2.time_input(
              "Hora Visita Educación:", value=datetime.now().time()
          )
          visita_educacion_sc = f"{fecha_ed.strftime('%d/%m/%Y')} {hora_ed.strftime('%I:%M %p')}"

      col_f1, col_f2 = st.columns(2)

      if t_via == "BE":
        frecuencia = col_f1.selectbox(
            "Concentración de la Bomba Elastomérica (BE):",
            [8, 6],
            format_func=lambda x: f"BE C/{x}h ({3 if x==8 else 4} dosis por bomba)",
            help=(
                "C/8h = Cada bomba contiene 3 dosis. C/6h = Cada bomba contiene"
                " 4 dosis."
            ),
        )
      elif t_via == "Tuberculosis":
        frecuencia = 24
        col_f1.info("Frecuencia fijada en C/24h (Dosis Diaria Única)")
      else:
        frecuencias = [4, 6, 8, 12, 24]
        def_frec_idx = (
            frecuencias.index(12)
            if t_via == "BE analgesia"
            else (
                frecuencias.index(val_frec) if val_frec in frecuencias else 2
            )
        )
        frecuencia = col_f1.selectbox(
            "Frecuencia (Horas):", frecuencias, index=def_frec_idx
        )

      def_dias = 3 if t_via == "BE analgesia" else val_dias
      dias = col_f2.number_input(
          "Días de Tratamiento Ordenados (Días / Bombas):",
          min_value=1,
          max_value=999,
          value=def_dias,
      )

      col_h1, col_h2 = st.columns(2)
      fecha_inicio = col_h1.date_input("Fecha Inicio", value=val_ini.date())
      hora_inicio = col_h2.time_input("Hora Inicio", value=val_ini.time())
      dt_inicio = datetime.combine(fecha_inicio, hora_inicio)

      # ----------------- PROTOCOLO BOMBA ELASTOMÉRICA (BE) ESTRICTO -----------------
      unidosis_puente = 0
      bombas_calc = 0
      dosis_final_cierre = 0

      if t_via == "BE":
        st.markdown(
            f"#### 🧪 Protocolo Clínico BE: C/{frecuencia}h (1 Día = 1 Bomba"
            " Elastomérica C/24h)"
        )
        if dt_inicio.hour >= 23 or dt_inicio.hour < 6:
          st.error(
              "🚨 **REGLA T3:** Después de las 23:00 NO se instala BE. Iniciar"
              " con **Unidosis IV de puente** y colocar BE en T1 o T2."
          )

        col_be1, col_be2 = st.columns(2)
        val_puente_prev = (
            safe_int(tto_edit_data[29], 0)
            if (tto_edit_data and len(tto_edit_data) > 29)
            else 0
        )
        unidosis_puente = col_be1.number_input(
            "Unidosis IV administradas como puente inicial (si aplica):",
            min_value=0,
            max_value=50,
            value=val_puente_prev,
        )

        total_unidosis_ord, bombas_calc, dosis_final_cierre, dosis_por_bomba = (
            calcular_esquema_be_exacto(dias, frecuencia, unidosis_puente)
        )

        col_be2.markdown(
            f"""
            <div style='background-color: #D1E7DD; color: #000000; padding: 12px; border-radius: 8px; font-size: 0.85rem;'>
                <strong>📦 CÁLCULO CLÍNICO BE EXACTO (C/{frecuencia}h):</strong><br>
                • Días ordenados: <strong>{dias} Días</strong> = <strong>{dias} Bombas Elastoméricas</strong> (1 cambio c/24h).<br>
                • Total unidosis clínicas: {total_unidosis_ord} dosis.<br>
                • Unidosis puente previas: {unidosis_puente} dosis.<br>
                • Cierre con unidosis IV: {dosis_final_cierre} dosis.
            </div>
            """,
            unsafe_allow_html=True,
        )

      col_aj1, col_aj2, col_aj3 = st.columns([1, 1, 2])
      tipo_unidad_ajuste = col_aj1.selectbox(
          "Unidad de Ajuste:",
          ["Unidosis", "BE Completa"],
          index=0 if "BE" not in t_via else 1,
      )
      val_ajuste_prev = (
          safe_int(tto_edit_data[23], 0)
          if (tto_edit_data and len(tto_edit_data) > 23)
          else 0
      )
      cant_ajuste = col_aj2.number_input(
          "Cantidad (+ Sumar / - Restar):",
          min_value=-99,
          max_value=99,
          value=val_ajuste_prev,
      )

      dosis_equiv_ajuste = cant_ajuste
      if tipo_unidad_ajuste == "BE Completa":
        dosis_por_be_adj = 4 if frecuencia == 6 else 3
        dosis_equiv_ajuste = cant_ajuste * dosis_por_be_adj

      motivo_ajuste = col_aj3.selectbox(
          "Motivo del Ajuste:",
          [
              "Ninguno / Dosis exactas de orden médica",
              "Paciente Ausente / No estaba",
              "Reprogramación de dosis",
              "Dosis no administrada por enfermería",
              "Suspensión médica anticipada",
              "Reposición autorizada por médico",
          ],
      )

      # ----------------- CÁLCULO DE FECHAS (TTO VS RETIRO CATÉTER DEPENDIENDO DE VÍA) -----------------
      medicamento_infusion_corta = any(
          m in t_nombre.upper() for m in ["VANCOMICINA", "ZAVICEFTA"]
      )
      
      requiere_retiro_cateter = t_via in ["IV", "BE", "BE analgesia", "PUSH"] or (t_via == "SC" and t_acceso == "SC")

      if t_via == "NPT":
        dt_desconexion = dt_inicio + timedelta(hours=npt_horas_infusion)
        npt_hora_desconexion_str = dt_desconexion.strftime("%H:%M")
        h_con = dt_inicio.hour
        h_des = dt_desconexion.hour

        t1_calc = (
            f"CONECTAR ({dt_inicio.strftime('%H:%M')})"
            if 6 <= h_con < 14
            else (
                f"DESCONECTAR ({npt_hora_desconexion_str})"
                if 6 <= h_des < 14
                else "—"
            )
        )
        t2_calc = (
            f"CONECTAR ({dt_inicio.strftime('%H:%M')})"
            if 14 <= h_con < 22
            else (
                f"DESCONECTAR ({npt_hora_desconexion_str})"
                if 6 <= h_des < 22
                else "—"
            )
        )
        t3_calc = (
            f"CONECTAR ({dt_inicio.strftime('%H:%M')})"
            if (h_con >= 22 or h_con < 6)
            else (
                f"DESCONECTAR ({npt_hora_desconexion_str})"
                if (h_des >= 22 or h_des < 6)
                else "—"
            )
        )

        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
        dosis_base_fijas = dias
        dosis_totales = dias
        nuevo_mapa_admin = {}

      elif t_via == "Tuberculosis":
        h_tb = dt_inicio.hour
        t1_calc = f"VISITA TB ({dt_inicio.strftime('%H:%M')})" if 6 <= h_tb < 14 else "—"
        t2_calc = f"VISITA TB ({dt_inicio.strftime('%H:%M')})" if 14 <= h_tb < 22 else "—"
        t3_calc = f"VISITA TB ({dt_inicio.strftime('%H:%M')})" if (h_tb >= 22 or h_tb < 6) else "—"
        
        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = None
        dosis_base_fijas = dias
        dosis_totales = dias
        nuevo_mapa_admin = {}

      elif t_via == "Materna":
        h_mat = dt_inicio.hour
        t1_calc = f"VISITA MATERNA ({dt_inicio.strftime('%H:%M')})" if 6 <= h_mat < 14 else "—"
        t2_calc = f"VISITA MATERNA ({dt_inicio.strftime('%H:%M')})" if 14 <= h_mat < 22 else "—"
        t3_calc = "—"
        
        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = None
        dosis_base_fijas = dias
        dosis_totales = dias
        nuevo_mapa_admin = {}

      elif t_via == "BE":
        h_be = dt_inicio.hour
        horario_visita_be = f"{dt_inicio.strftime('%H:%M')}H (CAMBIO BE C/24H)"
        t1_calc = horario_visita_be if 6 <= h_be < 14 else "—"
        t2_calc = horario_visita_be if 14 <= h_be < 22 else "—"
        t3_calc = horario_visita_be if (h_be >= 22 or h_be < 6) else "—"

        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = fecha_fin_calculada + timedelta(hours=24) if requiere_retiro_cateter else None
        dosis_base_fijas = dias
        dosis_totales = dias
        nuevo_mapa_admin = {}

      elif t_via == "BE analgesia":
        h_vig1 = dt_inicio.hour
        h_vig2 = (dt_inicio + timedelta(hours=12)).hour
        t1_calc = f"{h_vig1:02d}H (VIGILANCIA BE)" if 6 <= h_vig1 < 14 else "—"
        t2_calc = f"{h_vig2:02d}H (VIGILANCIA BE)" if 14 <= h_vig2 < 22 else "—"
        t3_calc = "—"

        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
        dosis_base_fijas = dias * 2
        dosis_totales = dosis_base_fijas
        nuevo_mapa_admin = {}

      elif medicamento_infusion_corta:
        dt_retiro = dt_inicio + timedelta(hours=2)
        h_inst = dt_inicio.hour
        t1_calc = (
            f"INSTALAR ({dt_inicio.strftime('%H:%M')}) -> RETIRAR"
            f" ({dt_retiro.strftime('%H:%M')})"
            if 6 <= h_inst < 14
            else "—"
        )
        t2_calc = (
            f"INSTALAR ({dt_inicio.strftime('%H:%M')}) -> RETIRAR"
            f" ({dt_retiro.strftime('%H:%M')})"
            if 14 <= h_inst < 22
            else "—"
        )
        t3_calc = (
            f"INSTALAR ({dt_inicio.strftime('%H:%M')}) -> RETIRAR"
            f" ({dt_retiro.strftime('%H:%M')})"
            if (h_inst >= 22 or h_inst < 6)
            else "—"
        )

        fecha_fin_calculada = dt_inicio + timedelta(days=dias)
        fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
        dosis_base_fijas = int((dias * 24) / frecuencia)
        dosis_totales = dosis_base_fijas
        nuevo_mapa_admin = {}

      else:
        # Vías regulares (IV, SC, NBZ, PUSH, VO, IM)
        horas_ciclo, formato_horas = obtener_horas_ciclo(dt_inicio, frecuencia)
        st.markdown("#### 🕒 Asignación de Responsable por Dosis del Día")
        cols_horarios = st.columns(len(formato_horas))
        nuevo_mapa_admin = {}

        mapa_guardado = {}
        if tto_edit_data and len(tto_edit_data) > 20 and tto_edit_data[20]:
          try:
            mapa_guardado = json.loads(tto_edit_data[20])
          except Exception:
            mapa_guardado = {}

        for idx, h_str in enumerate(formato_horas):
          with cols_horarios[idx]:
            def_idx = 1 if mapa_guardado.get(h_str) == "Cuidador" else 0
            resp_sel = st.selectbox(
                f"Dosis {h_str}:",
                ["SENC", "Cuidador"],
                index=def_idx,
                key=f"adm_dosis_{h_str}",
            )
            nuevo_mapa_admin[h_str] = resp_sel

        (
            fecha_fin_calculada,
            t1_calc,
            t2_calc,
            t3_calc,
            dosis_base_fijas,
            dosis_totales,
            d_senc,
            d_cuid,
        ) = calcular_tratamiento_mixto(
            dt_inicio, dias, frecuencia, nuevo_mapa_admin, dosis_equiv_ajuste
        )
        fecha_retiro_cateter_calc = fecha_fin_calculada if requiere_retiro_cateter else None
        visita_educacion = ""

      usuario_responsable = render_selectbox(
          "✍️ Profesional responsable que guarda/modifica (Obligatorio):",
          cat_gestores,
          "",
          key="prof_formular"
      )

      st.markdown("#### ⚖️ Conciliación Final y Fechas Clínicas")
      mb1, mb2, mb3, mb4 = st.columns(4)
      mb1.metric("Días Ordenados", f"{dias} Días")
      mb2.metric(
          "Ajuste Dosis",
          f"{cant_ajuste} ({tipo_unidad_ajuste})" if cant_ajuste != 0 else "0",
      )
      mb3.metric(
          "Fecha Fin Tratamiento",
          fecha_fin_calculada.strftime("%d/%m %I:%M %p"),
      )
      mb4.metric(
          "Fecha Retiro Catéter",
          fecha_retiro_cateter_calc.strftime("%d/%m %I:%M %p") if fecha_retiro_cateter_calc else "N/A",
      )

      c_p1, c_p2, c_p3 = st.columns(3)
      c_p1.metric("Turno T1 (06H-14H)", t1_calc)
      c_p2.metric("Turno T2 (14H-22H)", t2_calc)
      c_p3.metric("Turno T3 (22H-06H)", t3_calc)

      col_btn_guardar, col_btn_cancelar = st.columns([3, 1])
      
      with col_btn_guardar:
          btn_label = (
              "💾 Actualizar Cambios del Tratamiento"
              if tto_edit_data
              else "💾 Guardar Formulación de Tratamiento (Inicia en Dosis 1)"
          )
          if st.button(btn_label):
            if not usuario_responsable.strip():
              st.error("❌ Indique el profesional responsable.")
            elif not t_nombre.strip():
              st.error("❌ Ingrese el medicamento o terapia.")
            else:
              fecha_reg_str = datetime.now().isoformat()
              nov_txt = (
                  f"[{'ACTUALIZACIÓN' if tto_edit_data else 'INICIO NUEVO TTO'}"
                  f" {t_via}]: {t_nombre} {t_dosis} por"
                  f" {usuario_responsable.strip()}"
              )

              if t_via == "NPT":
                nov_txt += (
                    f" // NPT Infusión {npt_horas_infusion}h (Desconectar a las"
                    f" {npt_hora_desconexion_str})"
                )
              elif t_via == "BE":
                nov_txt += (
                    f" // Esquema BE (C/{frecuencia}h): {unidosis_puente} IV"
                    f" puente + {dias} Bombas Elastoméricas. Retiro catéter:"
                    f" {fecha_retiro_cateter_calc.strftime('%d/%m/%Y')}"
                )
              elif t_via == "BE analgesia":
                nov_txt += (
                    " // Vigilancia BE analgesia C/12h por 3 días. Retiro catéter:"
                    f" {fecha_retiro_cateter_calc.strftime('%d/%m/%Y')}"
                )
              elif t_via == "SC" and visita_educacion_sc:
                nov_txt += (
                    f" // Visita de educación SC programada:"
                    f" {visita_educacion_sc}"
                )

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
                        t_nombre.strip(),
                        t_dosis,
                        t_via,
                        t_acceso,
                        frecuencia,
                        dias,
                        dt_inicio.isoformat(),
                        fecha_fin_calculada.isoformat(),
                        fecha_retiro_cateter_calc.isoformat() if fecha_retiro_cateter_calc else "",
                        t1_calc,
                        t2_calc,
                        t3_calc,
                        json.dumps(nuevo_mapa_admin),
                        visita_educacion_sc,
                        dosis_base_fijas,
                        dosis_equiv_ajuste,
                        tipo_unidad_ajuste,
                        motivo_ajuste,
                        npt_dias_semana_json,
                        npt_horas_infusion,
                        npt_hora_desconexion_str,
                        unidosis_puente,
                        bombas_calc,
                        dosis_final_cierre,
                        nov_txt,
                        usuario_responsable.strip(),
                        tto_edit_data[0],
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
                        doc_busqueda,
                        t_nombre.strip(),
                        t_dosis,
                        t_via,
                        t_acceso,
                        frecuencia,
                        dias,
                        dt_inicio.isoformat(),
                        fecha_fin_calculada.isoformat(),
                        fecha_retiro_cateter_calc.isoformat() if fecha_retiro_cateter_calc else "",
                        t1_calc,
                        t2_calc,
                        t3_calc,
                        json.dumps(nuevo_mapa_admin),
                        visita_educacion_sc,
                        dosis_base_fijas,
                        dosis_equiv_ajuste,
                        tipo_unidad_ajuste,
                        motivo_ajuste,
                        npt_dias_semana_json,
                        npt_horas_infusion,
                        npt_hora_desconexion_str,
                        unidosis_puente,
                        bombas_calc,
                        dosis_final_cierre,
                        nov_txt,
                        usuario_responsable.strip(),
                        fecha_reg_str,
                    ),
                    fetch=False,
                )
                st.success(
                    "✅ Nuevo tratamiento formulado exitosamente (conteo desde"
                    " dosis 1)."
                )
    
              st.rerun()
              
      with col_btn_cancelar:
          if st.button("❌ Cancelar / Limpiar"):
              st.session_state.id_tto_editando = None
              st.rerun()

# ----------------- SECCIÓN 2: PESTAÑA NOVEDADES Y AJUSTES DE RUTA -----------------
elif menu == "📢 Novedades y Ajustes de Ruta":
  st.subheader("📢 Gestión de Novedades Clínicas y Modificación de Horarios")

  st.markdown("### 📋 Pacientes con Novedades Activas / Pendientes por Gestionar")
  query_novs_generales = """
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
  conn = sqlite3.connect(DB_FILE)
  df_novs_pendientes_general = pd.read_sql_query(query_novs_generales, conn)
  conn.close()

  doc_preseleccionado = ""
  if not df_novs_pendientes_general.empty:
    col_nf1, col_nf2, col_nf3 = st.columns(3)
    f_zonas_nov = col_nf1.multiselect(
        "Filtrar por Zona:",
        options=df_novs_pendientes_general["Zona"].dropna().unique(),
    )
    f_pisos_nov = col_nf2.multiselect(
        "Filtrar por Piso:",
        options=df_novs_pendientes_general["Piso"].dropna().unique(),
    )
    f_tipos_nov = col_nf3.multiselect(
        "Filtrar por Tipo Novedad:",
        options=df_novs_pendientes_general["Tipo Novedad"].dropna().unique(),
    )

    df_novs_filtradas = df_novs_pendientes_general.copy()
    if f_zonas_nov:
      df_novs_filtradas = df_novs_filtradas[
          df_novs_filtradas["Zona"].isin(f_zonas_nov)
      ]
    if f_pisos_nov:
      df_novs_filtradas = df_novs_filtradas[
          df_novs_filtradas["Piso"].isin(f_pisos_nov)
      ]
    if f_tipos_nov:
      df_novs_filtradas = df_novs_filtradas[
          df_novs_filtradas["Tipo Novedad"].isin(f_tipos_nov)
      ]

    st.info(f"📢 Pendientes: `{len(df_novs_filtradas)}` novedades activas")
    st.dataframe(df_novs_filtradas, use_container_width=True)

    pacientes_unicos = df_novs_filtradas[
        ["Cédula", "Nombre y Apellido", "Zona"]
    ].drop_duplicates()
    opciones_pacientes_nov = [
        f"{row['Cédula']} - {row['Nombre y Apellido']} ({row['Zona']})"
        for _, row in pacientes_unicos.iterrows()
    ]

    paciente_escogido = st.selectbox(
        "👉 Seleccionar paciente para desplegar y gestionar sus novedades:",
        ["-- Seleccionar paciente --"] + opciones_pacientes_nov,
    )
    if paciente_escogido != "-- Seleccionar paciente --":
      doc_preseleccionado = paciente_escogido.split(" - ")[0].strip()
  else:
    st.success("🎉 No hay novedades pendientes activas.")

  st.divider()
  doc_nov = st.text_input(
      "Digite la Cédula del Paciente a Gestionar:",
      value=doc_preseleccionado,
      key="doc_nov_input",
  ).strip()

  if doc_nov:
    p_data = run_query(
        "SELECT documento, nombre, plan, programa, zona, piso,"
        " alerta_permanente, observaciones_clinicas FROM pacientes WHERE"
        " documento = ?",
        (doc_nov,),
    )
    if p_data:
      p_doc, p_nom, p_plan, p_prog, p_zona, p_piso, p_alerta_f, p_obs_f = (
          p_data[0]
      )

      ttos_activos_nov = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC",
          (doc_nov,),
      )

      st.markdown(
          f"#### 👤 Paciente: {p_nom} ({p_doc}) · Zona `{p_zona}` ({p_piso})"
      )
      if str(p_alerta_f).strip():
        st.warning(f"📌 **ALERTA FIJA PERMANENTE:** `{p_alerta_f.strip()}`")
      if str(p_obs_f).strip():
        st.info(f"📝 **Observación Clínica General:** {p_obs_f.strip()}")

      tab_gestionar, tab_nuevo = st.tabs(
          ["🛎️ Novedades Pendientes y Gestión", "➕ Registrar Nueva Novedad"]
      )

      with tab_gestionar:
        novs_pend = run_query(
            "SELECT id, tipo_novedad, nivel_novedad, fecha_aplicacion,"
            " hora_aplicacion, detalle, responsable FROM registro_novedades"
            " WHERE documento = ? AND estado_novedad = 'Pendiente' ORDER BY id"
            " DESC",
            (doc_nov,),
        )
        if novs_pend:
          st.write(
              f"El paciente tiene **{len(novs_pend)} novedad(es)**. Puede"
              " resolverlas individualmente:"
          )
          for n in novs_pend:
            n_id, n_tipo, n_niv, n_fec, n_hor, n_det, n_resp = n
            with st.expander(
                f"📌 Novedad #{n_id}: {n_tipo} ({n_niv}) - {n_fec} {n_hor}",
                expanded=True,
            ):
              st.markdown(
                  f"**Detalle registrado:** {n_det}\n\n*Registrado por:* {n_resp}"
              )

              col_res1, col_res2 = st.columns(2)
              prof_ges = render_selectbox(
                  "Profesional que gestiona esta novedad (Firma):",
                  cat_gestores,
                  "",
                  key=f"prof_ges_{n_id}",
              )
              acc_ges = col_res2.text_input(
                  "Acción realizada:",
                  value="Novedad atendida satisfactoriamente.",
                  key=f"acc_ges_{n_id}",
              )

              if st.button(
                  f"✅ Resolver Únicamente esta Novedad (#{n_id})",
                  key=f"btn_res_{n_id}",
              ):
                if not prof_ges.strip():
                  st.error("Ingrese su nombre como responsable.")
                else:
                  fec_res = datetime.now().strftime("%d/%m/%Y %I:%M %p")
                  run_query(
                      """UPDATE registro_novedades SET estado_novedad = 'Gestionada', 
                                 fecha_gestion = ?, responsable_gestion = ? WHERE id = ?""",
                      (fec_res, prof_ges.strip(), n_id),
                      fetch=False,
                  )

                  restantes = run_query(
                      "SELECT COUNT(*) FROM registro_novedades WHERE documento"
                      " = ? AND estado_novedad = 'Pendiente'",
                      (doc_nov,),
                  )
                  if restantes and restantes[0][0] == 0 and ttos_activos_nov:
                    run_query(
                        """UPDATE tratamientos SET alerta_revisada = 'SI', 
                                   novedades = novedades || ' // [TODAS GESTIONADAS ' || ? || ' por ' || ? || ']' WHERE id = ?""",
                        (fec_res, prof_ges.strip(), ttos_activos_nov[0][0]),
                        fetch=False,
                    )
                  st.success(
                      f"✅ Novedad #{n_id} gestionada individualmente con"
                      " éxito."
                  )
                  st.rerun()
        else:
          st.success("No hay novedades pendientes para este paciente.")

      with tab_nuevo:
        col_nt1, col_nt2 = st.columns(2)
        tipo_nov = col_nt1.selectbox(
            "Categoría de Novedad:",
            [
                "Cambio de Horario / Reacomodo de Visita",
                "Visita Fallida / Reprogramación por no respuesta",
                "Novedad de Salida / Descanalización",
                "Modificación por Cambio de Orden Médica",
                "Paciente Ausente / No responde",
                "Familiar suspende dosis",
                "Otra Novedad de Ruta",
            ],
        )

        nivel_accion = col_nt2.selectbox(
            "Nivel de Impacto Operativo:",
            [
                "📢 Nivel 1: Informativa (Requiere ser gestionada para"
                " retirarse de la planilla)",
                "🕒 Nivel 2: Visita Puntual de un Solo Día (Automática, no"
                " genera alerta bloqueante)",
                "🔄 Nivel 3: Cambio Permanente de Todo el Ciclo (Automática sin"
                " reversa)",
            ],
        )

        if "Visita Fallida" in tipo_nov:
          st.markdown(
              "##### 🔄 Asistencia Rápida: Reprogramar Turno / Visita Fallida"
          )
          col_reprog1, col_reprog2 = st.columns(2)
          reprog_opcion = col_reprog1.radio(
              "Reprogramar para:",
              [
                  "Siguiente Turno (Turno posterior)",
                  "Mañana en el mismo horario",
              ],
              horizontal=True,
          )
          prof_rep = render_selectbox(
              "Profesional responsable de la reprogramación:", cat_gestores, ""
          )
          if st.button("🚀 Ejecutar Reprogramación Inmediata"):
            if not prof_rep.strip():
              st.error("Ingrese su nombre.")
            else:
              fec_hoy_str = datetime.now().strftime("%d/%m/%Y")
              hora_aut = datetime.now().strftime("%I:%M %p")
              nov_reprog_txt = (
                  f"[VISITA FALLIDA REPROGRAMADA para {reprog_opcion}]:"
                  f" Autoriza {prof_rep.strip()} (Reg: {hora_aut})"
              )
              if ttos_activos_nov:
                run_query(
                    """UPDATE tratamientos SET novedades=novedades || ' // ' || ? WHERE id=?""",
                    (nov_reprog_txt, ttos_activos_nov[0][0]),
                    fetch=False,
                )
              run_query(
                  """INSERT INTO registro_novedades (documento, tipo_novedad, nivel_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                             VALUES (?, ?, 'Informativa', ?, ?, ?, ?, 'Gestionada', ?)""",
                  (
                      doc_nov,
                      tipo_nov,
                      fec_hoy_str,
                      hora_aut,
                      f"Reprogramado por: {reprog_opcion}",
                      prof_rep.strip(),
                      datetime.now().isoformat(),
                  ),
                  fetch=False,
              )
              st.success(
                  "✅ Visita fallida reprogramada con éxito en la trazabilidad."
              )
              st.rerun()

        t_sel_para_nov = None
        if ttos_activos_nov and ("Nivel 2" in nivel_accion or "Nivel 3" in nivel_accion):
          st.markdown("#### 🕒 Tratamiento y Horarios Actuales de Visita")

          if len(ttos_activos_nov) > 1:
            opc_ttos = [
                f"ID #{t[0]}: {t[2]} ({t[3]}) - Vía {t[4]}"
                for t in ttos_activos_nov
            ]
            t_sel_idx = st.selectbox(
                "Seleccione el tratamiento al que aplica la novedad:", opc_ttos
            )
            t_id_escogido = int(t_sel_idx.split(":")[0].replace("ID #", ""))
            t_sel_para_nov = next(
                t for t in ttos_activos_nov if t[0] == t_id_escogido
            )
          else:
            t_sel_para_nov = ttos_activos_nov[0]

          st.markdown(
              f"""
            <div class="card-horarios-actuales">
                <span style="font-size: 0.85rem; font-weight: bold; color: #0033A0;">💊 TRATAMIENTO ACTUAL:</span> <strong>{t_sel_para_nov[2]}</strong> ({t_sel_para_nov[3]}) · Vía: <strong>{t_sel_para_nov[4]}</strong> · Acceso: <strong>{t_sel_para_nov[5]}</strong><br>
                <span style="font-size: 0.85rem; font-weight: bold; color: #0033A0;">⏱️ FRECUENCIA:</span> Cada {t_sel_para_nov[6]} horas por {t_sel_para_nov[7]} días · Fin: {datetime.fromisoformat(t_sel_para_nov[9]).strftime('%d/%m %I:%M %p')}<br>
                <hr style="margin: 8px 0; border: 0; border-top: 1px solid #C4D9F2;">
                <strong>🕒 HORARIOS DE VISITA ACTUALES:</strong><br>
                • <strong>Turno T1 (06H-14H):</strong> <code>{t_sel_para_nov[11]}</code><br>
                • <strong>Turno T2 (14H-22H):</strong> <code>{t_sel_para_nov[12]}</code><br>
                • <strong>Turno T3 (22H-06H):</strong> <code>{t_sel_para_nov[13]}</code>
            </div>
            """,
              unsafe_allow_html=True,
          )

        fec_n = st.date_input(
            "📅 Fecha a partir de la cual rige / aplica la novedad:",
            value=datetime.now().date(),
        )

        hora_puntual_n2 = None
        turno_afectado_n2 = "T2"
        hora_nueva_perm = None
        confirma_nivel3 = False

        if "Nivel 2" in nivel_accion:
          st.markdown(
              "##### 🕒 Configuración de la Visita Puntual para este Día"
          )
          col_p1, col_p2 = st.columns(2)
          hora_puntual_n2 = col_p1.time_input(
              "Nueva hora a visitar ese día específico (Ej: 17:00):",
              value=(datetime.now() + timedelta(hours=1)).time(),
          )
          h_p = hora_puntual_n2.hour
          if 6 <= h_p < 14:
            turno_afectado_n2 = "T1"
          elif 14 <= h_p < 22:
            turno_afectado_n2 = "T2"
          else:
            turno_afectado_n2 = "T3"
          col_p2.info(
              f"Turno afectado automáticamente: **{turno_afectado_n2}** (Visita"
              f" a las {hora_puntual_n2.strftime('%H:%M')})"
          )

        elif "Nivel 3" in nivel_accion:
          hora_nueva_perm = st.time_input(
              "Nueva hora fija de visita a partir de esa fecha (Ej: 20:00):",
              value=(datetime.now() + timedelta(hours=2)).time(),
          )

          if t_sel_para_nov:
            dt_simulado = datetime.combine(fec_n, hora_nueva_perm)
            try:
              mapa_sim = json.loads(t_sel_para_nov[20])
            except Exception:
              mapa_sim = {}
            frec_sim = safe_int(t_sel_para_nov[6], 8)
            dias_sim = safe_int(t_sel_para_nov[7], 5)
            ajuste_sim = safe_int(t_sel_para_nov[23], 0)

            n_fin_s, n_t1_s, n_t2_s, n_t3_s, _, _, _, _ = (
                calcular_tratamiento_mixto(
                    dt_simulado, dias_sim, frec_sim, mapa_sim, ajuste_sim
                )
            )

            st.markdown(
                f"""
                <div style="background-color: #FFF3CD; border: 1px solid #FFE69C; padding: 10px 14px; border-radius: 8px; font-size: 0.85rem; color: #000000; margin-bottom: 10px;">
                    <strong>🔄 HORARIOS QUE REGIRÁN A PARTIR DEL {fec_n.strftime('%d/%m/%Y')}:</strong><br>
                    • Turno T1 nuevo: <code>{n_t1_s}</code><br>
                    • Turno T2 nuevo: <code>{n_t2_s}</code><br>
                    • Turno T3 nuevo: <code>{n_t3_s}</code><br>
                    • Fecha fin recalculada: <strong>{n_fin_s.strftime('%d/%m/%Y %I:%M %p')}</strong>
                </div>
                """,
                unsafe_allow_html=True,
            )

          st.warning(
              f"⚠️ **ATENCIÓN:** Esta novedad modificará los turnos y la fecha"
              f" final a las **{hora_nueva_perm.strftime('%H:%M')}** a partir"
              f" del **{fec_n.strftime('%d/%m/%Y')}**."
          )
          confirma_nivel3 = st.checkbox(
              "¿Está seguro de aplicar este cambio definitivo?", value=False
          )

        default_detalle = ""
        if "Nivel 2" in nivel_accion and hora_puntual_n2:
          default_detalle = (
              f"Visita puntual el {fec_n.strftime('%d/%m/%Y')} a las"
              f" {hora_puntual_n2.strftime('%H:%M')} en {turno_afectado_n2}."
          )
        elif "Nivel 3" in nivel_accion and hora_nueva_perm:
          default_detalle = (
              f"Reprogramar permanentemente para las"
              f" {hora_nueva_perm.strftime('%H:%M')} a partir del"
              f" {fec_n.strftime('%d/%m/%Y')}."
          )

        det_n = st.text_area(
            "Descripción y Detalle de la Novedad (Obligatorio):",
            value=default_detalle,
        )
        resp_n = render_selectbox(
            "Profesional Responsable (Firma):", cat_gestores, "", key="prof_crea_nov"
        )

        if st.button("💾 Guardar Novedad"):
          if "Nivel 3" in nivel_accion and not confirma_nivel3:
            st.error(
                "❌ Debe marcar la casilla de confirmación para aplicar el"
                " cambio permanente."
            )
          elif not resp_n.strip():
            st.error("❌ Ingrese su nombre como responsable.")
          elif not det_n.strip():
            st.error("❌ Escriba el detalle de la novedad.")
          else:
            ahora_dt = datetime.now()
            hora_creacion_auto = ahora_dt.strftime("%I:%M %p")
            fecha_crea_str = ahora_dt.isoformat()
            fec_aplica_str = fec_n.strftime("%d/%m/%Y")

            estado_nov_inicial = (
                "Pendiente" if "Nivel 1" in nivel_accion else "Gestionada"
            )

            run_query(
                """INSERT INTO registro_novedades (documento, tipo_novedad, nivel_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    doc_nov,
                    tipo_nov,
                    nivel_accion.split(":")[0],
                    fec_aplica_str,
                    hora_creacion_auto,
                    det_n,
                    resp_n.strip(),
                    estado_nov_inicial,
                    fecha_crea_str,
                ),
                fetch=False,
            )

            t_objetivo = (
                t_sel_para_nov
                if t_sel_para_nov
                else (ttos_activos_nov[0] if ttos_activos_nov else None)
            )

            if t_objetivo:
              t_id_act = t_objetivo[0]
              es_para_hoy = fec_n <= ahora_dt.date()

              if "Nivel 3" in nivel_accion and hora_nueva_perm:
                nuevo_dt_ini = datetime.combine(
                    datetime.fromisoformat(t_objetivo[8]).date(),
                    hora_nueva_perm,
                )
                try:
                  mapa_ant = json.loads(t_objetivo[20])
                except Exception:
                  mapa_ant = {}
                frec_h = safe_int(t_objetivo[6], 8)
                dias_h = safe_int(t_objetivo[7], 5)
                ajuste_h = safe_int(t_objetivo[23], 0)

                n_fin, n_t1, n_t2, n_t3, _, tot_d, _, _ = (
                    calcular_tratamiento_mixto(
                        nuevo_dt_ini, dias_h, frec_h, mapa_ant, ajuste_h
                    )
                )

                if es_para_hoy:
                  nov_perm_txt = (
                      f"[CAMBIO PERMANENTE A"
                      f" {hora_nueva_perm.strftime('%H:%M')}]: {det_n} -"
                      f" Autoriza: {resp_n.strip()} (Reg: {hora_creacion_auto})"
                  )
                  run_query(
                      """UPDATE tratamientos SET fecha_inicio=?, fecha_fin=?, t1=?, t2=?, t3=?, 
                                   t1_futuro='', t2_futuro='', t3_futuro='', fecha_aplica_futuro='',
                                   novedades=novedades || ' // ' || ? WHERE id=?""",
                      (
                          nuevo_dt_ini.isoformat(),
                          n_fin.isoformat(),
                          n_t1,
                          n_t2,
                          n_t3,
                          nov_perm_txt,
                          t_id_act,
                      ),
                      fetch=False,
                  )
                  st.success("✅ Horarios actualizados en todo el ciclo.")
                else:
                  nov_perm_txt = (
                      f"[CAMBIO PROGRAMADO PARA {fec_aplica_str} a"
                      f" {hora_nueva_perm.strftime('%H:%M')}]: {det_n} -"
                      f" Autoriza: {resp_n.strip()} (Reg: {hora_creacion_auto})"
                  )
                  run_query(
                      """UPDATE tratamientos SET t1_futuro=?, t2_futuro=?, t3_futuro=?, fecha_aplica_futuro=?, 
                                   novedades=novedades || ' // ' || ? WHERE id=?""",
                      (n_t1, n_t2, n_t3, fec_n.isoformat(), nov_perm_txt, t_id_act),
                      fetch=False,
                  )
                  st.success(
                      f"✅ Cambio programado para el {fec_aplica_str}. Hoy el"
                      " paciente continúa con sus horarios normales."
                  )

              elif "Nivel 2" in nivel_accion and hora_puntual_n2:
                nov_punt_txt = (
                    f"[VISITA PUNTUAL EL {fec_aplica_str} a las"
                    f" {hora_puntual_n2.strftime('%H:%M')}]: {det_n} - Autoriza:"
                    f" {resp_n.strip()} (Reg: {hora_creacion_auto})"
                )
                run_query(
                    """UPDATE tratamientos SET puntual_fecha=?, puntual_turno=?, puntual_hora=?, 
                               novedades=novedades || ' // ' || ? WHERE id=?""",
                    (
                        fec_n.isoformat(),
                        turno_afectado_n2,
                        hora_puntual_n2.strftime("%H:%M"),
                        nov_punt_txt,
                        t_id_act,
                    ),
                    fetch=False,
                )
                st.success("✅ Visita puntual aplicada y guardada.")

              else:
                nov_inf_txt = (
                    f"📢 NOVEDAD INFORMATIVA [{fec_aplica_str}]: {det_n} -"
                    f" Reporta: {resp_n.strip()} (Reg: {hora_creacion_auto})"
                )
                run_query(
                    """UPDATE tratamientos SET alerta_revisada='NO', 
                               novedades=novedades || ' // ' || ? WHERE id=?""",
                    (nov_inf_txt, t_id_act),
                    fetch=False,
                )
                st.success(
                    "✅ Novedad informativa guardada y visible como alerta en"
                    " planilla."
                )

            st.rerun()

# ----------------- SECCIÓN 3: CENSO Y PLANILLA EN VIVO (RENDER HTML SEGURO) -----------------
elif menu == "📊 Censo y Planilla en Vivo":
  col_vista1, col_vista2 = st.columns([3, 1])
  filtro_estado = col_vista2.selectbox(
      "Filtrar por Estado:", ["Activos", "Completados / Descanalizados", "Todos"]
  )

  condicion_estado = "WHERE t.estado = 'Activo'"
  if filtro_estado == "Completados / Descanalizados":
    condicion_estado = (
        "WHERE t.estado IN ('Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"
    )
  elif filtro_estado == "Todos":
    condicion_estado = (
        "WHERE t.estado IN ('Activo', 'Completado', 'Cambiado por Orden Médica', 'Suspendido Médicamente')"
    )

  query_general = f"""
        SELECT 
            p.documento AS 'Documento',
            p.nombre AS 'Nombre y Apellido',
            p.plan AS 'Plan',
            p.programa AS 'Programa',
            p.municipio AS 'Municipio',
            p.barrio AS 'Barrio',
            p.piso AS 'Piso',
            p.zona AS 'Zona',
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
            COALESCE(t.t1_futuro, '') AS 'T1_Futuro',
            COALESCE(t.t2_futuro, '') AS 'T2_Futuro',
            COALESCE(t.t3_futuro, '') AS 'T3_Futuro',
            COALESCE(t.fecha_aplica_futuro, '') AS 'Fecha_Futuro',
            COALESCE(t.puntual_fecha, '') AS 'Puntual_Fecha',
            COALESCE(t.puntual_turno, '') AS 'Puntual_Turno',
            COALESCE(t.puntual_hora, '') AS 'Puntual_Hora',
            COALESCE(t.novedades, '') AS 'NOVEDAD',
            COALESCE(t.estado, 'Sin TTO') AS 'Estado',
            COALESCE(t.descanalizado, 'NO') AS 'Descanalizado',
            COALESCE(t.alerta_revisada, 'NO') AS 'Alerta Revisada'
        FROM pacientes p
        JOIN tratamientos t ON p.documento = t.documento
        {condicion_estado}
        ORDER BY p.zona ASC, p.nombre ASC
    """

  conn = sqlite3.connect(DB_FILE)
  df_general = pd.read_sql_query(query_general, conn)
  conn.close()

  if not df_general.empty:
    hoy_date_str = datetime.now().date().isoformat()

    for idx, row in df_general.iterrows():
      t1_v, t2_v, t3_v = row["T1"], row["T2"], row["T3"]
      f_fut = row.get("Fecha_Futuro", "")
      if f_fut and f_fut <= hoy_date_str:
        t1_v = row.get("T1_Futuro", t1_v)
        t2_v = row.get("T2_Futuro", t2_v)
        t3_v = row.get("T3_Futuro", t3_v)

      p_fec = row.get("Puntual_Fecha", "")
      p_tur = row.get("Puntual_Turno", "")
      p_hor = row.get("Puntual_Hora", "")
      if p_fec == hoy_date_str and p_hor:
        if p_tur == "T1":
          t1_v = f"{p_hor}H (PUNTUAL HOY)"
        elif p_tur == "T2":
          t2_v = f"{p_hor}H (PUNTUAL HOY)"
        elif p_tur == "T3":
          t3_v = f"{p_hor}H (PUNTUAL HOY)"

      df_general.at[idx, "T1"] = t1_v
      df_general.at[idx, "T2"] = t2_v
      df_general.at[idx, "T3"] = t3_v

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

    df_general["Alerta Fija"] = df_general["Alerta_Permanente_Texto"].apply(
        lambda txt: "SI" if txt.strip() else "NO"
    )

    c_f1, c_f2, c_f3, c_f4, c_f5, c_f6 = st.columns(6)
    filtro_zona = c_f1.multiselect(
        "Zona", options=df_general["Zona"].dropna().unique()
    )
    filtro_plan = c_f2.multiselect(
        "Plan", options=df_general["Plan"].dropna().unique()
    )
    filtro_muni = c_f3.multiselect(
        "Municipio", options=df_general["Municipio"].dropna().unique()
    )
    filtro_via = c_f4.multiselect(
        "Vía", options=df_general["Vía"].dropna().unique()
    )
    filtro_acceso = c_f5.multiselect(
        "Tipo Acceso", options=df_general["Tipo Acceso"].dropna().unique()
    )
    filtro_alerta = c_f6.multiselect(
        "Alertas", options=df_general["Estado Tratamiento"].dropna().unique()
    )

    df_filtrado = df_general.copy()
    if filtro_zona:
      df_filtrado = df_filtrado[df_filtrado["Zona"].isin(filtro_zona)]
    if filtro_plan:
      df_filtrado = df_filtrado[df_filtrado["Plan"].isin(filtro_plan)]
    if filtro_muni:
      df_filtrado = df_filtrado[df_filtrado["Municipio"].isin(filtro_muni)]
    if filtro_via:
      df_filtrado = df_filtrado[df_filtrado["Vía"].isin(filtro_via)]
    if filtro_acceso:
      df_filtrado = df_filtrado[df_filtrado["Tipo Acceso"].isin(filtro_acceso)]
    if filtro_alerta:
      df_filtrado = df_filtrado[
          df_filtrado["Estado Tratamiento"].isin(filtro_alerta)
      ]

    cols_ordenadas = [
        "Alerta Fija",
        "Documento",
        "Nombre y Apellido",
        "Plan",
        "Programa",
        "Municipio",
        "Barrio",
        "Piso",
        "Zona",
        "Aislamiento",
        "Tratamiento",
        "Dosis",
        "Vía",
        "Tipo Acceso",
        "Frec (h)",
        "Días",
        "Fecha Inicio",
        "Fecha Fin",
        "Fecha Retiro Cateter",
        "T1",
        "T2",
        "T3",
        "NOVEDAD",
        "Estado",
        "Estado Tratamiento",
        "Alerta_Permanente_Texto",
    ]
    df_filtrado = df_filtrado[cols_ordenadas]

    df_filtrado.reset_index(drop=True, inplace=True)
    df_filtrado.index = df_filtrado.index + 1
    df_filtrado.index.name = "N°"

    total_filtrados = len(df_filtrado[df_filtrado["Estado"] == "Activo"])
    total_picc = len(
        df_filtrado[
            (df_filtrado["Tipo Acceso"].str.upper() == "PICC")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_npt = len(
        df_filtrado[
            df_filtrado["Vía"].str.upper().str.contains("NPT")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_be = len(
        df_filtrado[
            df_filtrado["Vía"].str.upper().str.contains("BE")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_aislados = len(
        df_filtrado[
            (df_filtrado["Aislamiento"] == "SI")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_novedades = len(
        df_filtrado[
            df_filtrado["Estado Tratamiento"].str.contains("NOVEDAD ACTIVA")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_descanalizar = len(
        df_filtrado[
            df_filtrado["Estado Tratamiento"].str.contains("FALTA 1 DOSIS")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    zona_txt = f" ({', '.join(filtro_zona)})" if filtro_zona else " TOTAL"
    m1.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_filtrados}</div><div"
        f" class='metric-lbl'>ACTIVOS EN ZONA{zona_txt}</div></div>",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_picc}</div><div class='metric-lbl'>🩸 CON"
        " CATÉTER PICC</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_npt}</div><div class='metric-lbl'>🥣 CON"
        " NPT</div></div>",
        unsafe_allow_html=True,
    )
    m4.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_be}</div><div class='metric-lbl'>💉"
        " INFUSIÓN BE / ANALGESIA</div></div>",
        unsafe_allow_html=True,
    )
    m5.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_aislados}</div><div class='metric-lbl'>☣️"
        " AISLADOS</div></div>",
        unsafe_allow_html=True,
    )
    m6.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_novedades}</div><div class='metric-lbl'>📢"
        " NOVEDADES ACTIVAS</div></div>",
        unsafe_allow_html=True,
    )
    m7.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_descanalizar}</div><div"
        " class='metric-lbl'>⚠️ PENDIENTE RETIRO</div></div>",
        unsafe_allow_html=True,
    )

    st.markdown("##### 🎨 Convenciones de Color en la Planilla")
    c_c1, c_c2, c_c3, c_c4 = st.columns(4)
    with c_c1:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_PICC};'>🩸"
          " Catéter PICC</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_NPT};'>🥣"
          " Nutrición Parenteral (NPT)</div>",
          unsafe_allow_html=True,
      )
    with c_c2:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_BE};'>🟩"
          " Bomba Elastomérica (BE)</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color:"
          f" {COLOR_ANALGESIA};'>🟪 BE Analgesia / Dolor</div>",
          unsafe_allow_html=True,
      )
    with c_c3:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_TB};'>🟧"
          " Tuberculosis (TB)</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_PUSH};'>🟦"
          " Vía PUSH</div>",
          unsafe_allow_html=True,
      )
    with c_c4:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_AISLA};'>🟣"
          " Paciente Aislado</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_FALTA1};'>⚠️"
          " Falta 1 Dosis / Salida</div>",
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)

    # ----------------- TABLA HTML NATIVA CON TOOLTIP FLOTANTE (SEGURO CONTRA ERRORES) -----------------
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

      fin_formateada = (
          datetime.fromisoformat(r["Fecha Fin"]).strftime("%d/%m %I:%M %p")
          if r["Fecha Fin"] != "-"
          else "-"
      )
      retiro_formateado = (
          datetime.fromisoformat(r["Fecha Retiro Cateter"]).strftime(
              "%d/%m %I:%M %p"
          )
          if str(r["Fecha Retiro Cateter"]).strip()
          else "N/A"
      )

      html_filas += f"""
            <tr style="background-color: {bg_color};">
                <td><strong>{num_idx}</strong></td>
                <td>{celda_alerta_html}</td>
                <td><strong>{r['Documento']}</strong></td>
                <td style="text-align: left;">{r['Nombre y Apellido']}</td>
                <td>{r['Plan']}</td>
                <td>{r['Programa']}</td>
                <td>{r['Municipio']}</td>
                <td>{r['Barrio']}</td>
                <td>{r['Piso']}</td>
                <td><strong>{r['Zona']}</strong></td>
                <td>{r['Aislamiento']}</td>
                <td><strong>{r['Tratamiento']}</strong></td>
                <td>{r['Dosis']}</td>
                <td>{r['Vía']}</td>
                <td>{r['Tipo Acceso']}</td>
                <td>C/{r['Frec (h)']}h</td>
                <td>{r['Días']}</td>
                <td>{fin_formateada}</td>
                <td style="background-color: #FFF3CD; font-weight: bold;">{retiro_formateado}</td>
                <td><code>{r['T1']}</code></td>
                <td><code>{r['T2']}</code></td>
                <td><code>{r['T3']}</code></td>
                <td><strong>{r['Estado Tratamiento']}</strong></td>
                <td style="text-align: left; font-size: 0.78rem;">{r['NOVEDAD']}</td>
            </tr>
            """

    tabla_html_final = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        .tabla-censo-container {{
            width: 100%;
            height: 500px;
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
            width: 250px;
            background-color: #212529;
            color: #FFFFFF;
            text-align: left;
            border-radius: 8px;
            padding: 10px 14px;
            position: absolute;
            z-index: 999999;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            margin-bottom: 8px;
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
            top: 100%;
            left: 50%;
            margin-left: -6px;
            border-width: 6px;
            border-style: solid;
            border-color: #212529 transparent transparent transparent;
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
                            <th>Plan</th>
                            <th>Programa</th>
                            <th>Municipio</th>
                            <th>Barrio</th>
                            <th>Piso</th>
                            <th>Zona</th>
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
                            <th>Novedades / Alertas</th>
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
    df_descarga["Alerta Fija"] = df_descarga["Alerta_Permanente_Texto"].apply(
        lambda t: "SI" if str(t).strip() else "NO"
    )
    cols_excel = [
        "Alerta Fija",
        "Documento",
        "Nombre y Apellido",
        "Plan",
        "Programa",
        "Municipio",
        "Barrio",
        "Piso",
        "Zona",
        "Aislamiento",
        "Alerta_Permanente_Texto",
        "Tratamiento",
        "Dosis",
        "Vía",
        "Tipo Acceso",
        "Frec (h)",
        "Días",
        "Fecha Fin",
        "Fecha Retiro Cateter",
        "T1",
        "T2",
        "T3",
        "Estado Tratamiento",
        "NOVEDAD",
    ]
    df_descarga = df_descarga[cols_excel]
    df_descarga.rename(
        columns={"Alerta_Permanente_Texto": "Detalle Alerta Fija"}, inplace=True
    )

    excel_coloreado = exportar_excel_con_colores(df_descarga)

    st.download_button(
        label="📥 Descargar Planilla en Excel con Colores (.xlsx)",
        data=excel_coloreado,
        file_name=(
            f"Planilla_SENC_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.info("No hay registros activos.")

# ----------------- SECCIÓN 4: RESPALDO GENERAL EN EXCEL -----------------
elif menu == "💾 Respaldo General en Excel":
  st.subheader("💾 Copia de Seguridad Consolidada en Excel (Con Colores)")
  opcion_backup = st.radio(
      "Modalidad de Respaldo:",
      [
          "1. Planilla Consolidada Limpia (Solo pacientes y tratamientos"
          " ACTIVOS)",
          "2. Planilla Completa con Historial y Base Permanente de Pacientes"
          " (Auditoría)",
      ],
      index=0,
  )

  if st.button("📊 Generar y Descargar Archivo Excel Coloreado"):
    conn = sqlite3.connect(DB_FILE)
    if "Solo pacientes y tratamientos ACTIVOS" in opcion_backup:
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
                ORDER BY p.zona ASC, p.nombre ASC
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
                ORDER BY p.zona ASC, p.documento ASC, t.id DESC
            """
      nombre_archivo = f"HISTORIAL_AUDITORIA_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    df_b = pd.read_sql_query(query_b, conn)
    conn.close()

    excel_b_coloreado = exportar_excel_con_colores(df_b)

    st.download_button(
        label="⬇️ Descargar Archivo Excel Coloreado (.xlsx)",
        data=excel_b_coloreado,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ----------------- SECCIÓN 5: ANALÍTICA Y GRÁFICOS INTERACTIVOS (SVG NATIVO) -----------------
elif menu == "📈 Analítica y Gráficos del Censo":
  st.subheader("📈 Tablero Analítico y Gráficos Circulares SENC")
  st.caption(
      "Consulte métricas por rango de fechas de ingreso y distribución en"
      " gráficos circulares (pastel) por estado, vía, acceso y zona."
  )

  conn = sqlite3.connect(DB_FILE)
  df_pacientes_all = pd.read_sql_query("SELECT * FROM pacientes", conn)
  df_ttos_all = pd.read_sql_query("SELECT * FROM tratamientos", conn)
  conn.close()

  if not df_pacientes_all.empty:
    col_fec1, col_fec2 = st.columns(2)
    f_desde = col_fec1.date_input(
        "Fecha Ingreso Desde:",
        value=(datetime.now() - timedelta(days=30)).date(),
    )
    f_hasta = col_fec2.date_input(
        "Fecha Ingreso Hasta:", value=datetime.now().date()
    )

    df_pacientes_all["fecha_ingreso_dt"] = pd.to_datetime(
        df_pacientes_all["fecha_ingreso"], errors="coerce"
    )
    df_p_rango = df_pacientes_all[
        (df_pacientes_all["fecha_ingreso_dt"].dt.date >= f_desde)
        & (df_pacientes_all["fecha_ingreso_dt"].dt.date <= f_hasta)
    ]

    total_ingresos_rango = len(df_p_rango)
    activos_rango = len(df_p_rango[df_p_rango["estado_paciente"] == "Activo"])
    egresados_rango = len(
        df_p_rango[df_p_rango["estado_paciente"] == "Egresado"]
    )

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Ingresos en Periodo", f"{total_ingresos_rango} Pacientes")
    col_m2.metric("Pacientes Activos", f"{activos_rango} Activos")
    col_m3.metric("Pacientes Egresados / Alta", f"{egresados_rango} Egresados")
    pct_alta = (
        round((egresados_rango / total_ingresos_rango) * 100, 1)
        if total_ingresos_rango > 0
        else 0
    )
    col_m4.metric("% Tasa de Egreso / Alta", f"{pct_alta}%")

    st.markdown("---")

    col_g1, col_g2 = st.columns(2)

    with col_g1:
      data_estado = [activos_rango, egresados_rango]
      labels_estado = ["Activos", "Egresados"]
      colores_estado = ["#0033A0", "#6C757D"]
      st.markdown(
          generar_pie_chart_svg(
              data_estado,
              labels_estado,
              colores_estado,
              "Distribución por Estado de Pacientes",
          ),
          unsafe_allow_html=True,
      )

    with col_g2:
      df_ttos_activos = df_ttos_all[df_ttos_all["estado"] == "Activo"]
      if not df_ttos_activos.empty:
        vias_cuenta = df_ttos_activos["via_administracion"].value_counts()
        palette_vias = [
            "#0033A0",
            "#00AEC7",
            "#28A745",
            "#FFC107",
            "#DC3545",
            "#6F42C1",
            "#FD7E14",
            "#20C997",
        ]
        st.markdown(
            generar_pie_chart_svg(
                list(vias_cuenta.values),
                list(vias_cuenta.index),
                palette_vias[: len(vias_cuenta)],
                "Distribución por Vía / Terapia",
            ),
            unsafe_allow_html=True,
        )
      else:
        st.info("No hay tratamientos activos.")

    st.markdown("<br>", unsafe_allow_html=True)
    col_g3, col_g4 = st.columns(2)

    with col_g3:
      if not df_ttos_activos.empty:
        acc_cuenta = df_ttos_activos["tipo_acceso"].value_counts()
        palette_acc = ["#17A2B8", "#E83E8C", "#6C757D", "#FFC107"]
        st.markdown(
            generar_pie_chart_svg(
                list(acc_cuenta.values),
                list(acc_cuenta.index),
                palette_acc[: len(acc_cuenta)],
                "Distribución por Acceso Vascular",
            ),
            unsafe_allow_html=True,
        )

    with col_g4:
      zonas_cuenta = df_pacientes_all[
          df_pacientes_all["estado_paciente"] == "Activo"
      ]["zona"].value_counts()
      if not zonas_cuenta.empty:
        palette_zonas = [
            "#0033A0",
            "#00AEC7",
            "#20C997",
            "#28A745",
            "#FFC107",
            "#FD7E14",
            "#DC3545",
            "#6610F2",
            "#6F42C1",
            "#E83E8C",
        ]
        st.markdown(
            generar_pie_chart_svg(
                list(zonas_cuenta.values),
                list(zonas_cuenta.index),
                palette_zonas[: len(zonas_cuenta)],
                "Distribución por Zona Operativa",
            ),
            unsafe_allow_html=True,
        )

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
