import io
import json
import os
import sqlite3
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
import streamlit as st

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
        color: #6C757D;
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
    </style>
""",
    unsafe_allow_html=True,
)


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
            estado_paciente TEXT DEFAULT 'Activo',
            usuario_registro TEXT
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
            t1 TEXT,
            t2 TEXT,
            t3 TEXT,
            distribucion_admin TEXT DEFAULT '{}',
            visita_educacion TEXT DEFAULT '',
            dosis_base_fijas INTEGER DEFAULT 0,
            dosis_perdidas INTEGER DEFAULT 0,
            tipo_unidad_ajuste TEXT DEFAULT 'Unidosis', -- 'Unidosis' o 'BE Completa'
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
            fecha_retiro_cateter TEXT,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

  c.execute("""
        CREATE TABLE IF NOT EXISTS registro_novedades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT NOT NULL,
            tipo_novedad TEXT NOT NULL,
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

  # Migraciones preventivas
  c.execute("PRAGMA table_info(tratamientos)")
  cols_t = [col[1] for col in c.fetchall()]
  mig_trat = {
      "dosis_base_fijas": "INTEGER DEFAULT 0",
      "tipo_unidad_ajuste": "TEXT DEFAULT 'Unidosis'",
      "motivo_ajuste_dosis": "TEXT DEFAULT ''",
      "npt_dias_semana": "TEXT DEFAULT '[]'",
      "npt_horas_infusion": "INTEGER DEFAULT 12",
      "npt_hora_desconexion": "TEXT DEFAULT ''",
      "be_unidosis_puente": "INTEGER DEFAULT 0",
      "be_total_bombas": "INTEGER DEFAULT 0",
      "be_dosis_final_cierre": "INTEGER DEFAULT 0",
  }
  for col, defn in mig_trat.items():
    if col not in cols_t:
      c.execute(f"ALTER TABLE tratamientos ADD COLUMN {col} {defn}")

  conn.commit()
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


# ----------------- CÁLCULOS CLÍNICOS EXACTOS -----------------
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


# CONCILIACIÓN EXACTA: BE + UNIDOSIS
def calcular_esquema_be_exacto(
    total_unidosis_prescritas, unidosis_puente_previas, c_horas
):
  dosis_por_bomba = 4 if c_horas == 6 else 3
  dosis_restantes = max(0, total_unidosis_prescritas - unidosis_puente_previas)
  bombas_completas = dosis_restantes // dosis_por_bomba
  dosis_cierre = dosis_restantes % dosis_por_bomba
  return bombas_completas, dosis_cierre, dosis_por_bomba


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

  nov_up = str(novedades).upper() if novedades else ""

  if alerta_revisada != "SI":
    if (
        "NOVEDAD PENDIENTE" in nov_up
        or "VISITA PUNTUAL" in nov_up
        or "SALIDA" in nov_up
    ):
      return "📢 NOVEDAD ACTIVA (POR GESTIONAR)"

  if alerta_revisada == "SI":
    return "✅ NOVEDAD / ALERTA GESTIONADA"

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


# ----------------- PALETA DE COLORES PASTEL DISTINTIVOS -----------------
# Colores hex suaves para la web
COLOR_PICC = "#FFD1D1"  # Rosado suave
COLOR_NPT = "#FFF3CD"  # Crema ámbar suave
COLOR_BE = "#D1E7DD"  # Verde menta suave
COLOR_ANALGESIA = "#E2D9F3"  # Lavanda suave
COLOR_TB = "#FFE5D0"  # Melocotón suave
COLOR_PUSH = "#CFF4FC"  # Azul hielo / turquesa suave
COLOR_AISLA = "#E8D7F1"  # Lila suave
COLOR_FALTA1 = "#FFE0B2"  # Naranja claro advertencia


def estilo_filas(row):
  estado = str(row.get("Estado", ""))
  via = str(row.get("Vía", "")).upper()
  acceso = str(row.get("Tipo Acceso", "")).upper()
  alerta = str(row.get("Estado Tratamiento", ""))
  aislamiento = str(row.get("Aislamiento", "")).upper()

  if "DESCANALIZADO" in alerta or estado == "Completado":
    return ["color: #6C757D; font-style: italic;"] * len(row)

  if "NOVEDAD ACTIVA" in alerta:
    return [
        "background-color: #ffe0b2; color: #b71c1c; font-weight: bold;"
    ] * len(row)

  if "FALTA 1 DOSIS" in alerta:
    return [
        f"background-color: {COLOR_FALTA1}; color: #b71c1c; font-weight: bold;"
    ] * len(row)

  if acceso == "PICC":
    return [
        f"background-color: {COLOR_PICC}; color: #721c24; font-weight: bold;"
    ] * len(row)

  if "NPT" in via:
    return [
        f"background-color: {COLOR_NPT}; color: #664d03; font-weight: bold;"
    ] * len(row)
  elif "ANALGESIA" in via:
    return [
        f"background-color: {COLOR_ANALGESIA}; color: #4a148c; font-weight:"
        " bold;"
    ] * len(row)
  elif "BE" in via:
    return [
        f"background-color: {COLOR_BE}; color: #0f5132; font-weight: bold;"
    ] * len(row)
  elif "TUBERCULOSIS" in via or "TB" in via:
    return [
        f"background-color: {COLOR_TB}; color: #842029; font-weight: bold;"
    ] * len(row)
  elif via == "PUSH":
    return [
        f"background-color: {COLOR_PUSH}; color: #055160; font-weight: bold;"
    ] * len(row)

  if aislamiento == "SI":
    return [
        f"background-color: {COLOR_AISLA}; color: #4a148c; font-weight: bold;"
    ] * len(row)

  if "PRÓXIMO A FIN" in alerta:
    return [
        "background-color: #fff3cd; color: #856404; font-weight: bold;"
    ] * len(row)

  return [""] * len(row)


# Función para aplicar estilos con color directamente en el archivo Excel descargado
def exportar_excel_con_colores(df):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Planilla_SURA")
    workbook = writer.book
    worksheet = writer.sheets["Planilla_SURA"]

    # Fills para openpyxl (hex sin #)
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

      if fill_row:
        for col_idx in range(1, len(df.columns) + 1):
          worksheet.cell(row=row_idx + 2, column=col_idx).fill = fill_row

  return output.getvalue()


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

# Variable de sesión para controlar si se está editando un tratamiento existente con el lápiz
if "id_tto_editando" not in st.session_state:
  st.session_state.id_tto_editando = None

# ----------------- SECCIÓN 1: GESTIÓN DE PACIENTES Y FORMULACIÓN -----------------
if menu == "📋 Gestión de Pacientes":
  st.subheader("🔍 Localizador de Pacientes")
  doc_busqueda = st.text_input(
      "Documento de Identidad del Paciente:", ""
  ).strip()

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
    paciente_data = run_query(
        "SELECT * FROM pacientes WHERE documento = ?", (doc_busqueda,)
    )

    with st.expander(
        "👤 Información Demográfica del Paciente (Base Permanente)",
        expanded=not bool(paciente_data),
    ):
      col1, col2, col3 = st.columns(3)

      if paciente_data:
        p_doc, p_nom, p_plan, p_prog, p_piso, p_zona = paciente_data[0][0:6]
        p_muni = (
            paciente_data[0][6]
            if len(paciente_data[0]) > 6 and paciente_data[0][6]
            else "Medellín"
        )
        p_barrio = (
            paciente_data[0][7]
            if len(paciente_data[0]) > 7 and paciente_data[0][7]
            else ""
        )
        p_aisla = (
            paciente_data[0][8]
            if len(paciente_data[0]) > 8 and paciente_data[0][8]
            else "NO"
        )
        p_tipo_aisla = (
            paciente_data[0][9]
            if len(paciente_data[0]) > 9 and paciente_data[0][9]
            else "Ninguno"
        )
        p_estado_gral = (
            paciente_data[0][10]
            if len(paciente_data[0]) > 10 and paciente_data[0][10]
            else "Activo"
        )
        p_user = (
            paciente_data[0][11]
            if len(paciente_data[0]) > 11 and paciente_data[0][11]
            else ""
        )

        nombre = col1.text_input("Nombre y Apellido", value=p_nom)
        plan = col2.selectbox(
            "Plan",
            ["POS", "Póliza", "ARL"],
            index=["POS", "Póliza", "ARL"].index(p_plan),
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

        usuario_edita = st.text_input(
            "Tu Nombre / Profesional responsable:",
            value=p_user if p_user else "",
        )

        if st.button("Actualizar Ficha Paciente"):
          if usuario_edita:
            run_query(
                """UPDATE pacientes SET nombre=?, plan=?, programa=?, piso=?, zona=?, 
                           municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, estado_paciente='Activo', usuario_registro=? WHERE documento=?""",
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
                    usuario_edita,
                    doc_busqueda,
                ),
                fetch=False,
            )
            st.success(
                "Ficha demográfica guardada y reactivada en la base permanente."
            )
          else:
            st.error("Debes ingresar tu nombre como responsable.")
      else:
        st.info("Paciente no encontrado. Ingrese los datos para admitirlo:")
        nombre = col1.text_input("Nombre y Apellido")
        plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"])
        programa = col3.selectbox("Programa", programas_disponibles, index=0)

        col4, col5, col6, col7 = st.columns(4)
        municipio = col4.selectbox("Municipio", municipios_disponibles)
        barrio = col5.text_input("Barrio")
        piso = col6.selectbox("Piso", pisos)
        zona = col7.selectbox("Zona", zonas)

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

        usuario_crea = st.text_input("Tu Nombre / Profesional que registra:")

        if st.button("Admitir Paciente a la Base"):
          if nombre and usuario_crea:
            run_query(
                """INSERT INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, estado_paciente, usuario_registro) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?)""",
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
                    usuario_crea,
                ),
                fetch=False,
            )
            st.success("Paciente admitido a la base permanente.")
            st.rerun()
          else:
            st.error("Diligencie el nombre del paciente y su nombre.")

    if paciente_data:
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

          with st.expander(
              "✅ Egresar Paciente / Dar Alta Médica (Conserva historial en"
              " base de datos)"
          ):
            c_egr1, c_egr2 = st.columns(2)
            prof_egreso = c_egr1.text_input(
                "Profesional que autoriza el alta:", key="prof_alta"
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

      # ----------------- LISTADO DE TRATAMIENTOS CON BOTÓN DE LÁPIZ (✏️ MODIFICAR) -----------------
      st.markdown("### 💊 Tratamientos Activos del Paciente")
      if ttos_activos:
        for t in ttos_activos:
          alerta_t = evaluar_alerta_fin(t[9], t[6], t[20], t[21], t[18], t[22])
          col_info, col_btn = st.columns([5, 1])
          with col_info:
            st.info(
                f"**Medicamento:** {t[2]} ({t[3]}) · **Vía:** {t[4]} ·"
                f" **Acceso:** {t[5]} · **Frec:** C/{t[6]}h · **Fin:**"
                f" {datetime.fromisoformat(t[9]).strftime('%d/%m %I:%M %p')}\n\n*Turnos:*"
                f" T1: `{t[10]}` | T2: `{t[11]}` | T3: `{t[12]}` · **Estado:**"
                f" `{alerta_t}`"
            )
          with col_btn:
            if st.button(f"✏️ Modificar", key=f"btn_edit_{t[0]}"):
              st.session_state.id_tto_editando = t[0]
              st.rerun()
      else:
        st.info("El paciente no tiene tratamientos activos en este momento.")

      # Cargar datos si se está editando con el lápiz
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
          if st.button("❌ Cancelar Edición / Formular Nuevo"):
            st.session_state.id_tto_editando = None
            st.rerun()

      # ----------------- FORMULARIO: FORMULACIÓN O EDICIÓN COMPLETA -----------------
      st.markdown(
          "### ➕ Formular Nuevo Tratamiento o Guardar Cambios del Seleccionado"
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
          "PUSH",
      ]
      accesos = ["Periférico", "PICC", "SC", "Ninguno"]

      # Valores precargados si se edita
      val_med = tto_edit_data[2] if tto_edit_data else ""
      val_dos = tto_edit_data[3] if tto_edit_data else ""
      val_via = (
          tto_edit_data[4]
          if (tto_edit_data and tto_edit_data[4] in vias)
          else "IV"
      )
      val_acc = (
          tto_edit_data[5]
          if (tto_edit_data and tto_edit_data[5] in accesos)
          else "Periférico"
      )
      val_frec = safe_int(tto_edit_data[6], 8) if tto_edit_data else 8
      val_dias = safe_int(tto_edit_data[7], 5) if tto_edit_data else 5
      val_ini = (
          datetime.fromisoformat(tto_edit_data[8])
          if tto_edit_data
          else datetime.now()
      )

      col_t1, col_t2, col_t3, col_t4 = st.columns([2, 1, 1, 1])
      t_nombre = col_t1.text_input("Medicamento / Terapia:", value=val_med)
      t_dosis = col_t2.text_input("Dosis (Ej: 1g, 500mg, 1 Bolsa):", value=val_dos)
      t_via = col_t3.selectbox("Vía de Adm:", vias, index=vias.index(val_via))
      t_acceso = col_t4.selectbox(
          "Acceso:",
          accesos,
          index=accesos.index(val_acc) if val_acc in accesos else 0,
      )

      # ----------------- PROTOCOLO NPT: DÍAS ESPECÍFICOS Y DESCONEXIÓN -----------------
      npt_dias_semana_json = "[]"
      npt_horas_infusion = 12
      npt_hora_desconexion_str = ""

      if t_via == "NPT":
        st.markdown(
            "#### 🥣 Protocolo NPT: Programación de Infusión y Retiro"
        )
        col_npt1, col_npt2 = st.columns(2)
        dias_def = ["Lunes", "Martes", "Miércoles"]
        if tto_edit_data and tto_edit_data[19]:
          try:
            dias_def = json.loads(tto_edit_data[19])
          except Exception:
            dias_def = ["Lunes", "Martes", "Miércoles"]

        dias_semana_npt = col_npt1.multiselect(
            "Días de infusión en la semana (Descansa los demás):",
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
        horas_opc = [12, 14, 16, 18, 20, 24]
        h_inf_def = (
            safe_int(tto_edit_data[20], 12) if tto_edit_data else 12
        )
        npt_horas_infusion = col_npt2.selectbox(
            "Horas de infusión programadas:",
            horas_opc,
            index=horas_opc.index(h_inf_def) if h_inf_def in horas_opc else 0,
        )

      # ----------------- FRECUENCIA Y DÍAS -----------------
      col_f1, col_f2 = st.columns(2)
      frecuencias = [4, 6, 8, 12, 24]
      frecuencia = col_f1.selectbox(
          "Frecuencia (Horas)",
          frecuencias,
          index=(
              frecuencias.index(val_frec) if val_frec in frecuencias else 2
          ),
      )
      dias = col_f2.number_input(
          "Días de Tratamiento (Hasta 999 días):",
          min_value=1,
          max_value=999,
          value=val_dias,
      )

      col_h1, col_h2 = st.columns(2)
      fecha_inicio = col_h1.date_input("Fecha Inicio", value=val_ini.date())
      hora_inicio = col_h2.time_input("Hora Inicio", value=val_ini.time())
      dt_inicio = datetime.combine(fecha_inicio, hora_inicio)

      # ----------------- PROTOCOLO BOMBA ELASTOMÉRICA (BE) -----------------
      unidosis_puente = 0
      bombas_calc = 0
      dosis_final_calc = 0

      if "BE" in t_via:
        st.markdown(
            "#### 🧪 Protocolo de Bomba Elastomérica (BE) · Cambio cada 24"
            " horas"
        )
        if dt_inicio.hour >= 23 or dt_inicio.hour < 6:
          st.error(
              "🚨 **REGLA T3:** Después de las 23:00 NO se instala BE. Iniciar"
              " con **Unidosis IV de puente** y colocar BE en T1 o T2."
          )

        col_be1, col_be2 = st.columns(2)
        val_puente_prev = (
            safe_int(tto_edit_data[22], 0) if tto_edit_data else 0
        )
        unidosis_puente = col_be1.number_input(
            "Unidosis IV ya administradas antes de montar la BE:",
            min_value=0,
            max_value=50,
            value=val_puente_prev,
        )

        total_unidosis_clinicas = int((dias * 24) / frecuencia)
        bombas_calc, dosis_final_calc, d_x_b = calcular_esquema_be_exacto(
            total_unidosis_clinicas, unidosis_puente, frecuencia
        )

        col_be2.markdown(
            f"""
            <div style='background-color: #D1E7DD; color: #0F5132; padding: 10px; border-radius: 8px; font-size: 0.85rem;'>
                <strong>📦 Conciliación Matemática BE:</strong><br>
                • Total orden médica: <strong>{total_unidosis_clinicas} Unidosis</strong> equivalentes.<br>
                • Unidosis IV administradas: <strong>{unidosis_puente} dosis</strong>.<br>
                • Requiere: <strong>{bombas_calc} Bombas Elastoméricas</strong> ({d_x_b} dosis c/u).<br>
                • Dosis final para completar: <strong>{dosis_final_calc} Unidosis IV</strong>.
            </div>
            """,
            unsafe_allow_html=True,
        )

      # ----------------- ASIGNACIÓN Y AJUSTES DE DOSIS -----------------
      col_aj1, col_aj2, col_aj3 = st.columns([1, 1, 2])
      tipo_unidad_ajuste = col_aj1.selectbox(
          "Unidad de Ajuste:",
          ["Unidosis", "BE Completa"],
          index=0 if "BE" not in t_via else 1,
      )
      val_ajuste_prev = safe_int(tto_edit_data[16], 0) if tto_edit_data else 0
      cant_ajuste = col_aj2.number_input(
          "Cantidad (+ Sumar / - Restar):",
          min_value=-99,
          max_value=99,
          value=val_ajuste_prev,
      )

      # Conversión si ajusta en BE completas
      dosis_equiv_ajuste = cant_ajuste
      if tipo_unidad_ajuste == "BE Completa":
        dosis_por_be = 4 if frecuencia == 6 else 3
        dosis_equiv_ajuste = cant_ajuste * dosis_por_be

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

      # Cálculo de horarios y turnos
      if t_via == "NPT":
        dt_desconexion = dt_inicio + timedelta(hours=npt_horas_infusion)
        npt_hora_desconexion_str = dt_desconexion.strftime("%H:%M")
        h_con = dt_inicio.hour
        h_des = dt_desconexion.hour

        # Lógica NPT: Conectar en su hora y desconectar en la calculada
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
                if 14 <= h_des < 22
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
        dosis_base_fijas = dias
        dosis_totales = dias
        nuevo_mapa_admin = {}
        visita_educacion = ""
      else:
        horas_ciclo, formato_horas = obtener_horas_ciclo(dt_inicio, frecuencia)
        st.markdown("#### 🕒 Asignación de Responsable por Dosis del Día")
        cols_horarios = st.columns(len(formato_horas))
        nuevo_mapa_admin = {}

        mapa_guardado = {}
        if tto_edit_data and tto_edit_data[13]:
          try:
            mapa_guardado = json.loads(tto_edit_data[13])
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
        visita_educacion = ""

      usuario_responsable = st.text_input(
          "✍️ Profesional responsable que guarda/modifica (Obligatorio):"
      )

      # Balance visual en vivo
      st.markdown("#### ⚖️ Conciliación Final y Horarios Calculados")
      mb1, mb2, mb3, mb4 = st.columns(4)
      mb1.metric("Dosis Base Prescritas", f"{dosis_base_fijas} Dosis")
      mb2.metric(
          "Ajuste Dosis",
          f"{cant_ajuste} ({tipo_unidad_ajuste})" if cant_ajuste != 0 else "0",
      )
      mb3.metric("Total Dosis Finales", f"{dosis_totales} Dosis")
      mb4.metric("Fecha Fin Exacta", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))

      c_p1, c_p2, c_p3 = st.columns(3)
      c_p1.metric("Turno T1 (06H-14H)", t1_calc)
      c_p2.metric("Turno T2 (14H-22H)", t2_calc)
      c_p3.metric("Turno T3 (22H-06H)", t3_calc)

      btn_label = (
          "💾 Actualizar Cambios del Tratamiento"
          if tto_edit_data
          else "💾 Guardar Formulación de Tratamiento"
      )
      if st.button(btn_label):
        if not usuario_responsable.strip():
          st.error("❌ Digite su nombre como profesional responsable.")
        elif not t_nombre.strip():
          st.error("❌ Ingrese el medicamento o terapia.")
        else:
          fecha_reg_str = datetime.now().isoformat()
          nov_txt = (
              f"[{'ACTUALIZACIÓN' if tto_edit_data else 'INICIO'} {t_via}]:"
              f" {t_nombre} {t_dosis} por {usuario_responsable.strip()}"
          )

          if t_via == "NPT":
            nov_txt += (
                f" // NPT Infusión {npt_horas_infusion}h (Desconectar a las"
                f" {npt_hora_desconexion_str})"
            )
          elif "BE" in t_via:
            nov_txt += (
                f" // Esquema BE: {unidosis_puente} IV puente + {bombas_calc}"
                f" BE + {dosis_final_calc} IV cierre"
            )

          if tto_edit_data:
            # Actualizar tratamiento existente
            run_query(
                """
                            UPDATE tratamientos SET 
                                tratamiento=?, dosis=?, via_administracion=?, tipo_acceso=?, frecuencia_horas=?, 
                                dias=?, fecha_inicio=?, fecha_fin=?, t1=?, t2=?, t3=?, distribucion_admin=?, 
                                dosis_base_fijas=?, dosis_perdidas=?, tipo_unidad_ajuste=?, motivo_ajuste_dosis=?, 
                                npt_dias_semana=?, npt_horas_infusion=?, npt_hora_desconexion=?, be_unidosis_puente=?, 
                                be_total_bombas=?, be_dosis_final_cierre=?, novedades=novedades || ' // ' || ?, 
                                usuario_modificacion=?, alerta_revisada='NO'
                            WHERE id=?
                        """,
                (
                    t_nombre,
                    t_dosis,
                    t_via,
                    t_acceso,
                    frecuencia,
                    dias,
                    dt_inicio.isoformat(),
                    fecha_fin_calculada.isoformat(),
                    t1_calc,
                    t2_calc,
                    t3_calc,
                    json.dumps(nuevo_mapa_admin),
                    dosis_base_fijas,
                    dosis_equiv_ajuste,
                    tipo_unidad_ajuste,
                    motivo_ajuste,
                    npt_dias_semana_json,
                    npt_horas_infusion,
                    npt_hora_desconexion_str,
                    unidosis_puente,
                    bombas_calc,
                    dosis_final_calc,
                    nov_txt,
                    usuario_responsable.strip(),
                    tto_edit_data[0],
                ),
                fetch=False,
            )
            st.session_state.id_tto_editando = None
            st.success("✅ Tratamiento actualizado con éxito.")
          else:
            # Crear nuevo tratamiento
            run_query(
                """
                            INSERT INTO tratamientos (
                                documento, tratamiento, dosis, via_administracion, tipo_acceso, frecuencia_horas, 
                                dias, fecha_inicio, fecha_fin, t1, t2, t3, distribucion_admin, visita_educacion, 
                                dosis_base_fijas, dosis_perdidas, tipo_unidad_ajuste, motivo_ajuste_dosis, 
                                npt_dias_semana, npt_horas_infusion, npt_hora_desconexion, be_unidosis_puente, 
                                be_total_bombas, be_dosis_final_cierre, novedades, usuario_modificacion, 
                                estado, descanalizado, alerta_revisada, fecha_registro
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', 'NO', 'NO', ?)
                        """,
                (
                    doc_busqueda,
                    t_nombre,
                    t_dosis,
                    t_via,
                    t_acceso,
                    frecuencia,
                    dias,
                    dt_inicio.isoformat(),
                    fecha_fin_calculada.isoformat(),
                    t1_calc,
                    t2_calc,
                    t3_calc,
                    json.dumps(nuevo_mapa_admin),
                    visita_educacion,
                    dosis_base_fijas,
                    dosis_equiv_ajuste,
                    tipo_unidad_ajuste,
                    motivo_ajuste,
                    npt_dias_semana_json,
                    npt_horas_infusion,
                    npt_hora_desconexion_str,
                    unidosis_puente,
                    bombas_calc,
                    dosis_final_calc,
                    nov_txt,
                    usuario_responsable.strip(),
                    fecha_reg_str,
                ),
                fetch=False,
            )
            st.success("✅ Tratamiento formulado con éxito.")

          st.rerun()

# ----------------- SECCIÓN 2: PESTAÑA NOVEDADES Y AJUSTES DE RUTA -----------------
elif menu == "📢 Novedades y Ajustes de Ruta":
  st.subheader("📢 Gestión de Novedades Clínicas y Modificación de Horarios")

  st.markdown("### 📋 Pacientes con Novedades Activas / Pendientes por Gestionar")
  query_novs_generales = """
        SELECT 
            rn.documento AS 'Cédula',
            p.nombre AS 'Nombre y Apellido',
            p.piso AS 'Piso',
            p.zona AS 'Zona',
            p.plan AS 'Plan',
            rn.tipo_novedad AS 'Tipo Novedad',
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

    st.info(f"📢 Pendientes: `{len(df_novs_filtradas)}` pacientes")
    st.dataframe(df_novs_filtradas, use_container_width=True)

    opciones_pacientes_nov = [
        f"{row['Cédula']} - {row['Nombre y Apellido']} ({row['Zona']})"
        for _, row in df_novs_filtradas.iterrows()
    ]
    paciente_escogido = st.selectbox(
        "👉 Seleccionar paciente para resolver novedad:",
        ["-- Seleccionar paciente --"] + opciones_pacientes_nov,
    )
    if paciente_escogido != "-- Seleccionar paciente --":
      doc_preseleccionado = paciente_escogido.split(" - ")[0].strip()
  else:
    st.success("🎉 No hay novedades pendientes activas.")

  st.divider()
  doc_nov = st.text_input(
      "Digite la Cédula del Paciente:",
      value=doc_preseleccionado,
      key="doc_nov_input",
  ).strip()

  if doc_nov:
    p_data = run_query(
        "SELECT documento, nombre, plan, programa, zona, piso FROM pacientes"
        " WHERE documento = ?",
        (doc_nov,),
    )
    if p_data:
      p_doc, p_nom, p_plan, p_prog, p_zona, p_piso = p_data[0]
      t_data = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC LIMIT 1",
          (doc_nov,),
      )

      st.markdown(
          f"#### 👤 Paciente: {p_nom} ({p_doc}) · Zona `{p_zona}` ({p_piso})"
      )

      tab_gestionar, tab_nuevo = st.tabs(
          ["🛎️ Novedades Pendientes y Gestión", "➕ Registrar Nueva Novedad"]
      )

      with tab_gestionar:
        novs_pend = run_query(
            "SELECT id, tipo_novedad, fecha_aplicacion, hora_aplicacion,"
            " detalle, responsable FROM registro_novedades WHERE documento = ?"
            " AND estado_novedad = 'Pendiente'",
            (doc_nov,),
        )
        if novs_pend:
          for n in novs_pend:
            st.warning(
                f"📌 **{n[1]}** ({n[2]} {n[3]}): {n[4]} (Registró: {n[5]})"
            )

          prof_ges = st.text_input(
              "Profesional que atiende / gestiona (Firma):", key="prof_ges_res"
          )
          acc_ges = st.text_area(
              "Acción clínica realizada:",
              value="Novedad atendida y coordinada.",
          )

          if st.button("✅ Confirmar Novedad Resuelta y Retirar Alerta"):
            if not prof_ges.strip():
              st.error("Ingrese su nombre como responsable.")
            else:
              fec_res = datetime.now().strftime("%d/%m/%Y %I:%M %p")
              run_query(
                  """UPDATE registro_novedades SET estado_novedad = 'Gestionada', 
                             fecha_gestion = ?, responsable_gestion = ? WHERE documento = ? AND estado_novedad = 'Pendiente'""",
                  (fec_res, prof_ges.strip(), doc_nov),
                  fetch=False,
              )
              if t_data:
                run_query(
                    """UPDATE tratamientos SET alerta_revisada = 'SI', 
                               novedades = novedades || ' // [GESTIONADA ' || ? || ' por ' || ? || ']' WHERE id = ?""",
                    (fec_res, prof_ges.strip(), t_data[0][0]),
                    fetch=False,
                )
              st.success("✅ Novedad resuelta y alerta retirada.")
              st.rerun()
        else:
          st.success("No hay novedades pendientes para este paciente.")

      with tab_nuevo:
        tipo_nov = st.selectbox(
            "Tipo de Novedad:",
            [
                "Cambio de Horario / Reacomodo de Visita",
                "Novedad de Salida / Descanalización",
                "Modificación por Cambio de Orden Médica",
                "Paciente Ausente / No responde",
                "Familiar suspende dosis",
                "Otra Novedad",
            ],
        )
        fec_n = st.date_input("Fecha Novedad:", value=datetime.now().date())
        hor_n = st.time_input("Hora Novedad:", value=datetime.now().time())
        det_n = st.text_area("Detalle:")
        resp_n = st.text_input(
            "Profesional Responsable (Firma):", key="prof_crea_nov"
        )

        if st.button("💾 Guardar Novedad"):
          if not resp_n.strip():
            st.error("Ingrese su nombre.")
          else:
            run_query(
                """INSERT INTO registro_novedades (documento, tipo_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                           VALUES (?, ?, ?, ?, ?, ?, 'Pendiente', ?)""",
                (
                    doc_nov,
                    tipo_nov,
                    fec_n.strftime("%d/%m/%Y"),
                    hor_n.strftime("%I:%M %p"),
                    det_n,
                    resp_n.strip(),
                    datetime.now().isoformat(),
                ),
                fetch=False,
            )
            if t_data:
              run_query(
                  "UPDATE tratamientos SET alerta_revisada = 'NO', novedades ="
                  " novedades || ' // 📢 NOVEDAD PENDIENTE: ' || ? WHERE id ="
                  " ?",
                  (det_n, t_data[0][0]),
                  fetch=False,
              )
            st.success("✅ Novedad registrada con éxito.")
            st.rerun()

# ----------------- SECCIÓN 3: CENSO Y PLANILLA EN VIVO -----------------
elif menu == "📊 Censo y Planilla en Vivo":
  col_vista1, col_vista2 = st.columns([3, 1])
  filtro_estado = col_vista2.selectbox(
      "Filtrar por Estado:", ["Activos", "Completados / Descanalizados", "Todos"]
  )

  condicion_estado = "WHERE t.estado = 'Activo'"
  if filtro_estado == "Completados / Descanalizados":
    condicion_estado = "WHERE t.estado = 'Completado'"
  elif filtro_estado == "Todos":
    condicion_estado = "WHERE t.estado IN ('Activo', 'Completado')"

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
            COALESCE(t.tratamiento, 'Sin Tratamiento') AS 'Tratamiento',
            COALESCE(t.dosis, '-') AS 'Dosis',
            COALESCE(t.via_administracion, '-') AS 'Vía',
            COALESCE(t.tipo_acceso, 'Periférico') AS 'Tipo Acceso',
            COALESCE(t.frecuencia_horas, 0) AS 'Frec (h)',
            COALESCE(t.dias, 0) AS 'Días',
            COALESCE(t.fecha_inicio, '-') AS 'Fecha Inicio',
            COALESCE(t.fecha_fin, '-') AS 'Fecha Fin',
            COALESCE(t.t1, '-') AS 'T1',
            COALESCE(t.t2, '-') AS 'T2',
            COALESCE(t.t3, '-') AS 'T3',
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

    df_filtrado.reset_index(drop=True, inplace=True)
    df_filtrado.index = df_filtrado.index + 1
    df_filtrado.index.name = "N°"

    # Conteo
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
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#721C24;'>{total_picc}</div><div"
        " class='metric-lbl'>🩸 CON CATÉTER PICC</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#664D03;'>{total_npt}</div><div"
        " class='metric-lbl'>🥣 CON NPT</div></div>",
        unsafe_allow_html=True,
    )
    m4.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#0F5132;'>{total_be}</div><div"
        " class='metric-lbl'>💉 INFUSIÓN BE / ANALGESIA</div></div>",
        unsafe_allow_html=True,
    )
    m5.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#4A148C;'>{total_aislados}</div><div"
        " class='metric-lbl'>☣️ AISLADOS</div></div>",
        unsafe_allow_html=True,
    )
    m6.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#E65100;'>{total_novedades}</div><div"
        " class='metric-lbl'>📢 NOVEDADES ACTIVAS</div></div>",
        unsafe_allow_html=True,
    )
    m7.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#B71C1C;'>{total_descanalizar}</div><div"
        " class='metric-lbl'>⚠️ PENDIENTE RETIRO</div></div>",
        unsafe_allow_html=True,
    )

    # Convenciones Compactas en Colores Pastel
    st.markdown("##### 🎨 Convenciones de Color en la Planilla")
    c_c1, c_c2, c_c3, c_c4 = st.columns(4)
    with c_c1:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_PICC}; color:"
          " #721c24;'>🩸 Catéter PICC</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_NPT}; color:"
          " #664d03;'>🥣 Nutrición Parenteral (NPT)</div>",
          unsafe_allow_html=True,
      )
    with c_c2:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_BE}; color:"
          " #0f5132;'>🟩 Bomba Elastomérica (BE)</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_ANALGESIA};"
          " color: #4a148c;'>🟪 BE Analgesia / Dolor</div>",
          unsafe_allow_html=True,
      )
    with c_c3:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_TB}; color:"
          " #842029;'>🟧 Tuberculosis (TB)</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_PUSH}; color:"
          " #055160;'>🟦 Vía PUSH</div>",
          unsafe_allow_html=True,
      )
    with c_c4:
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_AISLA};"
          " color: #4a148c;'>🟣 Paciente Aislado</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<div class='conv-box' style='background-color: {COLOR_FALTA1};"
          " color: #b71c1c;'>⚠️ Falta 1 Dosis / Salida</div>",
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.dataframe(
        df_filtrado.style.apply(estilo_filas, axis=1), use_container_width=True
    )

    # Descarga de Excel CON COLORES APLICADOS
    excel_coloreado = exportar_excel_con_colores(df_filtrado)

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
                    p.piso AS 'Piso', p.zona AS 'Zona', p.aislamiento AS 'Aislamiento', 
                    t.tratamiento AS 'Medicamento', t.dosis AS 'Dosis', t.via_administracion AS 'Vía', 
                    t.tipo_acceso AS 'Tipo Acceso', t.frecuencia_horas AS 'Frec (h)', t.dias AS 'Días', 
                    t.fecha_inicio AS 'Fecha Inicio', t.fecha_fin AS 'Fecha Fin', 
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
                    p.piso AS 'Piso', p.zona AS 'Zona', p.estado_paciente AS 'Estado Paciente',
                    t.tratamiento AS 'Medicamento', t.dosis AS 'Dosis', t.via_administracion AS 'Vía', 
                    t.tipo_acceso AS 'Tipo Acceso', t.dias AS 'Días', t.fecha_inicio AS 'Fecha Inicio', 
                    t.fecha_fin AS 'Fecha Fin', t.t1 AS 'T1', t.t2 AS 'T2', t.t3 AS 'T3', 
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
