import io
import json
import os
import sqlite3
from datetime import datetime, timedelta
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
    
    /* Convenciones de color en grid compacto */
    .conv-box {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-bottom: 4px;
        width: 100%;
        text-align: center;
    }
    
    /* Pie de página autoría fijo */
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
            motivo_ajuste_dosis TEXT DEFAULT '',
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

  c.execute("PRAGMA table_info(tratamientos)")
  cols_t = [col[1] for col in c.fetchall()]
  if "dosis_base_fijas" not in cols_t:
    c.execute(
        "ALTER TABLE tratamientos ADD COLUMN dosis_base_fijas INTEGER DEFAULT 0"
    )
  if "motivo_ajuste_dosis" not in cols_t:
    c.execute(
        "ALTER TABLE tratamientos ADD COLUMN motivo_ajuste_dosis TEXT DEFAULT"
        " ''"
    )

  c.execute("PRAGMA table_info(registro_novedades)")
  cols_nov = [col[1] for col in c.fetchall()]
  if "estado_novedad" not in cols_nov:
    c.execute(
        "ALTER TABLE registro_novedades ADD COLUMN estado_novedad TEXT DEFAULT"
        " 'Pendiente'"
    )
  if "fecha_gestion" not in cols_nov:
    c.execute("ALTER TABLE registro_novedades ADD COLUMN fecha_gestion TEXT")
  if "responsable_gestion" not in cols_nov:
    c.execute(
        "ALTER TABLE registro_novedades ADD COLUMN responsable_gestion TEXT"
    )

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


# ----------------- CÁLCULOS CLÍNICOS -----------------
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
        "background-color: #ffe0b2; color: #e65100; font-weight: bold;"
    ] * len(row)

  if "FALTA 1 DOSIS" in alerta:
    return [
        "background-color: #ffcc80; color: #b71c1c; font-weight: bold;"
    ] * len(row)

  if acceso == "PICC":
    return [
        "background-color: #ffd6d6; color: #721c24; font-weight: bold;"
    ] * len(row)

  if "NPT" in via:
    return [
        "background-color: #fff8e1; color: #5d4037; font-weight: bold;"
    ] * len(row)
  elif "ANALGESIA" in via:
    return [
        "background-color: #ede7f6; color: #4a148c; font-weight: bold;"
    ] * len(row)
  elif "BE" in via:
    return [
        "background-color: #e8f5e9; color: #1b5e20; font-weight: bold;"
    ] * len(row)
  elif "TUBERCULOSIS" in via or "TB" in via:
    return [
        "background-color: #ffe0b2; color: #bf360c; font-weight: bold;"
    ] * len(row)
  elif via == "PUSH":
    return [
        "background-color: #e0f7fa; color: #006064; font-weight: bold;"
    ] * len(row)

  if aislamiento == "SI":
    return [
        "background-color: #e1bee7; color: #4a148c; font-weight: bold;"
    ] * len(row)

  if "PRÓXIMO A FIN" in alerta:
    return [
        "background-color: #fff3cd; color: #856404; font-weight: bold;"
    ] * len(row)

  return [""] * len(row)


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

  # TARJETA DE AUTORÍA POSICIONADA AL PIE
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

# ----------------- SECCIÓN 1: GESTIÓN DE PACIENTES -----------------
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
        "👤 Información Demográfica del Paciente",
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
        p_user = (
            paciente_data[0][10]
            if len(paciente_data[0]) > 10 and paciente_data[0][10]
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
                           municipio=?, barrio=?, aislamiento=?, tipo_aislamiento=?, usuario_registro=? WHERE documento=?""",
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
            st.success("Datos actualizados correctamente.")
          else:
            st.error("Debes ingresar tu nombre como responsable.")
      else:
        st.info("Paciente nuevo. Ingrese los datos de admisión:")
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

        if st.button("Ingresar Nuevo Paciente"):
          if nombre and usuario_crea:
            run_query(
                """INSERT INTO pacientes (documento, nombre, plan, programa, piso, zona, municipio, barrio, aislamiento, tipo_aislamiento, usuario_registro) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
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
            st.success("Paciente admitido con éxito.")
            st.rerun()
          else:
            st.error("Diligencie el nombre del paciente y su nombre.")

    if paciente_data:
      st.divider()

      tratamiento_activo = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC LIMIT 1",
          (doc_busqueda,),
      )

      st.subheader("💊 Formulación / Configuración Base de Tratamiento")

      vias = [
          "IV",
          "SC",
          "IM",
          "BE",
          "NPT",
          "BE analgesia",
          "Tuberculosis",
          "Materna",
          "PUSH",
      ]
      accesos = ["Periférico", "PICC", "SC", "Ninguno"]

      via_val = (
          tratamiento_activo[0][4]
          if tratamiento_activo and tratamiento_activo[0][4] in vias
          else "IV"
      )
      acceso_val = (
          tratamiento_activo[0][5]
          if tratamiento_activo and tratamiento_activo[0][5] in accesos
          else "Periférico"
      )

      mapa_guardado = {}
      if tratamiento_activo and len(tratamiento_activo[0]) > 13:
        try:
          mapa_guardado = json.loads(tratamiento_activo[0][13])
        except Exception:
          mapa_guardado = {}

      col_t1, col_t2, col_t3, col_t4 = st.columns([2, 1, 1, 1])
      t_nombre = col_t1.text_input(
          "Medicamento / Tratamiento",
          value=tratamiento_activo[0][2] if tratamiento_activo else "",
      )
      t_dosis = col_t2.text_input(
          "Dosis", value=tratamiento_activo[0][3] if tratamiento_activo else ""
      )
      t_via = col_t3.selectbox("Vía", vias, index=vias.index(via_val))
      t_acceso = col_t4.selectbox(
          "Acceso", accesos, index=accesos.index(acceso_val)
      )

      col_f1, col_f2 = st.columns(2)
      frecuencias = [4, 6, 8, 12, 24]
      frec_val = safe_int(
          tratamiento_activo[0][6] if tratamiento_activo else 8, 8
      )
      frecuencia = col_f1.selectbox(
          "Frecuencia (Horas)",
          frecuencias,
          index=frecuencias.index(frec_val) if frec_val in frecuencias else 2,
      )

      dias_val = safe_int(
          tratamiento_activo[0][7] if tratamiento_activo else 5, 5
      )
      dias = col_f2.number_input(
          "Días de Tratamiento (Hasta 999 días):",
          min_value=1,
          max_value=999,
          value=dias_val,
      )

      col_h1, col_h2 = st.columns(2)
      default_inicio = (
          datetime.fromisoformat(tratamiento_activo[0][8])
          if tratamiento_activo
          else datetime.now()
      )
      fecha_inicio = col_h1.date_input(
          "Fecha Inicio", value=default_inicio.date()
      )
      hora_inicio = col_h2.time_input(
          "Hora Inicio", value=default_inicio.time()
      )

      dt_inicio = datetime.combine(fecha_inicio, hora_inicio)
      horas_ciclo, formato_horas = obtener_horas_ciclo(dt_inicio, frecuencia)

      st.markdown("#### 🕒 Asignación de Responsable por Dosis del Día")
      cols_horarios = st.columns(len(formato_horas))
      nuevo_mapa_admin = {}
      hay_cuidador = False

      for idx, h_str in enumerate(formato_horas):
        with cols_horarios[idx]:
          default_admin_idx = 0
          if mapa_guardado.get(h_str) == "Cuidador":
            default_admin_idx = 1

          resp_sel = st.selectbox(
              f"Dosis {h_str}:",
              ["SENC", "Cuidador"],
              index=default_admin_idx,
              key=f"admin_{h_str}",
          )
          nuevo_mapa_admin[h_str] = resp_sel
          if resp_sel == "Cuidador":
            hay_cuidador = True

      visita_educacion = ""
      if hay_cuidador:
        st.info("👨‍⚕️ **Dosis con Cuidador:** Programar visita de educación.")
        col_ed1, col_ed2 = st.columns(2)
        fecha_ed = col_ed1.date_input(
            "Fecha Visita Educación:", value=datetime.now().date()
        )
        hora_ed = col_ed2.time_input(
            "Hora Visita Educación:", value=datetime.now().time()
        )
        visita_educacion = f"{fecha_ed.strftime('%d/%m/%Y')} {hora_ed.strftime('%I:%M %p')}"

      col_rep1, col_rep2 = st.columns([1, 2])
      ajuste_val = safe_int(
          tratamiento_activo[0][16]
          if tratamiento_activo and len(tratamiento_activo[0]) > 16
          else 0,
          0,
      )
      ajuste_dosis = col_rep1.number_input(
          "Ajuste Dosis (+ Sumar / - Restar)",
          min_value=-99,
          max_value=99,
          value=ajuste_val,
          help=(
              "Permite agregar o restar dosis adicionales al ciclo prescrito."
          ),
      )

      motivos_ajuste_lista = [
          "Ninguno / Dosis exactas de orden médica",
          "Paciente no se encontraba en domicilio (Ausente)",
          "Reprogramación de dosis",
          "Dosis no administrada por enfermería",
          "Suspensión anticipada de dosis médica",
          "Reposición autorizada",
      ]
      motivo_guardado = (
          tratamiento_activo[0][17]
          if tratamiento_activo and len(tratamiento_activo[0]) > 17
          else ""
      )
      idx_m_ajuste = (
          motivos_ajuste_lista.index(motivo_guardado)
          if motivo_guardado in motivos_ajuste_lista
          else 0
      )

      if ajuste_dosis != 0:
        motivo_ajuste = col_rep2.selectbox(
            "Motivo de Dosis Adicional / Restada (Obligatorio):",
            motivos_ajuste_lista[1:],
            index=max(0, idx_m_ajuste - 1),
        )
      else:
        motivo_ajuste = col_rep2.selectbox(
            "Estado de Ajuste de Dosis:",
            ["Ninguno / Dosis exactas de orden médica"],
            index=0,
            disabled=True,
        )

      col_mod1, col_mod2 = st.columns(2)
      motivo_modificacion = col_mod1.selectbox(
          "Motivo de formulación / cambio (Obligatorio):",
          [
              "Inicio de Esquema / Admisión",
              "Cambio de Orden Médica",
              "Prórroga de Días",
              "Ajuste de Dosis",
              "Reanudación de Tratamiento",
          ],
      )

      usuario_cambio = col_mod2.text_input(
          "✍️ Nombre del Profesional Responsable (Obligatorio):", value=""
      )

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
          dt_inicio, dias, frecuencia, nuevo_mapa_admin, ajuste_dosis
      )

      st.markdown("#### ⚖️ Conciliación y Balance de Dosis Totales")
      mb1, mb2, mb3, mb4 = st.columns(4)
      mb1.metric("1. Dosis Base Prescritas (Fijas)", f"{dosis_base_fijas} Dosis")

      txt_ajuste = "0 (Sin cambios)"
      if ajuste_dosis > 0:
        txt_ajuste = f"+{ajuste_dosis} Dosis de MÁS ({motivo_ajuste})"
      elif ajuste_dosis < 0:
        txt_ajuste = f"{ajuste_dosis} Dosis de MENOS ({motivo_ajuste})"

      mb2.metric("2. Dosis de Más / Menos", txt_ajuste)
      mb3.metric("3. Total Dosis Finales a Cumplir", f"{dosis_totales} Dosis")
      mb4.metric("Fecha Fin Exacta", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))

      st.markdown("#### 📋 Horarios de Visitas Calculados")
      c_p1, c_p2, c_p3 = st.columns(3)
      c_p1.metric("Turno T1 (06H-14H)", t1_calc)
      c_p2.metric("Turno T2 (14H-22H)", t2_calc)
      c_p3.metric("Turno T3 (22H-06H)", t3_calc)

      if st.button("💾 Guardar Formulación con Firma"):
        if not usuario_cambio.strip():
          st.error("❌ Ingrese su nombre como profesional responsable.")
        elif not t_nombre.strip():
          st.error("❌ Ingrese el medicamento / tratamiento.")
        else:
          fecha_reg_str = datetime.now().isoformat()
          fecha_ini_str = dt_inicio.isoformat()
          fecha_fin_str = fecha_fin_calculada.isoformat()
          json_admin = json.dumps(nuevo_mapa_admin)

          nov_anterior = (
              tratamiento_activo[0][18]
              if tratamiento_activo and len(tratamiento_activo[0]) > 18
              else ""
          )
          nov_actualizada = (
              f"[{motivo_modificacion.upper()}]: Registrado por"
              f" {usuario_cambio.strip()} el"
              f" {datetime.now().strftime('%d/%m/%Y %I:%M %p')}"
          )
          if ajuste_dosis != 0:
            nov_actualizada += (
                f" // [AJUSTE: {txt_ajuste} por {usuario_cambio.strip()}]"
            )
          if nov_anterior:
            nov_actualizada = f"{nov_actualizada} // {nov_anterior}"

          if tratamiento_activo:
            run_query(
                "UPDATE tratamientos SET estado = 'Modificado' WHERE id = ?",
                (tratamiento_activo[0][0],),
                fetch=False,
            )

          run_query(
              """
                        INSERT INTO tratamientos (
                            documento, tratamiento, dosis, via_administracion, tipo_acceso, frecuencia_horas, 
                            dias, fecha_inicio, fecha_fin, t1, t2, t3, distribucion_admin, visita_educacion, 
                            dosis_base_fijas, dosis_perdidas, motivo_ajuste_dosis, novedades, usuario_modificacion, 
                            estado, descanalizado, alerta_revisada, fecha_registro
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', 'NO', 'NO', ?)
                    """,
              (
                  doc_busqueda,
                  t_nombre,
                  t_dosis,
                  t_via,
                  t_acceso,
                  frecuencia,
                  dias,
                  fecha_ini_str,
                  fecha_fin_str,
                  t1_calc,
                  t2_calc,
                  t3_calc,
                  json_admin,
                  visita_educacion,
                  dosis_base_fijas,
                  ajuste_dosis,
                  motivo_ajuste if ajuste_dosis != 0 else "",
                  nov_actualizada,
                  usuario_cambio.strip(),
                  fecha_reg_str,
              ),
              fetch=False,
          )

          st.success(
              f"Formulación guardada y visible en tiempo real por:"
              f" {usuario_cambio}."
          )
          st.rerun()

      if tratamiento_activo:
        st.write("### 📌 Tratamiento Activo Actual")
        frec_act = safe_int(tratamiento_activo[0][6], 8)
        estado_act = (
            tratamiento_activo[0][20]
            if len(tratamiento_activo[0]) > 20
            else "Activo"
        )
        desc_act = (
            tratamiento_activo[0][21] if len(tratamiento_activo[0]) > 21 else "NO"
        )
        nov_act = (
            tratamiento_activo[0][18] if len(tratamiento_activo[0]) > 18 else ""
        )
        rev_act = (
            tratamiento_activo[0][22] if len(tratamiento_activo[0]) > 22 else "NO"
        )

        estado_alerta_act = evaluar_alerta_fin(
            tratamiento_activo[0][9],
            frec_act,
            estado_act,
            desc_act,
            nov_act,
            rev_act,
        )

        d_base = (
            tratamiento_activo[0][15]
            if len(tratamiento_activo[0]) > 15
            else (dias_val * 24 // frec_act)
        )
        d_ajust = (
            tratamiento_activo[0][16] if len(tratamiento_activo[0]) > 16 else 0
        )
        mot_ajust = (
            tratamiento_activo[0][17] if len(tratamiento_activo[0]) > 17 else ""
        )

        df_activo = pd.DataFrame([{
            "Tratamiento": tratamiento_activo[0][2],
            "Dosis": tratamiento_activo[0][3],
            "Vía": tratamiento_activo[0][4],
            "Tipo Acceso": tratamiento_activo[0][5],
            "Frecuencia": f"C/{frec_act}h",
            "Días": tratamiento_activo[0][7],
            "Dosis Base Fijas": d_base,
            "Ajuste Dosis (+/-)": (
                f"{d_ajust} ({mot_ajust})" if d_ajust != 0 else "0"
            ),
            "Fecha Fin": datetime.fromisoformat(
                tratamiento_activo[0][9]
            ).strftime("%d/%m/%Y %I:%M %p"),
            "T1": tratamiento_activo[0][10],
            "T2": tratamiento_activo[0][11],
            "T3": tratamiento_activo[0][12],
            "Estado Tratamiento": estado_alerta_act,
            "Registrado Por": (
                tratamiento_activo[0][19]
                if len(tratamiento_activo[0]) > 19
                else ""
            ),
            "NOVEDAD": nov_act,
        }])
        st.dataframe(
            df_activo.style.apply(estilo_filas, axis=1), use_container_width=True
        )

      st.divider()
      st.subheader("📜 Movimientos y Trazabilidad en Tiempo Real")
      historial = run_query(
          """SELECT tratamiento, dosis, via_administracion, tipo_acceso, frecuencia_horas, 
                           dias, fecha_inicio, fecha_fin, t1, t2, t3, 
                           usuario_modificacion, novedades, estado, fecha_registro 
                    FROM tratamientos WHERE documento = ? ORDER BY id DESC""",
          (doc_busqueda,),
      )
      if historial:
        df_hist = pd.DataFrame(
            historial,
            columns=[
                "Tratamiento",
                "Dosis",
                "Vía",
                "Tipo Acceso",
                "Frec (h)",
                "Días",
                "Inicio",
                "Fin",
                "T1",
                "T2",
                "T3",
                "Modificado Por",
                "NOVEDAD",
                "Estado",
                "Fecha Registro",
            ],
        )
        st.dataframe(df_hist, use_container_width=True)

# ----------------- SECCIÓN 2: NOVEDADES Y AJUSTES DE RUTA -----------------
elif menu == "📢 Novedades y Ajustes de Ruta":
  st.subheader("📢 Gestión de Novedades Clínicas y Modificación de Horarios")

  st.markdown("### 📋 Pacientes con Novedades Activas / Pendientes por Gestionar")
  st.caption(
      "Filtra por Zona, Piso o Tipo de Novedad para revisar los pacientes de tu"
      " sector y selecciona o digita la cédula para atenderla."
  )

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

    conteo_filtrado = len(df_novs_filtradas)
    zona_subtxt = f" ({', '.join(f_zonas_nov)})" if f_zonas_nov else ""
    st.info(
        f"📢 **Novedades Pendientes Encontradas:** `{conteo_filtrado}` pacientes"
        f"{zona_subtxt}"
    )

    st.dataframe(df_novs_filtradas, use_container_width=True)

    opciones_pacientes_nov = [
        f"{row['Cédula']} - {row['Nombre y Apellido']} ({row['Zona']})"
        for _, row in df_novs_filtradas.iterrows()
    ]
    paciente_escogido = st.selectbox(
        "👉 O selecciona directamente un paciente del listado anterior:",
        ["-- Seleccionar paciente para gestionar --"] + opciones_pacientes_nov,
    )
    if paciente_escogido != "-- Seleccionar paciente para gestionar --":
      doc_preseleccionado = paciente_escogido.split(" - ")[0].strip()

  else:
    st.success("🎉 No hay novedades pendientes activas en este momento.")

  st.divider()

  st.markdown("### 🔍 Gestionar Novedad o Registrar Nuevo Reacomodo")
  doc_nov = st.text_input(
      "Digite la Cédula del Paciente a Gestionar:",
      value=doc_preseleccionado,
      key="doc_nov_input",
  ).strip()

  if doc_nov:
    p_data = run_query(
        "SELECT documento, nombre, plan, programa, zona, piso, municipio,"
        " barrio, aislamiento FROM pacientes WHERE documento = ?",
        (doc_nov,),
    )

    if not p_data:
      st.error(
          "❌ Paciente no encontrado. Debe crearlo primero en 'Gestión de"
          " Pacientes'."
      )
    else:
      p_doc, p_nom, p_plan, p_prog, p_zona, p_piso, p_muni, p_barrio, p_ais = (
          p_data[0]
      )

      t_data = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC LIMIT 1",
          (doc_nov,),
      )

      st.markdown("#### 👤 Ficha del Paciente")
      col_res1, col_res2, col_res3, col_res4 = st.columns(4)
      col_res1.markdown(f"**Nombre:** {p_nom}\n\n**Cédula:** {p_doc}")
      col_res2.markdown(
          f"**Zona:** `{p_zona}` ({p_piso})\n\n**Plan / Programa:** {p_plan} -"
          f" {p_prog}"
      )

      if t_data:
        t_id, _, t_med, t_dos, t_via, t_acc, t_frec, t_dias, t_ini, t_fin = (
            t_data[0][0:10]
        )
        t_t1, t_t2, t_t3 = t_data[0][10:13]
        t_dist_json = t_data[0][13] if len(t_data[0]) > 13 else "{}"
        t_nov_actual = (
            t_data[0][18]
            if len(t_data[0]) > 18
            else (t_data[0][16] if len(t_data[0]) > 16 else "")
        )

        col_res3.markdown(
            f"**Medicamento:** {t_med} ({t_dos})\n\n**Vía/Acceso:** {t_via} -"
            f" {t_acc}"
        )
        col_res4.markdown(
            f"**Frecuencia:** Cada {t_frec}h\n\n**Fin TTO:**"
            f" {datetime.fromisoformat(t_fin).strftime('%d/%m %I:%M %p')}"
        )

        st.info(
            f"🕒 **Horarios Actuales de Ruta:** T1: `{t_t1}` | T2: `{t_t2}` |"
            f" T3: `{t_t3}`"
        )
      else:
        st.warning("⚠️ El paciente no tiene un tratamiento activo formulado.")

      st.markdown("---")

      novs_pendientes_db = run_query(
          """SELECT id, tipo_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable 
                    FROM registro_novedades WHERE documento = ? AND estado_novedad = 'Pendiente' ORDER BY id DESC""",
          (doc_nov,),
      )

      tab_gestionar, tab_nuevo = st.tabs(
          ["🛎️ Novedades Pendientes y Gestión", "➕ Registrar Nueva Novedad"]
      )

      with tab_gestionar:
        st.markdown("### 🛎️ Novedades y Alertas Pendientes de este Paciente")

        hay_novedad_activa = bool(novs_pendientes_db) or (
            t_data
            and (
                "NOVEDAD PENDIENTE" in t_nov_actual.upper()
                or "VISITA PUNTUAL" in t_nov_actual.upper()
                or "SALIDA" in t_nov_actual.upper()
            )
        )

        if not hay_novedad_activa:
          st.success(
              "✅ Este paciente no tiene novedades pendientes por gestionar."
          )
        else:
          if novs_pendientes_db:
            for nov_item in novs_pendientes_db:
              n_id, n_tipo, n_fec, n_hor, n_det, n_resp = nov_item
              st.warning(
                  f"📌 **{n_tipo}** (Para: {n_fec} {n_hor})\n\n**Detalle:**"
                  f" {n_det}\n\n*Registrado por:* {n_resp}"
              )
          elif t_data and t_nov_actual:
            st.warning(f"📌 **Novedad en Tratamiento:**\n\n{t_nov_actual}")

          st.markdown("#### ✍️ Formulario para Resolver y Retirar Alerta")
          c_ges1, c_ges2 = st.columns(2)
          prof_gestiona = c_ges1.text_input(
              "Tu Nombre / Profesional que gestiona (Obligatorio):",
              key="prof_res_directo",
          )
          accion_tipo = c_ges2.selectbox(
              "Acción realizada:",
              [
                  "Visita Realizada / Novedad Atendida",
                  "Paciente Coordinado y Notificado",
                  "Paciente Gestionado / Descanalizado",
                  "Cerrar Alerta",
              ],
              key="acc_tipo_sel",
          )

          det_accion = st.text_area(
              "Observación / Detalle de la gestión:",
              value="Novedad atendida satisfactoriamente.",
              key="det_res_directo",
          )

          if st.button(
              "✅ Confirmar Gestión y Retirar Alerta de la Planilla",
              key="btn_resolver_alerta",
          ):
            if not prof_gestiona.strip():
              st.error(
                  "❌ Ingrese su nombre como profesional responsable de la"
                  " gestión."
              )
            else:
              fec_gestion_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
              nov_resuelta_txt = (
                  f"[GESTIONADA {fec_gestion_str} por"
                  f" {prof_gestiona.strip()}]: {accion_tipo} - {det_accion}"
              )

              run_query(
                  """UPDATE registro_novedades SET estado_novedad = 'Gestionada', 
                             fecha_gestion = ?, responsable_gestion = ? WHERE documento = ? AND estado_novedad = 'Pendiente'""",
                  (fec_gestion_str, prof_gestiona.strip(), doc_nov),
                  fetch=False,
              )

              if t_data:
                run_query(
                    """UPDATE tratamientos SET novedades = ?, alerta_revisada = 'SI', 
                               detalle_revision_alerta = ?, responsable_revision_alerta = ? WHERE id = ?""",
                    (
                        nov_resuelta_txt,
                        det_accion,
                        prof_gestiona.strip(),
                        t_data[0][0],
                    ),
                    fetch=False,
                )

              st.success(
                  "✅ Novedad gestionada y confirmada. La alerta fue retirada de"
                  " la planilla en tiempo real."
              )
              st.rerun()

      with tab_nuevo:
        st.markdown("### ➕ Registrar una Nueva Novedad o Reacomodo de Ruta")

        tipo_nov = st.selectbox(
            "Tipo de Novedad a Registrar:",
            [
                "Cambio de Horario / Reacomodo de Visita (Paciente solicita más"
                " tarde o más temprano)",
                "Novedad de Salida / Descanalización",
                "Modificación por Cambio de Orden Médica",
                "Paciente Ausente / No responde llamado",
                "Familiar rechaza visita / Suspende dosis",
                "Otra Novedad de Ruta",
            ],
            key="tipo_nov_sel",
        )

        col_fec1, col_fec2 = st.columns(2)
        fecha_aplica = col_fec1.date_input(
            "Fecha en que aplica la novedad:",
            value=datetime.now().date(),
            key="fec_aplica_input",
        )
        hora_aplica = col_fec2.time_input(
            "Hora en que aplica la novedad:",
            value=datetime.now().time(),
            key="hor_aplica_input",
        )

        if "Cambio de Horario" in tipo_nov and t_data:
          st.markdown("##### ⏱️ Configuración del Reacomodo de Horarios")
          modo_cambio = st.radio(
              "Modalidad de aplicación:",
              [
                  "1. Modificar automáticamente todo el ciclo permanente"
                  " (Corre todas las visitas y recalcula la fecha fin)",
                  "2. Visita puntual de un solo día (Activa alerta en planilla"
                  " sin alterar ciclo permanente)",
              ],
              key="modo_h_sel",
          )

          col_h_ant, col_h_nuev = st.columns(2)
          hora_original = col_h_ant.time_input(
              "Hora original que se visitaba (Ej: 18:00):",
              value=datetime.now().time(),
              key="h_orig_input",
          )
          hora_nueva = col_h_nuev.time_input(
              "Nueva hora solicitada por paciente (Ej: 20:00):",
              value=(datetime.now() + timedelta(hours=2)).time(),
              key="h_nuev_input",
          )

          detalle_motivo = st.text_area(
              "Motivo detallado del cambio de horario (Obligatorio):",
              value=(
                  f"Paciente solicita reprogramar visita de las"
                  f" {hora_original.strftime('%H:%M')} para las"
                  f" {hora_nueva.strftime('%H:%M')}."
              ),
              key="det_h_input",
          )

          resp_firma = st.text_input(
              "✍️ Nombre del Profesional Responsable (Obligatorio):",
              key="resp_nov_h_tab",
          )

          if st.button(
              "🚀 Aplicar Cambio de Horario y Activar Alerta", key="btn_aplica_h"
          ):
            if not resp_firma.strip():
              st.error("❌ Debe ingresar su nombre como responsable.")
            else:
              fecha_creacion_str = datetime.now().isoformat()
              fec_aplica_str = fecha_aplica.strftime("%d/%m/%Y")

              run_query(
                  """INSERT INTO registro_novedades (documento, tipo_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                             VALUES (?, ?, ?, ?, ?, ?, 'Pendiente', ?)""",
                  (
                      doc_nov,
                      tipo_nov,
                      fec_aplica_str,
                      hora_aplica.strftime("%I:%M %p"),
                      detalle_motivo,
                      resp_firma.strip(),
                      fecha_creacion_str,
                  ),
                  fetch=False,
              )

              if "todo el ciclo permanente" in modo_cambio:
                nuevo_dt_inicio = datetime.combine(
                    datetime.fromisoformat(t_ini).date(), hora_nueva
                )
                try:
                  mapa_admin_ant = json.loads(t_dist_json)
                except Exception:
                  mapa_admin_ant = {}

                frec_h = safe_int(t_frec, 8)
                dias_h = safe_int(t_dias, 5)
                d_base_ant = (
                    safe_int(t_data[0][15], dias_h * 24 // frec_h)
                    if len(t_data[0]) > 15
                    else (dias_h * 24 // frec_h)
                )
                ajuste_h = (
                    safe_int(t_data[0][16], 0) if len(t_data[0]) > 16 else 0
                )
                mot_ajuste_ant = (
                    t_data[0][17] if len(t_data[0]) > 17 else ""
                )

                (
                    nueva_fecha_fin,
                    nuevo_t1,
                    nuevo_t2,
                    nuevo_t3,
                    _,
                    tot_d,
                    _,
                    _,
                ) = calcular_tratamiento_mixto(
                    nuevo_dt_inicio, dias_h, frec_h, mapa_admin_ant, ajuste_h
                )

                nueva_nov_txt = (
                    f"📢 NOVEDAD PENDIENTE [HORARIO MODIFICADO]: Visitar a las"
                    f" {hora_nueva.strftime('%H:%M')} (Antes"
                    f" {hora_original.strftime('%H:%M')}) - Autoriza:"
                    f" {resp_firma.strip()}"
                )

                run_query(
                    "UPDATE tratamientos SET estado = 'Modificado' WHERE id = ?",
                    (t_id,),
                    fetch=False,
                )

                run_query(
                    """
                                  INSERT INTO tratamientos (
                                      documento, tratamiento, dosis, via_administracion, tipo_acceso, frecuencia_horas, 
                                      dias, fecha_inicio, fecha_fin, t1, t2, t3, distribucion_admin, visita_educacion, 
                                      dosis_base_fijas, dosis_perdidas, motivo_ajuste_dosis, novedades, usuario_modificacion, 
                                      estado, descanalizado, alerta_revisada, fecha_registro
                                  ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', 'NO', 'NO', ?)
                              """,
                    (
                        doc_nov,
                        t_med,
                        t_dos,
                        t_via,
                        t_acc,
                        frec_h,
                        dias_h,
                        nuevo_dt_inicio.isoformat(),
                        nueva_fecha_fin.isoformat(),
                        nuevo_t1,
                        nuevo_t2,
                        nuevo_t3,
                        t_dist_json,
                        t_data[0][14],
                        d_base_ant,
                        ajuste_h,
                        mot_ajuste_ant,
                        nueva_nov_txt,
                        resp_firma.strip(),
                        fecha_creacion_str,
                    ),
                    fetch=False,
                )

                st.success(
                    "✅ Horario recalculado en todo el ciclo y activada alerta en"
                    " planilla."
                )
              else:
                aviso_puntual = (
                    f"📢 NOVEDAD PENDIENTE [VISITA PUNTUAL {fec_aplica_str}]:"
                    f" Visitar a las {hora_nueva.strftime('%H:%M')} (Original:"
                    f" {hora_original.strftime('%H:%M')}) - Autoriza:"
                    f" {resp_firma.strip()}"
                )
                run_query(
                    "UPDATE tratamientos SET novedades = ?,"
                    " usuario_modificacion = ?, alerta_revisada = 'NO' WHERE id"
                    " = ?",
                    (aviso_puntual, resp_firma.strip(), t_id),
                    fetch=False,
                )
                st.success(
                    f"✅ Alerta de visita puntual para el {fec_aplica_str}"
                    " activada en planilla."
                )

              st.rerun()

        else:
          detalle_general = st.text_area(
              "Descripción de la Novedad / Motivo:",
              value=f"Se registra {tipo_nov}.",
              key="det_gen_input",
          )
          resp_firma_gen = st.text_input(
              "✍️ Nombre del Profesional Responsable (Obligatorio):",
              key="resp_nov_gen_tab",
          )

          if st.button(
              "💾 Guardar Novedad y Activar Alerta en Planilla",
              key="btn_guarda_gen",
          ):
            if not resp_firma_gen.strip():
              st.error("❌ Ingrese su nombre como responsable.")
            else:
              fecha_creacion_str = datetime.now().isoformat()
              fec_aplica_str = fecha_aplica.strftime("%d/%m/%Y")

              run_query(
                  """INSERT INTO registro_novedades (documento, tipo_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion)
                             VALUES (?, ?, ?, ?, ?, ?, 'Pendiente', ?)""",
                  (
                      doc_nov,
                      tipo_nov,
                      fec_aplica_str,
                      hora_aplica.strftime("%I:%M %p"),
                      detalle_general,
                      resp_firma_gen.strip(),
                      fecha_creacion_str,
                  ),
                  fetch=False,
              )

              if t_data:
                prefijo = (
                    "🚨 [NOVEDAD DE SALIDA]"
                    if "Salida" in tipo_nov
                    else "📢 NOVEDAD PENDIENTE"
                )
                nueva_nov_t = (
                    f"{prefijo} [{tipo_nov.upper()} {fec_aplica_str}]:"
                    f" {detalle_general} - Firma: {resp_firma_gen.strip()}"
                )
                run_query(
                    "UPDATE tratamientos SET novedades = ?,"
                    " usuario_modificacion = ?, alerta_revisada = 'NO' WHERE id"
                    " = ?",
                    (nueva_nov_t, resp_firma_gen.strip(), t_data[0][0]),
                    fetch=False,
                )

              st.success(
                  "✅ Novedad registrada y visible como ALERTA en la planilla."
              )
              st.rerun()

      st.markdown("---")
      st.markdown("#### 📜 Histórico de Novedades Registradas de este Paciente")
      novs_paciente = run_query(
          """SELECT tipo_novedad, fecha_aplicacion, hora_aplicacion, detalle, responsable, estado_novedad, fecha_creacion 
                    FROM registro_novedades WHERE documento = ? ORDER BY id DESC""",
          (doc_nov,),
      )
      if novs_paciente:
        df_novs = pd.DataFrame(
            novs_paciente,
            columns=[
                "Tipo Novedad",
                "Fecha Aplica",
                "Hora Aplica",
                "Detalle / Motivo",
                "Responsable",
                "Estado",
                "Fecha Registro Sistema",
            ],
        )
        st.dataframe(df_novs, use_container_width=True)
      else:
        st.info("No hay novedades registradas previamente para este paciente.")

# ----------------- SECCIÓN 3: BASE EN VIVO / PLANILLA -----------------
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
            p.tipo_aislamiento AS 'Tipo Aislamiento',
            COALESCE(t.tratamiento, 'Sin Tratamiento') AS 'Tratamiento',
            COALESCE(t.dosis, '-') AS 'Dosis',
            COALESCE(t.via_administracion, '-') AS 'Vía',
            COALESCE(t.tipo_acceso, 'Periférico') AS 'Tipo Acceso',
            COALESCE(t.frecuencia_horas, 0) AS 'Frec (h)',
            COALESCE(t.dias, 0) AS 'Días',
            COALESCE(t.visita_educacion, '') AS 'Visita Educación',
            COALESCE(t.dosis_base_fijas, 0) AS 'Dosis Base Fijas',
            COALESCE(t.dosis_perdidas, 0) AS 'Ajuste Dosis (+/-)',
            COALESCE(t.motivo_ajuste_dosis, '') AS 'Motivo Ajuste Dosis',
            COALESCE(t.fecha_inicio, '-') AS 'Fecha Inicio',
            COALESCE(t.fecha_fin, '-') AS 'Fecha Fin',
            COALESCE(t.t1, '-') AS 'T1',
            COALESCE(t.t2, '-') AS 'T2',
            COALESCE(t.t3, '-') AS 'T3',
            COALESCE(t.usuario_modificacion, '-') AS 'Gestor Responsable',
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

    c_f1, c_f2, c_f3, c_f4, c_f5, c_f6, c_f7 = st.columns(7)
    filtro_zona = c_f1.multiselect(
        "Zona", options=df_general["Zona"].dropna().unique()
    )
    filtro_plan = c_f2.multiselect(
        "Plan (POS/Póliza/ARL)", options=df_general["Plan"].dropna().unique()
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
        "Alertas / Novedades",
        options=df_general["Estado Tratamiento"].dropna().unique(),
    )
    filtro_turno = c_f7.selectbox(
        "Filtro Turno:",
        ["Todos", "T1 (SENC)", "T2 (SENC)", "T3 (SENC)", "Con Cuidador"],
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

    if filtro_turno == "T1 (SENC)":
      df_filtrado = df_filtrado[
          df_filtrado["T1"].str.contains("H", na=False)
          & ~df_filtrado["T1"].str.contains("CUIDADOR", na=False)
      ]
    elif filtro_turno == "T2 (SENC)":
      df_filtrado = df_filtrado[
          df_filtrado["T2"].str.contains("H", na=False)
          & ~df_filtrado["T2"].str.contains("CUIDADOR", na=False)
      ]
    elif filtro_turno == "T3 (SENC)":
      df_filtrado = df_filtrado[
          df_filtrado["T3"].str.contains("H", na=False)
          & ~df_filtrado["T3"].str.contains("CUIDADOR", na=False)
      ]
    elif filtro_turno == "Con Cuidador":
      df_filtrado = df_filtrado[
          df_filtrado["T1"].str.contains("CUIDADOR", na=False)
          | df_filtrado["T2"].str.contains("CUIDADOR", na=False)
          | df_filtrado["T3"].str.contains("CUIDADOR", na=False)
      ]

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
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#721C24;'>{total_picc}</div><div"
        " class='metric-lbl'>🩸 CON CATÉTER PICC</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#5D4037;'>{total_npt}</div><div"
        " class='metric-lbl'>🥣 CON NPT</div></div>",
        unsafe_allow_html=True,
    )
    m4.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#1B5E20;'>{total_be}</div><div"
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

    # ----------------- CONVENCIONES DE COLOR COMPACTAS (HORIZONTAL) -----------------
    st.markdown("##### 🎨 Convenciones de Color en la Planilla")
    c_c1, c_c2, c_c3, c_c4 = st.columns(4)
    with c_c1:
      st.markdown(
          "<div class='conv-box' style='background-color: #ffd6d6; color:"
          " #721c24;'>🟥 Catéter PICC</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<div class='conv-box' style='background-color: #fff8e1; color:"
          " #5d4037;'>🟨 Nutrición Parenteral (NPT)</div>",
          unsafe_allow_html=True,
      )
    with c_c2:
      st.markdown(
          "<div class='conv-box' style='background-color: #e8f5e9; color:"
          " #1b5e20;'>🟩 Bomba Elastomérica (BE)</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<div class='conv-box' style='background-color: #ede7f6; color:"
          " #4a148c;'>🟪 BE Analgesia / Dolor</div>",
          unsafe_allow_html=True,
      )
    with c_c3:
      st.markdown(
          "<div class='conv-box' style='background-color: #ffe0b2; color:"
          " #bf360c;'>🟧 Tuberculosis (TB) / Novedad</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<div class='conv-box' style='background-color: #e0f7fa; color:"
          " #006064;'>🟦 Vía PUSH</div>",
          unsafe_allow_html=True,
      )
    with c_c4:
      st.markdown(
          "<div class='conv-box' style='background-color: #e1bee7; color:"
          " #4a148c;'>🟣 Paciente Aislado</div>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<div class='conv-box' style='background-color: #ffcc80; color:"
          " #b71c1c;'>⚠️ Falta 1 Dosis / Retiro</div>",
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.dataframe(
        df_filtrado.style.apply(estilo_filas, axis=1), use_container_width=True
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_filtrado.to_excel(writer, sheet_name="Planilla_SURA")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Planilla Filtrada (.xlsx)",
        data=excel_data,
        file_name=(
            f"Planilla_SENC_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.info("No hay registros bajo el estado o filtros seleccionados.")

# ----------------- SECCIÓN 4: RESPALDO GENERAL EN EXCEL -----------------
elif menu == "💾 Respaldo General en Excel":
  st.subheader("💾 Copia de Seguridad de la Base de Datos en Excel")
  opcion_backup = st.radio(
      "Modalidad de Respaldo:",
      [
          "1. Planilla Consolidada Limpia (Solo pacientes y tratamientos"
          " ACTIVOS, sin historial)",
          "2. Planilla Completa con Historial de Cambios y Auditoría (Todos los"
          " movimientos y versiones)",
      ],
      index=0,
  )

  col_b1, col_b2 = st.columns([2, 1])

  with col_b1:
    if "Solo pacientes y tratamientos ACTIVOS" in opcion_backup:
      if st.button("📊 Descargar Planilla Limpia (Sin Historial)"):
        query_limpia = """
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
                        p.tipo_aislamiento AS 'Tipo Aislamiento',
                        t.tratamiento AS 'Medicamento',
                        t.dosis AS 'Dosis',
                        t.via_administracion AS 'Vía Adm',
                        t.tipo_acceso AS 'Tipo Acceso',
                        t.frecuencia_horas AS 'Frecuencia (h)',
                        t.dias AS 'Días TTO',
                        t.dosis_base_fijas AS 'Dosis Base Prescritas',
                        t.dosis_perdidas AS 'Ajuste Dosis (+/-)',
                        t.motivo_ajuste_dosis AS 'Motivo Ajuste',
                        t.fecha_inicio AS 'Fecha Inicio',
                        t.fecha_fin AS 'Fecha Fin',
                        t.t1 AS 'T1',
                        t.t2 AS 'T2',
                        t.t3 AS 'T3',
                        t.visita_educacion AS 'Visita Educación',
                        t.novedades AS 'NOVEDAD',
                        t.usuario_modificacion AS 'Gestor Responsable',
                        t.fecha_registro AS 'Fecha Última Modificación'
                    FROM pacientes p
                    INNER JOIN tratamientos t ON p.documento = t.documento
                    WHERE t.estado = 'Activo'
                    ORDER BY p.zona ASC, p.nombre ASC
                """
        conn = sqlite3.connect(DB_FILE)
        df_limpia = pd.read_sql_query(query_limpia, conn)
        conn.close()

        output_b = io.BytesIO()
        with pd.ExcelWriter(output_b, engine="openpyxl") as writer:
          df_limpia.to_excel(
              writer, index=False, sheet_name="Censo_Activo_Limpio"
          )

        b_data = output_b.getvalue()
        b_nombre = (
            "CENSO_ACTIVO_LIMPIO_SURA_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        st.download_button(
            label="⬇️ Descargar Archivo Excel (.xlsx)",
            data=b_data,
            file_name=b_nombre,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    else:
      if st.button("📊 Descargar con Historial Completo y Auditoría"):
        query_historial = """
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
                        p.tipo_aislamiento AS 'Tipo Aislamiento',
                        t.tratamiento AS 'Medicamento',
                        t.dosis AS 'Dosis',
                        t.via_administracion AS 'Vía Adm',
                        t.tipo_acceso AS 'Tipo Acceso',
                        t.frecuencia_horas AS 'Frecuencia (h)',
                        t.dias AS 'Días TTO',
                        t.dosis_base_fijas AS 'Dosis Base Prescritas',
                        t.dosis_perdidas AS 'Ajuste Dosis (+/-)',
                        t.motivo_ajuste_dosis AS 'Motivo Ajuste',
                        t.fecha_inicio AS 'Fecha Inicio',
                        t.fecha_fin AS 'Fecha Fin',
                        t.t1 AS 'T1',
                        t.t2 AS 'T2',
                        t.t3 AS 'T3',
                        t.visita_educacion AS 'Visita Educación',
                        t.novedades AS 'NOVEDAD',
                        t.estado AS 'Estado TTO',
                        t.descanalizado AS 'Descanalizado',
                        t.alerta_revisada AS 'Alerta Atendida',
                        t.detalle_revision_alerta AS 'Acción Alerta',
                        t.responsable_revision_alerta AS 'Gestionado Por',
                        t.usuario_modificacion AS 'Gestor Que Guardó',
                        t.fecha_registro AS 'Fecha Movimiento'
                    FROM pacientes p
                    LEFT JOIN tratamientos t ON p.documento = t.documento
                    ORDER BY p.zona ASC, p.documento ASC, t.id DESC
                """
        conn = sqlite3.connect(DB_FILE)
        df_hist = pd.read_sql_query(query_historial, conn)
        conn.close()

        output_b = io.BytesIO()
        with pd.ExcelWriter(output_b, engine="openpyxl") as writer:
          df_hist.to_excel(
              writer, index=False, sheet_name="Historial_Completo_SURA"
          )

        b_data = output_b.getvalue()
        b_nombre = (
            "HISTORIAL_AUDITORIA_SURA_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        st.download_button(
            label="⬇️ Descargar Archivo Excel (.xlsx)",
            data=b_data,
            file_name=b_nombre,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

# ----------------- PIE DE PÁGINA / AUTORÍA AL FINAL DE LA PÁGINA -----------------
st.markdown(
    """
    <div class="footer-autor">
        🏥 <strong>Sistema de Gestión y Control de Tratamientos Domiciliarios (SENC)</strong><br>
        Diseñado y desarrollado por <strong>Andrés Medina</strong> · Salud en Casa SURA Colombia
    </div>
""",
    unsafe_allow_html=True,
)
