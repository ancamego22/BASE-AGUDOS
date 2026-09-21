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
        padding: 1rem;
        box-shadow: 0px 3px 10px rgba(0,0,0,0.04);
        border: 1px solid #E6EAF0;
        text-align: center;
    }
    .metric-val {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0033A0;
    }
    .metric-lbl {
        font-size: 0.82rem;
        color: #6C757D;
        font-weight: 600;
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
            dosis_perdidas INTEGER DEFAULT 0,
            novedades TEXT,
            usuario_modificacion TEXT DEFAULT '',
            estado TEXT DEFAULT 'Activo',
            descanalizado TEXT DEFAULT 'NO',
            fecha_registro TEXT NOT NULL,
            fecha_retiro_cateter TEXT,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

  c.execute("PRAGMA table_info(pacientes)")
  cols_pacientes = [col[1] for col in c.fetchall()]
  if "municipio" not in cols_pacientes:
    c.execute(
        "ALTER TABLE pacientes ADD COLUMN municipio TEXT DEFAULT 'Medellín'"
    )
  if "barrio" not in cols_pacientes:
    c.execute("ALTER TABLE pacientes ADD COLUMN barrio TEXT DEFAULT ''")
  if "aislamiento" not in cols_pacientes:
    c.execute(
        "ALTER TABLE pacientes ADD COLUMN aislamiento TEXT DEFAULT 'NO'"
    )
  if "tipo_aislamiento" not in cols_pacientes:
    c.execute(
        "ALTER TABLE pacientes ADD COLUMN tipo_aislamiento TEXT DEFAULT"
        " 'Ninguno'"
    )
  if "usuario_registro" not in cols_pacientes:
    c.execute("ALTER TABLE pacientes ADD COLUMN usuario_registro TEXT")

  c.execute("PRAGMA table_info(tratamientos)")
  cols_trat = [col[1] for col in c.fetchall()]
  columnas_nuevas = {
      "tipo_acceso": "TEXT DEFAULT 'Periférico'",
      "t1": "TEXT",
      "t2": "TEXT",
      "t3": "TEXT",
      "distribucion_admin": "TEXT DEFAULT '{}'",
      "visita_educacion": "TEXT DEFAULT ''",
      "dosis_perdidas": "INTEGER DEFAULT 0",
      "novedades": "TEXT",
      "usuario_modificacion": "TEXT DEFAULT ''",
      "descanalizado": "TEXT DEFAULT 'NO'",
      "fecha_retiro_cateter": "TEXT",
  }
  for col_name, col_def in columnas_nuevas.items():
    if col_name not in cols_trat:
      c.execute(f"ALTER TABLE tratamientos ADD COLUMN {col_name} {col_def}")

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


# ----------------- CÁLCULOS CLÍNICOS EXACTOS (SIN ALTERAR DOSIS) -----------------
def obtener_horas_ciclo(dt_inicio, frecuencia_horas):
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
  # 1. Total exacto de dosis clínicas que el paciente debe recibir
  total_dosis_base = int((dias * 24) / frecuencia_horas)
  total_dosis = max(1, total_dosis_base + ajuste_dosis)

  # 2. La fecha fin depende estrictamente del total de dosis, sin importar quién la aplique
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

  # Desglose total de dosis durante todo el tratamiento
  dosis_totales_senc = dosis_diarias_senc * dias
  dosis_totales_cuidador = dosis_diarias_cuidador * dias

  return (
      fecha_fin,
      t1_str,
      t2_str,
      t3_str,
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
):
  if estado == "Completado" or descanalizado == "SI":
    return "✅ DESCANALIZADO / CERRADO"

  try:
    dt_fin = datetime.fromisoformat(fecha_fin_str)
    ahora = datetime.now()
    horas_restantes = (dt_fin - ahora).total_seconds() / 3600

    if horas_restantes <= frecuencia_horas:
      return "⚠️ FALTA 1 DOSIS (DESCANALIZAR)"
    elif horas_restantes <= 48:
      return "🔔 PRÓXIMO A FIN (<=48h - Programar)"
    elif novedades and (
        "SALIDA" in novedades.upper() or "RETIRA" in novedades.upper()
    ):
      return "🚨 NOVEDAD DE SALIDA"
    return "En curso"
  except Exception:
    return "En curso"


def estilo_filas(row):
  estado = str(row.get("Estado", ""))
  acceso = str(row.get("Tipo Acceso", "")).upper()
  alerta = str(row.get("Estado Tratamiento", ""))
  aislamiento = str(row.get("Aislamiento", "")).upper()

  if "DESCANALIZADO" in alerta or estado == "Completado":
    return ["color: #6C757D; font-style: italic;"] * len(row)

  if "FALTA 1 DOSIS" in alerta:
    return [
        "background-color: #ffcc80; color: #b71c1c; font-weight: bold;"
    ] * len(row)

  if "NOVEDAD DE SALIDA" in alerta:
    return [
        "background-color: #ffcdd2; color: #b71c1c; font-weight: bold;"
    ] * len(row)

  if acceso == "PICC":
    return [
        "background-color: #ffd6d6; color: #721c24; font-weight: bold;"
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
          "📊 Censo y Planilla en Vivo",
          "💾 Respaldo General en Excel",
      ],
  )
  st.markdown("---")
  st.caption("🔒 Acceso Seguro · SURA Colombia")

# ----------------- ENCABEZADO SUPERIOR -----------------
st.markdown(
    f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Panel de Control Clínico y Cronograma SENC</h2>
            <p class="hospital-sub">Programa de Hospitalización Domiciliaria Agudos · SURA</p>
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

  if doc_busqueda:
    paciente_data = run_query(
        "SELECT * FROM pacientes WHERE documento = ?", (doc_busqueda,)
    )

    with st.expander(
        "👤 Información Demográfica y Aislamiento del Paciente",
        expanded=not bool(paciente_data),
    ):
      col1, col2, col3 = st.columns(3)

      programas_disponibles = [
          "Agudos",
          "BE Analgesia",
          "PUSH",
          "Tuberculosis",
          "Materna",
          "BK Inducido",
      ]
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
        programa = col3.selectbox(
            "Programa",
            programas_disponibles,
            index=(
                programas_disponibles.index(p_prog)
                if p_prog in programas_disponibles
                else 0
            ),
        )

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
            "¿El paciente tiene algún tipo de aislamiento? (Obligatorio)",
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
            st.success("Datos demográficos actualizados en tiempo real.")
          else:
            st.error("Debes ingresar tu nombre como responsable.")
      else:
        st.info("Paciente no registrado. Digite los datos para admisión:")
        nombre = col1.text_input("Nombre y Apellido")
        plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"])
        programa = col3.selectbox("Programa", programas_disponibles)

        col4, col5, col6, col7 = st.columns(4)
        municipio = col4.selectbox("Municipio", municipios_disponibles)
        barrio = col5.text_input("Barrio")
        piso = col6.selectbox("Piso", pisos)
        zona = col7.selectbox("Zona", zonas)

        st.markdown("#### ☣️ Medidas de Bioseguridad y Aislamiento")
        c_ais1, c_ais2 = st.columns(2)
        tiene_aislamiento = c_ais1.radio(
            "¿El paciente tiene algún tipo de aislamiento? (Obligatorio)",
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

      # ----------------- PANEL DE DESCANALIZACIÓN / CIERRE -----------------
      if tratamiento_activo:
        with st.container():
          st.markdown(
              "### 🏁 Cierre de Tratamiento y Confirmación de Descanalización"
          )
          with st.expander(
              "✅ Confirmar que el paciente YA SE DESCANALIZÓ (Retiro de"
              " Catéter)"
          ):
            c_cierre1, c_cierre2 = st.columns(2)
            profesional_cierre = c_cierre1.text_input(
                "Profesional que confirma:", key="prof_desc"
            )
            fecha_cierre = c_cierre2.date_input(
                "Fecha de Retiro / Descanalización:", value=datetime.now().date()
            )
            nota_cierre = st.text_input(
                "Novedad de cierre:",
                value="PACIENTE DESCANALIZADO - FIN TTO",
                key="nota_desc",
            )

            if st.button("Confirmar Descanalización y Retirar Alerta"):
              if not profesional_cierre.strip():
                st.error("Debe ingresar el nombre del profesional que confirma.")
              else:
                fecha_retiro_str = (
                    f"{fecha_cierre.strftime('%d/%m/%Y')} por"
                    f" {profesional_cierre.strip()}"
                )
                nov_actual = (
                    tratamiento_activo[0][16]
                    if len(tratamiento_activo[0]) > 16
                    and tratamiento_activo[0][16]
                    else ""
                )
                nueva_novedad = (
                    f"{nov_actual} // [DESCANALIZADO]: {nota_cierre} ("
                    f" {fecha_retiro_str})".strip()
                )

                run_query(
                    """UPDATE tratamientos SET estado = 'Completado', descanalizado = 'SI', 
                               novedades = ?, fecha_retiro_cateter = ? WHERE id = ?""",
                    (nueva_novedad, fecha_retiro_str, tratamiento_activo[0][0]),
                    fetch=False,
                )
                st.success(
                    "✅ Paciente descanalizado con éxito. Alerta retirada del"
                    " sistema."
                )
                st.rerun()

      # ----------------- FORMULACIÓN DE TRATAMIENTO -----------------
      st.subheader("💊 Formulación / Modificación de Tratamiento")

      vias = ["IV", "SC", "IM", "BE", "NPT"]
      accesos = ["Periférico", "PICC"]

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
      frec_val = tratamiento_activo[0][6] if tratamiento_activo else 8
      frecuencia = col_f1.selectbox(
          "Frecuencia (Horas)",
          frecuencias,
          index=frecuencias.index(frec_val) if frec_val in frecuencias else 2,
      )
      dias = col_f2.number_input(
          "Días de Tratamiento",
          min_value=1,
          max_value=90,
          value=tratamiento_activo[0][7] if tratamiento_activo else 5,
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

      # ----------------- ASIGNACIÓN EXACTA DE DOSIS -----------------
      st.markdown("#### 🕒 Asignación de Responsable por Dosis del Día")
      st.caption(
          "Las dosis por Cuidador se contabilizan en el total del tratamiento"
          " (días y fin exactos), pero no saturan la ruta física de SENC."
      )

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
        st.info(
            "👨‍⚕️ **Dosis asignadas a Cuidador:** Puede programar una visita de"
            " educación de enfermería."
        )
        col_ed1, col_ed2 = st.columns(2)
        fecha_ed = col_ed1.date_input(
            "Fecha Visita Educación:", value=datetime.now().date()
        )
        hora_ed = col_ed2.time_input(
            "Hora Visita Educación:", value=datetime.now().time()
        )
        visita_educacion = f"{fecha_ed.strftime('%d/%m/%Y')} {hora_ed.strftime('%I:%M %p')}"

      col_rep1, col_rep2 = st.columns([1, 2])
      ajuste_dosis = col_rep1.number_input(
          "Ajuste Dosis (+ Sumar / - Restar)",
          min_value=-30,
          max_value=30,
          value=tratamiento_activo[0][15] if tratamiento_activo else 0,
      )
      novedades = col_rep2.text_area(
          "NOVEDAD (PICC, Salida, Prórroga, Familiar Retira)",
          value=tratamiento_activo[0][16] if tratamiento_activo else "",
      )

      usuario_cambio = st.text_input(
          "✍️ Nombre del Profesional que realiza el cambio (Obligatorio):",
          value="",
      )

      (
          fecha_fin_calculada,
          t1_calc,
          t2_calc,
          t3_calc,
          dosis_totales,
          d_senc,
          d_cuid,
      ) = calcular_tratamiento_mixto(
          dt_inicio, dias, frecuencia, nuevo_mapa_admin, ajuste_dosis
      )

      # ----------------- CUADRO DE CONCILIACIÓN EXACTA DE DOSIS -----------------
      st.markdown("#### ⚖️ Conciliación y Balance de Dosis Totales")
      mb1, mb2, mb3, mb4 = st.columns(4)
      mb1.metric("Dosis Totales Requeridas", f"{dosis_totales} Dosis")
      mb2.metric("Aportadas por SENC", f"{d_senc} Dosis")
      mb3.metric("Aportadas por Cuidador", f"{d_cuid} Dosis")
      mb4.metric("Fecha Fin Exacta", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))

      st.markdown("#### 📋 Horarios de Visitas Calculados")
      c_p1, c_p2, c_p3 = st.columns(3)
      c_p1.metric("Turno T1 (06H-14H)", t1_calc)
      c_p2.metric("Turno T2 (14H-22H)", t2_calc)
      c_p3.metric("Turno T3 (22H-06H)", t3_calc)

      alerta_fin = evaluar_alerta_fin(
          fecha_fin_calculada.isoformat(), frecuencia, "Activo", "NO", novedades
      )
      if "FALTA 1 DOSIS" in alerta_fin:
        st.error(
            f"⚠️ **ALERTA CRÍTICA:** {alerta_fin}. Se debe coordinar retiro"
            " inmediato de catéter."
        )
      elif "NOVEDAD DE SALIDA" in alerta_fin:
        st.warning(
            "🚨 **ALERTA DE SALIDA:** Paciente con novedad activa de salida."
        )

      if st.button("💾 Guardar Tratamiento con Firma"):
        if not usuario_cambio.strip():
          st.error("❌ Ingrese su nombre como profesional responsable.")
        elif not t_nombre.strip():
          st.error("❌ Ingrese el medicamento / tratamiento.")
        else:
          fecha_reg_str = datetime.now().isoformat()
          fecha_ini_str = dt_inicio.isoformat()
          fecha_fin_str = fecha_fin_calculada.isoformat()
          json_admin = json.dumps(nuevo_mapa_admin)

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
                            dosis_perdidas, novedades, usuario_modificacion, estado, descanalizado, fecha_registro
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', 'NO', ?)
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
                  ajuste_dosis,
                  novedades,
                  usuario_cambio.strip(),
                  fecha_reg_str,
              ),
              fetch=False,
          )

          st.success(
              f"Tratamiento registrado y visible en tiempo real por:"
              f" {usuario_cambio}."
          )
          st.rerun()

      if tratamiento_activo:
        st.write("### 📌 Tratamiento Activo Actual")
        estado_alerta_act = evaluar_alerta_fin(
            tratamiento_activo[0][9],
            tratamiento_activo[0][6],
            tratamiento_activo[0][18]
            if len(tratamiento_activo[0]) > 18
            else "Activo",
            tratamiento_activo[0][19]
            if len(tratamiento_activo[0]) > 19
            else "NO",
            tratamiento_activo[0][16],
        )

        df_activo = pd.DataFrame([{
            "Tratamiento": tratamiento_activo[0][2],
            "Dosis": tratamiento_activo[0][3],
            "Vía": tratamiento_activo[0][4],
            "Tipo Acceso": tratamiento_activo[0][5],
            "Frecuencia": f"C/{tratamiento_activo[0][6]}h",
            "Días": tratamiento_activo[0][7],
            "Fecha Fin": datetime.fromisoformat(
                tratamiento_activo[0][9]
            ).strftime("%d/%m/%Y %I:%M %p"),
            "T1": tratamiento_activo[0][10],
            "T2": tratamiento_activo[0][11],
            "T3": tratamiento_activo[0][12],
            "Visita Educación": (
                tratamiento_activo[0][14]
                if len(tratamiento_activo[0]) > 14
                else ""
            ),
            "Ajuste Dosis": tratamiento_activo[0][15],
            "Estado Tratamiento": estado_alerta_act,
            "Registrado Por": (
                tratamiento_activo[0][17]
                if len(tratamiento_activo[0]) > 17
                else ""
            ),
            "NOVEDAD": tratamiento_activo[0][16],
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

# ----------------- SECCIÓN 2: BASE EN VIVO / PLANILLA -----------------
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
            COALESCE(t.fecha_inicio, '-') AS 'Fecha Inicio',
            COALESCE(t.fecha_fin, '-') AS 'Fecha Fin',
            COALESCE(t.t1, '-') AS 'T1',
            COALESCE(t.t2, '-') AS 'T2',
            COALESCE(t.t3, '-') AS 'T3',
            COALESCE(t.dosis_perdidas, 0) AS 'Ajuste Dosis',
            COALESCE(t.usuario_modificacion, '-') AS 'Gestor Responsable',
            COALESCE(t.novedades, '') AS 'NOVEDAD',
            COALESCE(t.estado, 'Sin TTO') AS 'Estado',
            COALESCE(t.descanalizado, 'NO') AS 'Descanalizado'
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
        )
        if row["Fecha Fin"] != "-"
        else "Sin tratamiento",
        axis=1,
    )

    c_f1, c_f2, c_f3, c_f4, c_f5, c_f6, c_f7 = st.columns(7)
    filtro_zona = c_f1.multiselect(
        "Zona", options=df_general["Zona"].dropna().unique()
    )
    filtro_muni = c_f2.multiselect(
        "Municipio", options=df_general["Municipio"].dropna().unique()
    )
    filtro_prog = c_f3.multiselect(
        "Programa", options=df_general["Programa"].dropna().unique()
    )
    filtro_acceso = c_f4.multiselect(
        "Tipo Acceso", options=df_general["Tipo Acceso"].dropna().unique()
    )
    filtro_aisla = c_f5.multiselect(
        "Aislamiento", options=df_general["Aislamiento"].dropna().unique()
    )
    filtro_alerta = c_f6.multiselect(
        "Alertas Retiro/Salida",
        options=df_general["Estado Tratamiento"].dropna().unique(),
    )
    filtro_turno = c_f7.selectbox(
        "Filtro por Turno Visita:",
        ["Todos", "T1 (SENC)", "T2 (SENC)", "T3 (SENC)", "Con Cuidador"],
    )

    df_filtrado = df_general.copy()
    if filtro_zona:
      df_filtrado = df_filtrado[df_filtrado["Zona"].isin(filtro_zona)]
    if filtro_muni:
      df_filtrado = df_filtrado[df_filtrado["Municipio"].isin(filtro_muni)]
    if filtro_prog:
      df_filtrado = df_filtrado[df_filtrado["Programa"].isin(filtro_prog)]
    if filtro_acceso:
      df_filtrado = df_filtrado[df_filtrado["Tipo Acceso"].isin(filtro_acceso)]
    if filtro_aisla:
      df_filtrado = df_filtrado[df_filtrado["Aislamiento"].isin(filtro_aisla)]
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
    total_descanalizar = len(
        df_filtrado[
            df_filtrado["Estado Tratamiento"].str.contains("FALTA 1 DOSIS")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_aislados = len(
        df_filtrado[
            (df_filtrado["Aislamiento"] == "SI")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )
    total_picc = len(
        df_filtrado[
            (df_filtrado["Tipo Acceso"].str.upper() == "PICC")
            & (df_filtrado["Estado"] == "Activo")
        ]
    )

    m1, m2, m3, m4 = st.columns(4)
    zona_txt = f" ({', '.join(filtro_zona)})" if filtro_zona else " TOTAL"
    m1.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_filtrados}</div><div"
        f" class='metric-lbl'>PACIENTES ACTIVOS EN ZONA{zona_txt}</div></div>",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#B71C1C;'>{total_descanalizar}</div><div"
        " class='metric-lbl'>⚠️ PENDIENTES DESCANALIZAR (FALTA 1"
        " DOSIS)</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#4A148C;'>{total_aislados}</div><div"
        " class='metric-lbl'>☣️ PACIENTES CON AISLAMIENTO</div></div>",
        unsafe_allow_html=True,
    )
    m4.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#721C24;'>{total_picc}</div><div"
        " class='metric-lbl'>PACIENTES CON PICC</div></div>",
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

# ----------------- SECCIÓN 3: RESPALDO GENERAL EN EXCEL -----------------
elif menu == "💾 Respaldo General en Excel":
  st.subheader("💾 Copia de Seguridad Completa en Excel (.xlsx)")
  st.markdown("""
    Descarga una copia completa de toda la base de datos estructurada en un único libro de Excel:
    * **Hoja 1:** Censo Demográfico Completo (Municipios, Barrios, Zonas, Aislamiento).
    * **Hoja 2:** Tratamientos Activos y Programación T1-T2-T3 (SENC vs. Cuidador).
    * **Hoja 3:** Trazabilidad Histórica de Modificaciones y Auditoría.
    """)

  if st.button("📊 Generar y Descargar Backup Consolidado en Excel"):
    conn = sqlite3.connect(DB_FILE)
    df_p = pd.read_sql_query("SELECT * FROM pacientes", conn)
    df_t = pd.read_sql_query(
        "SELECT * FROM tratamientos WHERE estado = 'Activo'", conn
    )
    df_h = pd.read_sql_query(
        "SELECT * FROM tratamientos ORDER BY id DESC", conn
    )
    conn.close()

    output_backup = io.BytesIO()
    with pd.ExcelWriter(output_backup, engine="openpyxl") as writer:
      df_p.to_excel(writer, index=False, sheet_name="Pacientes_Demografico")
      df_t.to_excel(writer, index=False, sheet_name="Tratamientos_Activos")
      df_h.to_excel(writer, index=False, sheet_name="Auditoria_Movimientos")

    b_data = output_backup.getvalue()
    b_nombre = (
        f"BACKUP_COMPLETO_SURA_SENC_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )

    st.download_button(
        label="⬇️ Descargar Archivo Excel de Respaldo (.xlsx)",
        data=b_data,
        file_name=b_nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
