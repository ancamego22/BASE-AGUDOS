import io
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

# ----------------- LOGO OFICIAL SURA Y ESTILOS TIPO DASHBOARD HOSPITALARIO -----------------
# Logo institucional SURA en SVG con el ave azul y turquesa
SVG_LOGO_SURA = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 240" width="100%" height="auto">
  <g fill="#0033A0">
    <path d="M42.2 135.8c-14.8-5.3-25.2-12.8-25.2-28.7 0-18.4 15.6-32.3 38.6-32.3 16.5 0 28.5 6.6 35.8 14.8l-12.6 13.8c-5.7-6.2-13.8-10.4-23.2-10.4-12.5 0-20.1 6.8-20.1 14.5 0 8.2 6.5 11.9 16.6 15.6l10.8 3.9c20.3 7.3 30.6 16.5 30.6 32.5 0 20.3-16.7 34.6-41.9 34.6-19.7 0-33.6-7.8-41.6-18.5l13.6-13.1c6.5 8 16.4 13.1 28 13.1 14.8 0 23.3-7.7 23.3-16.3 0-8.8-6.2-12.8-17.6-16.9l-16.6-6.5z"/>
    <path d="M129.5 77.2h18.2v67.8c0 19.3 10.9 28.9 28.1 28.9s28.1-9.6 28.1-28.9V77.2h18.2v67.8c0 30.4-19.3 46.8-46.3 46.8s-46.3-16.4-46.3-46.8V77.2z"/>
    <path d="M257.6 77.2h18.2v24.8c6.6-16.2 21.6-26.4 39.8-25.4v18.7c-21.7-1-39.8 11.2-39.8 35.6v59.3h-18.2V77.2z"/>
    <path d="M375.4 77.2h18.2v112h-18.2v-14.8c-8.5 10.6-21.1 16.8-35.8 16.8-26.4 0-46.1-20.3-46.1-47.8s19.7-47.8 46.1-47.8c14.7 0 27.3 6.2 35.8 16.8V77.2zm-33.8 20.2c-17.5 0-30.4 13.8-30.4 30.8s12.9 30.8 30.4 30.8c17.5 0 30.4-13.8 30.4-30.8s-12.9-30.8-30.4-30.8z"/>
  </g>
  <g fill="#00AEC7">
    <path d="M441 5.5c41.3 22 75.8 54.3 98.7 93.8-12.5-4.4-48-18.3-98.7-17.2 27.8-24.8 58.6-47.4 92.4-66.2-28.2-3.8-63.5-7.3-92.4-10.4z"/>
    <path d="M448 42.4c33.4 20.2 61.2 49.3 79.5 84.7-16-7.8-45.7-19-86.4-16.1 27.8-21.1 53.7-44.5 78.4-68.6h-71.5z"/>
    <path d="M455 79.2c25.6 18.5 46.8 44.3 60.5 75.6-18.2-11.2-46.6-19.8-77.5-15.5 24.2-18.3 47-38.6 68.3-60.1h-51.3z"/>
    <path d="M462 116c17.8 16.7 32.4 39.4 41.5 66.5-20.4-14.7-47.5-20.5-70.6-15.2 20.6-15.6 39.9-33 57.6-51.3h-28.5z"/>
    <path d="M472 153c10.1 14.8 18.1 34.6 22.5 57.5-22.6-18.1-48.4-21.3-65.7-14.8 17.5-12.8 33.7-27.1 48.2-42.7h-5z"/>
    <path d="M495 192c-15.2 12.5-33.3 21.8-52.6 27.2 16.5-6.8 31.6-16.1 44.5-27.2h8.1z"/>
    <path d="M532 188c-23.7 20.5-52.4 35.3-83.5 43.1 26.6-11.4 50.8-27.2 71.3-43.1h12.2z"/>
  </g>
</svg>
"""

st.markdown(
    """
    <style>
    /* Estilo Mediline / Dashboard Hospitalario */
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
    
    /* Contenedor Superior Corporativo */
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

    /* Tarjetas Métricas */
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
        font-size: 0.85rem;
        color: #6C757D;
        font-weight: 600;
    }

    /* Botón corporativo SURA */
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

DB_FILE = "control_pacientes.db"


# ----------------- BASE DE DATOS -----------------
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
            tipo_acceso TEXT NOT NULL,
            frecuencia_horas INTEGER NOT NULL,
            dias INTEGER NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_fin TEXT NOT NULL,
            t1 TEXT,
            t2 TEXT,
            t3 TEXT,
            dosis_perdidas INTEGER DEFAULT 0,
            novedades TEXT,
            usuario_modificacion TEXT NOT NULL,
            estado TEXT DEFAULT 'Activo',
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)
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


# ----------------- CÁLCULOS CLÍNICOS Y HORARIOS -----------------
def calcular_tratamiento(dt_inicio, dias, frecuencia_horas, ajuste_dosis=0):
  total_dosis_base = int((dias * 24) / frecuencia_horas)
  total_dosis = max(1, total_dosis_base + ajuste_dosis)

  fecha_fin = dt_inicio + timedelta(hours=frecuencia_horas * (total_dosis - 1))

  dosis_diarias = 24 // frecuencia_horas
  horas_ciclo = [
      (dt_inicio + timedelta(hours=i * frecuencia_horas))
      for i in range(dosis_diarias)
  ]

  t1_list, t2_list, t3_list = [], [], []
  for dt in horas_ciclo:
    h = dt.hour
    formato_h = f"{h:02d}H" if h != 0 else "24H"

    if 6 <= h < 14:
      t1_list.append(formato_h)
    elif 14 <= h < 22:
      t2_list.append(formato_h)
    else:
      t3_list.append(formato_h)

  t1_str = " - ".join(t1_list) if t1_list else ""
  t2_str = " - ".join(t2_list) if t2_list else ""
  t3_str = " - ".join(t3_list) if t3_list else ""

  return fecha_fin, t1_str, t2_str, t3_str


def evaluar_alerta_fin(fecha_fin_str):
  try:
    dt_fin = datetime.fromisoformat(fecha_fin_str)
    ahora = datetime.now()
    horas_restantes = (dt_fin - ahora).total_seconds() / 3600

    if horas_restantes < 0:
      return "⚠️ VENCIDO (Programar Retiro)"
    elif horas_restantes <= 48:
      return "🔔 PRÓXIMO A FIN (<=48h - Programar Retiro)"
    return "En curso"
  except Exception:
    return "En curso"


def estilo_filas(row):
  acceso = str(row.get("Tipo Acceso", "")).upper()
  alerta = str(row.get("Estado Tratamiento", ""))

  if acceso == "PICC":
    return [
        "background-color: #ffd6d6; color: #721c24; font-weight: bold;"
    ] * len(row)
  elif "Programar Retiro" in alerta:
    return [
        "background-color: #fff3cd; color: #856404; font-weight: bold;"
    ] * len(row)
  return [""] * len(row)


# ----------------- BARRA LATERAL (SIDEBAR) -----------------
with st.sidebar:
  st.markdown(
      f"<div style='text-align: center; padding: 10px;'>{SVG_LOGO_SURA}</div>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='text-align: center; font-size: 0.85rem; color: #00AEC7"
      " !important;'>SALUD EN CASA · GESTIÓN AGUDOS</p>",
      unsafe_allow_html=True,
  )
  st.markdown("---")
  menu = st.radio(
      "MENÚ OPERATIVO",
      [
          "📋 Gestión de Pacientes",
          "📊 Censo y Planilla en Vivo",
          "💾 Respaldo y Base Local",
      ],
  )
  st.markdown("---")
  st.caption("🔒 Acceso Seguro · SURA Colombia")

# ----------------- ENCABEZADO SUPERIOR -----------------
st.markdown(
    """
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Panel de Control Clínico y Cronograma de Dosis</h2>
            <p class="hospital-sub">Programa de Hospitalización Domiciliaria Agudos</p>
        </div>
        <div style="text-align: right;">
            <span style="background: #E8F4F8; color: #0033A0; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: bold;">
                📅 Ciclos T1 · T2 · T3
            </span>
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
        "👤 Información del Paciente", expanded=not bool(paciente_data)
    ):
      col1, col2, col3 = st.columns(3)

      if paciente_data:
        p_doc, p_nom, p_plan, p_prog, p_piso, p_zona, p_user = (
            paciente_data[0][0],
            paciente_data[0][1],
            paciente_data[0][2],
            paciente_data[0][3],
            paciente_data[0][4],
            paciente_data[0][5],
            paciente_data[0][6] if len(paciente_data[0]) > 6 else "",
        )
        nombre = col1.text_input("Nombre y Apellido", value=p_nom)
        plan = col2.selectbox(
            "Plan",
            ["POS", "Póliza", "ARL"],
            index=["POS", "Póliza", "ARL"].index(p_plan),
        )
        programa = col3.selectbox("Programa", ["Agudos"], index=0)

        col4, col5, col6 = st.columns(3)
        pisos = ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
        piso = col4.selectbox(
            "Piso",
            pisos,
            index=pisos.index(p_piso) if p_piso in pisos else 0,
        )

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
        zona = col5.selectbox(
            "Zona",
            zonas,
            index=zonas.index(p_zona) if p_zona in zonas else 0,
        )

        usuario_edita = col6.text_input(
            "Tu Nombre / Profesional responsable:", value=p_user
        )

        if st.button("Actualizar Ficha Paciente"):
          if usuario_edita:
            run_query(
                "UPDATE pacientes SET nombre=?, plan=?, programa=?, piso=?,"
                " zona=?, usuario_registro=? WHERE documento=?",
                (
                    nombre,
                    plan,
                    programa,
                    piso,
                    zona,
                    usuario_edita,
                    doc_busqueda,
                ),
                fetch=False,
            )
            st.success("Datos demográficos actualizados.")
          else:
            st.error("Debes ingresar tu nombre como responsable.")
      else:
        st.info("Paciente no encontrado. Digite los datos para admisión:")
        nombre = col1.text_input("Nombre y Apellido")
        plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"])
        programa = col3.selectbox("Programa", ["Agudos"])

        col4, col5, col6 = st.columns(3)
        piso = col4.selectbox(
            "Piso", ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"]
        )
        zona = col5.selectbox(
            "Zona",
            [
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
            ],
        )
        usuario_crea = col6.text_input("Tu Nombre / Profesional que registra:")

        if st.button("Ingresar Nuevo Paciente"):
          if nombre and usuario_crea:
            run_query(
                "INSERT INTO pacientes VALUES (?, ?, ?, ?, ?, ?, ?)",
                (
                    doc_busqueda,
                    nombre,
                    plan,
                    programa,
                    piso,
                    zona,
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
      st.subheader("💊 Formulación / Modificación de Tratamiento")

      tratamiento_activo = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC LIMIT 1",
          (doc_busqueda,),
      )

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

      with st.form("form_tratamiento_sura"):
        col_t1, col_t2, col_t3, col_t4 = st.columns([2, 1, 1, 1])
        t_nombre = col_t1.text_input(
            "Medicamento / Tratamiento",
            value=tratamiento_activo[0][2] if tratamiento_activo else "",
        )
        t_dosis = col_t2.text_input(
            "Dosis",
            value=tratamiento_activo[0][3] if tratamiento_activo else "",
        )
        t_via = col_t3.selectbox(
            "Vía", vias, index=vias.index(via_val)
        )
        t_acceso = col_t4.selectbox(
            "Acceso", accesos, index=accesos.index(acceso_val)
        )

        col_f1, col_f2 = st.columns(2)
        frecuencias = [4, 6, 8, 12, 24]
        frec_val = tratamiento_activo[0][6] if tratamiento_activo else 8
        frecuencia = col_f1.selectbox(
            "Frecuencia (Horas)",
            frecuencias,
            index=(
                frecuencias.index(frec_val) if frec_val in frecuencias else 2
            ),
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

        col_rep1, col_rep2 = st.columns([1, 2])
        ajuste_dosis = col_rep1.number_input(
            "Ajuste Dosis (+ Sumar / - Restar)",
            min_value=-30,
            max_value=30,
            value=tratamiento_activo[0][13] if tratamiento_activo else 0,
        )
        novedades = col_rep2.text_area(
            "NOVEDAD (PICC, Prórroga, Familiar Retira)",
            value=tratamiento_activo[0][14] if tratamiento_activo else "",
        )

        # Campo obligatorio de auditoría
        usuario_cambio = st.text_input(
            "✍️ Nombre del Profesional que realiza o autoriza el cambio"
            " (Obligatorio):",
            value="",
        )

        dt_inicio = datetime.combine(fecha_inicio, hora_inicio)
        fecha_fin_calculada, t1_calc, t2_calc, t3_calc = calcular_tratamiento(
            dt_inicio, dias, frecuencia, ajuste_dosis
        )

        st.markdown("#### 🕒 Horarios de Visitas Calculados")
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
        c_p1.metric("Turno T1 (06H-14H)", t1_calc if t1_calc else "—")
        c_p2.metric("Turno T2 (14H-22H)", t2_calc if t2_calc else "—")
        c_p3.metric("Turno T3 (22H-06H)", t3_calc if t3_calc else "—")
        c_p4.metric("Fin Estimado", fecha_fin_calculada.strftime("%d/%m %I:%M %p"))

        alerta_fin = evaluar_alerta_fin(fecha_fin_calculada.isoformat())
        if "Programar Retiro" in alerta_fin:
          st.warning(f"{alerta_fin}")

        btn_guardar = st.form_submit_button("💾 Guardar Tratamiento con Firma")

      if btn_guardar:
        if not usuario_cambio.strip():
          st.error(
              "❌ Debe registrar el nombre de la persona que realiza el cambio."
          )
        else:
          fecha_reg_str = datetime.now().isoformat()
          fecha_ini_str = dt_inicio.isoformat()
          fecha_fin_str = fecha_fin_calculada.isoformat()

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
                            dias, fecha_inicio, fecha_fin, t1, t2, t3, dosis_perdidas, novedades, usuario_modificacion, estado, fecha_registro
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?)
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
                  ajuste_dosis,
                  novedades,
                  usuario_cambio.strip(),
                  fecha_reg_str,
              ),
              fetch=False,
          )

          st.success(
              f"Tratamiento registrado correctamente por: {usuario_cambio}."
          )
          st.rerun()

      if tratamiento_activo:
        st.write("### 📌 Tratamiento Activo Actual")
        estado_alerta_act = evaluar_alerta_fin(tratamiento_activo[0][9])
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
            "Ajuste Dosis": tratamiento_activo[0][13],
            "Estado Tratamiento": estado_alerta_act,
            "Registrado Por": tratamiento_activo[0][15],
            "NOVEDAD": tratamiento_activo[0][14],
        }])
        st.dataframe(
            df_activo.style.apply(estilo_filas, axis=1), use_container_width=True
        )

      st.divider()
      st.subheader("📜 Trazabilidad de Auditoría (Historial Completo)")
      historial = run_query(
          "SELECT tratamiento, dosis, via_administracion, tipo_acceso,"
          " frecuencia_horas, dias, fecha_inicio, fecha_fin, t1, t2, t3,"
          " dosis_perdidas, usuario_modificacion, novedades, estado,"
          " fecha_registro FROM tratamientos WHERE documento = ? ORDER BY id"
          " DESC",
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
                "Ajuste Dosis",
                "Modificado Por",
                "NOVEDAD",
                "Estado",
                "Fecha Registro",
            ],
        )
        st.dataframe(df_hist, use_container_width=True)

# ----------------- SECCIÓN 2: BASE EN VIVO / PLANILLA -----------------
elif menu == "📊 Censo y Planilla en Vivo":
  query_general = """
        SELECT 
            p.documento AS 'Documento',
            p.nombre AS 'Nombre y Apellido',
            p.plan AS 'Plan',
            p.piso AS 'Piso',
            p.zona AS 'Zona',
            t.tratamiento AS 'Tratamiento',
            t.dosis AS 'Dosis',
            t.via_administracion AS 'Vía',
            t.tipo_acceso AS 'Tipo Acceso',
            t.frecuencia_horas AS 'Frec (h)',
            t.dias AS 'Días',
            t.fecha_inicio AS 'Fecha Inicio',
            t.fecha_fin AS 'Fecha Fin',
            t.t1 AS 'T1',
            t.t2 AS 'T2',
            t.t3 AS 'T3',
            t.dosis_perdidas AS 'Ajuste Dosis',
            t.usuario_modificacion AS 'Gestor Responsable',
            t.novedades AS 'NOVEDAD'
        FROM pacientes p
        LEFT JOIN tratamientos t ON p.documento = t.documento AND t.estado = 'Activo'
    """

  conn = sqlite3.connect(DB_FILE)
  df_general = pd.read_sql_query(query_general, conn)
  conn.close()

  if not df_general.empty:
    df_general["Estado Tratamiento"] = df_general["Fecha Fin"].apply(
        lambda f: evaluar_alerta_fin(f) if pd.notnull(f) else "Sin tratamiento"
    )

    # Tarjetas Métricas Superiores
    total_pacientes = len(df_general)
    total_picc = len(
        df_general[df_general["Tipo Acceso"].str.upper() == "PICC"]
    )
    total_retiros = len(
        df_general[df_general["Estado Tratamiento"].str.contains("Retiro")]
    )

    m1, m2, m3 = st.columns(3)
    m1.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_pacientes}</div><div"
        " class='metric-lbl'>CENSO ACTIVO</div></div>",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#721C24;'>{total_picc}</div><div"
        " class='metric-lbl'>PACIENTES CON PICC</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#856404;'>{total_retiros}</div><div"
        " class='metric-lbl'>PROGRAMAR RETIRO CATÉTER</div></div>",
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # Filtros
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    filtro_piso = col_f1.multiselect(
        "Piso", options=df_general["Piso"].dropna().unique()
    )
    filtro_zona = col_f2.multiselect(
        "Zona", options=df_general["Zona"].dropna().unique()
    )
    filtro_acceso = col_f3.multiselect(
        "Tipo Acceso", options=df_general["Tipo Acceso"].dropna().unique()
    )
    filtro_alerta = col_f4.multiselect(
        "Alertas Retiro",
        options=df_general["Estado Tratamiento"].dropna().unique(),
    )

    df_filtrado = df_general.copy()
    if filtro_piso:
      df_filtrado = df_filtrado[df_filtrado["Piso"].isin(filtro_piso)]
    if filtro_zona:
      df_filtrado = df_filtrado[df_filtrado["Zona"].isin(filtro_zona)]
    if filtro_acceso:
      df_filtrado = df_filtrado[df_filtrado["Tipo Acceso"].isin(filtro_acceso)]
    if filtro_alerta:
      df_filtrado = df_filtrado[
          df_filtrado["Estado Tratamiento"].isin(filtro_alerta)
      ]

    st.dataframe(
        df_filtrado.style.apply(estilo_filas, axis=1), use_container_width=True
    )

    # Exportación a Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_filtrado.to_excel(writer, index=False, sheet_name="Planilla_SURA")
      ws = writer.sheets["Planilla_SURA"]

      from openpyxl.styles import PatternFill

      fill_picc = PatternFill(
          start_color="FFD6D6", end_color="FFD6D6", fill_type="solid"
      )
      fill_alerta = PatternFill(
          start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"
      )

      col_acceso_idx = (
          df_filtrado.columns.get_loc("Tipo Acceso") + 1
          if "Tipo Acceso" in df_filtrado.columns
          else None
      )
      col_alerta_idx = (
          df_filtrado.columns.get_loc("Estado Tratamiento") + 1
          if "Estado Tratamiento" in df_filtrado.columns
          else None
      )

      for row_idx in range(2, len(df_filtrado) + 2):
        val_acceso = (
            ws.cell(row=row_idx, column=col_acceso_idx).value
            if col_acceso_idx
            else ""
        )
        val_alerta = (
            ws.cell(row=row_idx, column=col_alerta_idx).value
            if col_alerta_idx
            else ""
        )

        fill_aplicar = None
        if str(val_acceso).upper() == "PICC":
          fill_aplicar = fill_picc
        elif "Programar Retiro" in str(val_alerta):
          fill_aplicar = fill_alerta

        if fill_aplicar:
          for c_idx in range(1, len(df_filtrado.columns) + 1):
            ws.cell(row=row_idx, column=c_idx).fill = fill_aplicar

    excel_data = output.getvalue()
    st.download_button(
        label="📥 Descargar Planilla Completa (.xlsx)",
        data=excel_data,
        file_name=(
            f"Planilla_Agudos_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.info("No hay pacientes activos registrados.")

# ----------------- SECCIÓN 3: RESPALDO Y CONFIGURACIÓN LOCAL -----------------
elif menu == "💾 Respaldo y Base Local":
  st.subheader("💾 Gestión de Base de Datos y Servidor Local")
  st.markdown("""
    **Para asegurar que la base nunca se pierda:**
    1. Descarga periódicamente la copia de respaldo `.db`.
    2. Cuando estés en el PC de la sede que nunca se apaga, copia este proyecto allí y corre `streamlit run app.py`. Todo se almacenará en el disco duro físico del PC de forma permanente.
    """)

  c1, c2 = st.columns(2)
  with c1:
    st.markdown("#### 📤 Descargar Base Actual")
    if os.path.exists(DB_FILE):
      with open(DB_FILE, "rb") as f:
        bytes_db = f.read()
      st.download_button(
          label="⬇️ Descargar Archivo SQLite (.db)",
          data=bytes_db,
          file_name=(
              f"Backup_SURA_Agudos_{datetime.now().strftime('%Y%m%d_%H%M')}.db"
          ),
          mime="application/octet-stream",
      )
  with c2:
    st.markdown("#### 📥 Restaurar Base")
    subido = st.file_uploader("Subir archivo de respaldo (.db)", type=["db"])
    if subido and st.button("Restaurar"):
      with open(DB_FILE, "wb") as f:
        f.write(subido.getbuffer())
      st.success("✅ Base de datos cargada.")
      st.rerun()
