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
        font-size: 0.85rem;
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
            dosis_perdidas INTEGER DEFAULT 0,
            novedades TEXT,
            usuario_modificacion TEXT DEFAULT '',
            estado TEXT DEFAULT 'Activo',
            fecha_registro TEXT NOT NULL,
            fecha_retiro_cateter TEXT,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)

  # Migraciones preventivas
  c.execute("PRAGMA table_info(pacientes)")
  columnas_pacientes = [col[1] for col in c.fetchall()]
  if "usuario_registro" not in columnas_pacientes:
    c.execute("ALTER TABLE pacientes ADD COLUMN usuario_registro TEXT")

  c.execute("PRAGMA table_info(tratamientos)")
  columnas_trat = [col[1] for col in c.fetchall()]
  columnas_requeridas = {
      "tipo_acceso": "TEXT DEFAULT 'Periférico'",
      "t1": "TEXT",
      "t2": "TEXT",
      "t3": "TEXT",
      "usuario_modificacion": "TEXT DEFAULT ''",
      "dosis_perdidas": "INTEGER DEFAULT 0",
      "novedades": "TEXT",
      "fecha_retiro_cateter": "TEXT",
  }
  for col_name, col_def in columnas_requeridas.items():
    if col_name not in columnas_trat:
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


# ----------------- CÁLCULOS CLÍNICOS -----------------
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


def evaluar_alerta_fin(fecha_fin_str, estado="Activo"):
  if estado == "Completado":
    return "✅ FIN TTO / Retiro Realizado"

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
  estado = str(row.get("Estado", ""))
  acceso = str(row.get("Tipo Acceso", "")).upper()
  alerta = str(row.get("Estado Tratamiento", ""))

  if estado == "Completado":
    return ["color: #6C757D; font-style: italic;"] * len(row)

  if acceso == "PICC":
    return [
        "background-color: #ffd6d6; color: #721c24; font-weight: bold;"
    ] * len(row)
  elif "Programar Retiro" in alerta:
    return [
        "background-color: #fff3cd; color: #856404; font-weight: bold;"
    ] * len(row)
  return [""] * len(row)


# ----------------- BARRA LATERAL -----------------
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
      " !important; font-weight: 600;'>SALUD EN CASA · GESTIÓN AGUDOS</p>",
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
    f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Panel de Control Clínico y Cronograma de Dosis</h2>
            <p class="hospital-sub">Programa de Hospitalización Domiciliaria Agudos</p>
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
        "👤 Información del Paciente", expanded=not bool(paciente_data)
    ):
      col1, col2, col3 = st.columns(3)

      if paciente_data:
        p_doc = paciente_data[0][0]
        p_nom = paciente_data[0][1]
        p_plan = paciente_data[0][2]
        p_prog = paciente_data[0][3]
        p_piso = paciente_data[0][4]
        p_zona = paciente_data[0][5]
        p_user = paciente_data[0][6] if len(paciente_data[0]) > 6 else ""

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
            "Tu Nombre / Profesional responsable:",
            value=p_user if p_user else "",
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
        st.info("Paciente no registrado. Digite los datos para admisión:")
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

      # Buscar tratamiento activo vigente
      tratamiento_activo = run_query(
          "SELECT * FROM tratamientos WHERE documento = ? AND estado = 'Activo'"
          " ORDER BY id DESC LIMIT 1",
          (doc_busqueda,),
      )

      # ----------------- PANEL DE CIERRE / FIN TRATAMIENTO -----------------
      if tratamiento_activo:
        with st.container():
          st.markdown(
              "### 🏁 Cierre de Tratamiento y Retiro de Acceso Vascular"
          )
          st.info(
              "Si el paciente ya finalizó el esquema y enfermería retiró el"
              " catéter, márcalo aquí para cerrar el caso y retirarlo del censo"
              " de pendientes."
          )

          with st.expander(
              "✅ Registrar FIN DE TRATAMIENTO (Retiro de Catéter)"
          ):
            c_cierre1, c_cierre2 = st.columns(2)
            profesional_cierre = c_cierre1.text_input(
                "Profesional que confirma retiro:", key="prof_cierre"
            )
            fecha_cierre = c_cierre2.date_input(
                "Fecha de Retiro:", value=datetime.now().date()
            )
            nota_cierre = st.text_input(
                "Detalle de retiro / motivo de fin:",
                value="FIN TTO / Catéter retirado sin novedades",
                key="nota_cierre",
            )

            if st.button("Confirmar Fin de Tratamiento y Cierre"):
              if not profesional_cierre.strip():
                st.error("Debe ingresar el nombre del profesional que confirma.")
              else:
                fecha_retiro_str = (
                    f"{fecha_cierre.strftime('%d/%m/%Y')} por"
                    f" {profesional_cierre.strip()}"
                )
                nov_actual = (
                    tratamiento_activo[0][14]
                    if tratamiento_activo[0][14]
                    else ""
                )
                nueva_novedad = (
                    f"{nov_actual} // [CERRADO]: {nota_cierre} ("
                    f" {fecha_retiro_str})".strip()
                )

                run_query(
                    "UPDATE tratamientos SET estado = 'Completado', novedades ="
                    " ?, fecha_retiro_cateter = ? WHERE id = ?",
                    (nueva_novedad, fecha_retiro_str, tratamiento_activo[0][0]),
                    fetch=False,
                )
                st.success(
                    "✅ Tratamiento marcado como COMPLETADO. El catéter ya no"
                    " figura como pendiente."
                )
                st.rerun()

      # ----------------- FORMULARIO DE FORMULACIÓN / MODIFICACIÓN -----------------
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

        usuario_cambio = st.text_input(
            "✍️ Nombre del Profesional que realiza el cambio (Obligatorio):",
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
            "Registrado Por": (
                tratamiento_activo[0][15]
                if len(tratamiento_activo[0]) > 15
                else ""
            ),
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
  col_vista1, col_vista2 = st.columns([3, 1])
  filtro_estado = col_vista2.selectbox(
      "Filtrar por Estado:", ["Activos", "Completados (Fin TTO)", "Todos"]
  )

  condicion_estado = "WHERE t.estado = 'Activo'"
  if filtro_estado == "Completados (Fin TTO)":
    condicion_estado = "WHERE t.estado = 'Completado'"
  elif filtro_estado == "Todos":
    condicion_estado = "WHERE t.estado IN ('Activo', 'Completado')"

  query_general = f"""
        SELECT 
            p.documento AS 'Documento',
            p.nombre AS 'Nombre y Apellido',
            p.plan AS 'Plan',
            p.piso AS 'Piso',
            p.zona AS 'Zona',
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
            COALESCE(t.dosis_perdidas, 0) AS 'Ajuste Dosis',
            COALESCE(t.usuario_modificacion, '-') AS 'Gestor Responsable',
            COALESCE(t.novedades, '') AS 'NOVEDAD',
            COALESCE(t.estado, 'Sin TTO') AS 'Estado'
        FROM pacientes p
        JOIN tratamientos t ON p.documento = t.documento
        {condicion_estado}
    """

  conn = sqlite3.connect(DB_FILE)
  df_general = pd.read_sql_query(query_general, conn)
  conn.close()

  if not df_general.empty:
    df_general["Estado Tratamiento"] = df_general.apply(
        lambda row: evaluar_alerta_fin(row["Fecha Fin"], row["Estado"])
        if row["Fecha Fin"] != "-"
        else "Sin tratamiento",
        axis=1,
    )

    # Tarjetas métricas superiores (solo contabilizan tratamientos realmente activos)
    total_pacientes = len(df_general[df_general["Estado"] == "Activo"])
    total_picc = len(
        df_general[
            (df_general["Tipo Acceso"].astype(str).str.upper() == "PICC")
            & (df_general["Estado"] == "Activo")
        ]
    )
    total_retiros = len(
        df_general[
            df_general["Estado Tratamiento"].astype(str).str.contains("Retiro")
            & (df_general["Estado"] == "Activo")
        ]
    )

    m1, m2, m3 = st.columns(3)
    m1.markdown(
        f"<div class='metric-card'><div"
        f" class='metric-val'>{total_pacientes}</div><div class='metric-lbl'>CENSO"
        " ACTIVO</div></div>",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#721C24;'>{total_picc}</div><div"
        " class='metric-lbl'>PACIENTES PICC ACTIVOS</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-card'><div class='metric-val'"
        f" style='color:#856404;'>{total_retiros}</div><div"
        " class='metric-lbl'>PENDIENTES RETIRO CATÉTER</div></div>",
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

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
      col_estado_idx = (
          df_filtrado.columns.get_loc("Estado") + 1
          if "Estado" in df_filtrado.columns
          else None
      )

      for row_idx in range(2, len(df_filtrado) + 2):
        val_estado = (
            ws.cell(row=row_idx, column=col_estado_idx).value
            if col_estado_idx
            else ""
        )
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
        if val_estado == "Activo":
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
    st.info("No hay registros bajo el estado seleccionado.")

# ----------------- SECCIÓN 3: RESPALDO Y CONFIGURACIÓN LOCAL -----------------
elif menu == "💾 Respaldo y Base Local":
  st.subheader("💾 Gestión de Base de Datos y Servidor Local")
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
