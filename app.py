import io
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# ----------------- CONFIGURACIÓN DE PÁGINA -----------------
st.set_page_config(
    page_title="SURA - Gestión de Tratamientos Agudos",
    page_icon="🏥",
    layout="wide",
)

# ----------------- ESTILOS CORPORATIVOS SURA -----------------
st.markdown(
    """
    <style>
    /* Colores corporativos SURA */
    :root {
        --azul-sura: #0033A0;
        --aqua-sura: #00AEC7;
        --amarillo-sura: #FFE946;
        --gris-fondo: #F4F6F9;
    }
    
    /* Encabezado principal */
    .sura-header {
        background: linear-gradient(90deg, #0033A0 0%, #002270 100%);
        padding: 1.2rem 2rem;
        border-radius: 12px;
        color: white;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1.5rem;
        box-shadow: 0px 4px 10px rgba(0, 51, 160, 0.15);
    }
    .sura-title {
        font-size: 1.6rem;
        font-weight: 700;
        margin: 0;
        color: #FFFFFF;
    }
    .sura-subtitle {
        font-size: 0.95rem;
        color: #00AEC7;
        margin: 0;
        font-weight: 500;
    }
    
    /* Botones primarios en Azul SURA */
    div.stButton > button:first-child {
        background-color: #0033A0;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        transition: 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        background-color: #00AEC7;
        color: white;
    }

    /* Pestañas de Turnos T1, T2, T3 */
    button[data-baseweb="tab"] {
        font-weight: 600;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Logo oficial SURA en PNG transparente
LOGO_SURA_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/c/ca/Logo_sura.png/320px-Logo_sura.png"


# ----------------- BASE DE DATOS -----------------
def init_db():
  conn = sqlite3.connect("control_pacientes.db")
  c = conn.cursor()
  c.execute("""
        CREATE TABLE IF NOT EXISTS pacientes (
            documento TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            plan TEXT NOT NULL,
            programa TEXT NOT NULL,
            piso TEXT NOT NULL,
            zona TEXT NOT NULL
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
            estado TEXT DEFAULT 'Activo',
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (documento) REFERENCES pacientes (documento)
        )
    """)
  conn.commit()
  conn.close()


init_db()


def run_query(query, params=(), fetch=True):
  conn = sqlite3.connect("control_pacientes.db")
  c = conn.cursor()
  c.execute(query, params)
  data = c.fetchall() if fetch else None
  conn.commit()
  conn.close()
  return data


# ----------------- CÁLCULO DE FECHAS (SUMA O RESTA DE DOSIS) -----------------
def calcular_tratamiento(dt_inicio, dias, frecuencia_horas, ajuste_dosis=0):
  total_dosis_base = int((dias * 24) / frecuencia_horas)
  # Permite sumar (reponer) o restar (suspensión anticipada)
  total_dosis = max(1, total_dosis_base + ajuste_dosis)

  fecha_fin = dt_inicio + timedelta(hours=frecuencia_horas * (total_dosis - 1))

  # Horas dentro del ciclo de 24h
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


# ----------------- ESTILOS PARA LA PLANILLA (PICC) -----------------
def resaltar_picc(row):
  if str(row.get("Tipo Acceso", "")).upper() == "PICC":
    return [
        "background-color: #ffd6d6; color: #721c24; font-weight: bold;"
    ] * len(row)
  return [""] * len(row)


# ----------------- BANNER SUPERIOR SURA -----------------
st.markdown(
    f"""
    <div class="sura-header">
        <div>
            <h1 class="sura-title">Gestión y Control de Tratamientos</h1>
            <p class="sura-subtitle">Salud en Casa · Atención Domiciliaria Agudos</p>
        </div>
        <div>
            <img src="{LOGO_SURA_URL}" width="140" style="background: white; padding: 6px 12px; border-radius: 8px;">
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

# Barra lateral institucional
st.sidebar.image(LOGO_SURA_URL, width=150)
st.sidebar.markdown("### Navegación")
menu = st.sidebar.radio(
    "Seleccione el módulo:",
    ["Gestión de Pacientes", "Base en Tiempo Real (Planilla)"],
)

# ----------------- SECCIÓN 1: GESTIÓN DE PACIENTES -----------------
if menu == "Gestión de Pacientes":
  st.subheader("🔍 Buscar o Registrar Paciente")
  doc_busqueda = st.text_input("Ingrese Documento del Paciente:", "").strip()

  if doc_busqueda:
    paciente_data = run_query(
        "SELECT * FROM pacientes WHERE documento = ?", (doc_busqueda,)
    )

    with st.expander(
        "👤 Información General del Paciente", expanded=not bool(paciente_data)
    ):
      col1, col2, col3 = st.columns(3)

      if paciente_data:
        p_doc, p_nom, p_plan, p_prog, p_piso, p_zona = paciente_data[0]
        nombre = col1.text_input("Nombre y Apellido", value=p_nom)
        plan = col2.selectbox(
            "Plan",
            ["POS", "Póliza", "ARL"],
            index=["POS", "Póliza", "ARL"].index(p_plan),
        )
        programa = col3.selectbox("Programa", ["Agudos"], index=0)

        col4, col5 = st.columns(2)
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

        if st.button("Actualizar Datos Personales"):
          run_query(
              "UPDATE pacientes SET nombre=?, plan=?, programa=?, piso=?,"
              " zona=? WHERE documento=?",
              (nombre, plan, programa, piso, zona, doc_busqueda),
              fetch=False,
          )
          st.success("Datos personales actualizados correctamente.")
      else:
        st.info("Paciente nuevo. Diligencie los campos para crearlo:")
        nombre = col1.text_input("Nombre y Apellido")
        plan = col2.selectbox("Plan", ["POS", "Póliza", "ARL"])
        programa = col3.selectbox("Programa", ["Agudos"])

        col4, col5 = st.columns(2)
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

        if st.button("Guardar Paciente"):
          if nombre:
            run_query(
                "INSERT INTO pacientes VALUES (?, ?, ?, ?, ?, ?)",
                (doc_busqueda, nombre, plan, programa, piso, zona),
                fetch=False,
            )
            st.success("Paciente registrado con éxito.")
            st.rerun()
          else:
            st.error("Por favor complete el nombre y apellido.")

    if paciente_data:
      st.divider()
      st.subheader("💊 Formulación y Control de Tratamiento")

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

      with st.form("form_tratamiento"):
        col_t1, col_t2, col_t3, col_t4 = st.columns([2, 1, 1, 1])
        t_nombre = col_t1.text_input(
            "Tratamiento / Medicamento",
            value=tratamiento_activo[0][2] if tratamiento_activo else "",
        )
        t_dosis = col_t2.text_input(
            "Dosis",
            value=tratamiento_activo[0][3] if tratamiento_activo else "",
        )
        t_via = col_t3.selectbox(
            "Vía de Adm.", vias, index=vias.index(via_val)
        )
        t_acceso = col_t4.selectbox(
            "Tipo de Acceso", accesos, index=accesos.index(acceso_val)
        )

        col_f1, col_f2 = st.columns(2)
        frecuencias = [4, 6, 8, 12, 24]
        frec_val = tratamiento_activo[0][6] if tratamiento_activo else 8
        frecuencia = col_f1.selectbox(
            "Frecuencia (Cada cuántas horas)",
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
        # Permite números positivos (sumar) y negativos (restar)
        ajuste_dosis = col_rep1.number_input(
            "Ajuste de Dosis (+ Sumar / - Restar)",
            min_value=-30,
            max_value=30,
            value=tratamiento_activo[0][13] if tratamiento_activo else 0,
            help=(
                "Usa números positivos para reponer dosis perdidas (ej: 1, 2) o"
                " negativos si se suspenden dosis anticipadamente (ej: -1, -2)."
            ),
        )
        novedades = col_rep2.text_area(
            "NOVEDAD / Observaciones (ej: PICC, FAMILIAR RETIRA)",
            value=tratamiento_activo[0][14] if tratamiento_activo else "",
        )

        dt_inicio = datetime.combine(fecha_inicio, hora_inicio)
        fecha_fin_calculada, t1_calc, t2_calc, t3_calc = calcular_tratamiento(
            dt_inicio, dias, frecuencia, ajuste_dosis
        )

        st.write("#### 📋 Horarios de Visita Planificados")
        c_prev1, c_prev2, c_prev3, c_prev4 = st.columns(4)
        c_prev1.metric("T1 (06H - 14H)", t1_calc if t1_calc else "—")
        c_prev2.metric("T2 (14H - 22H)", t2_calc if t2_calc else "—")
        c_prev3.metric("T3 (22H - 06H)", t3_calc if t3_calc else "—")
        c_prev4.metric(
            "Fecha y Hora Fin", fecha_fin_calculada.strftime("%d/%m %I:%M %p")
        )

        if t_acceso == "PICC":
          st.warning(
              "⚠️ **ALERTA CLÍNICA:** Paciente con catéter PICC. Aparecerá"
              " resaltado en la planilla."
          )

        btn_guardar = st.form_submit_button(
            "💾 Guardar / Modificar Tratamiento"
        )

      if btn_guardar:
        fecha_registro_str = datetime.now().isoformat()
        fecha_inicio_str = dt_inicio.isoformat()
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
                        dias, fecha_inicio, fecha_fin, t1, t2, t3, dosis_perdidas, novedades, estado, fecha_registro
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activo', ?)
                """,
            (
                doc_busqueda,
                t_nombre,
                t_dosis,
                t_via,
                t_acceso,
                frecuencia,
                dias,
                fecha_inicio_str,
                fecha_fin_str,
                t1_calc,
                t2_calc,
                t3_calc,
                ajuste_dosis,
                novedades,
                fecha_registro_str,
            ),
            fetch=False,
        )

        st.success(
            "Tratamiento registrado. El registro anterior se archivó en el"
            " historial."
        )
        st.rerun()

      if tratamiento_activo:
        st.write("### 📌 Tratamiento Activo Actual")
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
            "NOVEDAD": tratamiento_activo[0][14],
        }])
        st.dataframe(
            df_activo.style.apply(resaltar_picc, axis=1),
            use_container_width=True,
        )

      st.divider()
      st.subheader("📜 Historial de Cambios Clínicos")
      historial = run_query(
          "SELECT tratamiento, dosis, via_administracion, tipo_acceso,"
          " frecuencia_horas, dias, fecha_inicio, fecha_fin, t1, t2, t3,"
          " dosis_perdidas, novedades, estado, fecha_registro FROM tratamientos"
          " WHERE documento = ? ORDER BY id DESC",
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
                "Frecuencia (h)",
                "Días",
                "Inicio",
                "Fin",
                "T1",
                "T2",
                "T3",
                "Ajuste Dosis",
                "NOVEDAD",
                "Estado",
                "Fecha Registro",
            ],
        )
        st.dataframe(
            df_hist.style.apply(resaltar_picc, axis=1), use_container_width=True
        )

# ----------------- SECCIÓN 2: BASE EN TIEMPO REAL (PLANILLA) -----------------
elif menu == "Base en Tiempo Real (Planilla)":
  st.subheader("📊 Censo y Planilla de Tratamientos en Tiempo Real")

  query_general = """
        SELECT 
            p.documento AS 'Documento',
            p.nombre AS 'Nombre y Apellido',
            p.plan AS 'Plan',
            p.programa AS 'Programa',
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
            t.novedades AS 'NOVEDAD'
        FROM pacientes p
        LEFT JOIN tratamientos t ON p.documento = t.documento AND t.estado = 'Activo'
    """

  conn = sqlite3.connect("control_pacientes.db")
  df_general = pd.read_sql_query(query_general, conn)
  conn.close()

  if not df_general.empty:
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    filtro_piso = col_f1.multiselect(
        "Filtrar por Piso", options=df_general["Piso"].dropna().unique()
    )
    filtro_zona = col_f2.multiselect(
        "Filtrar por Zona", options=df_general["Zona"].dropna().unique()
    )
    filtro_plan = col_f3.multiselect(
        "Filtrar por Plan", options=df_general["Plan"].dropna().unique()
    )
    filtro_acceso = col_f4.multiselect(
        "Filtrar por Acceso", options=df_general["Tipo Acceso"].dropna().unique()
    )

    df_filtrado = df_general.copy()
    if filtro_piso:
      df_filtrado = df_filtrado[df_filtrado["Piso"].isin(filtro_piso)]
    if filtro_zona:
      df_filtrado = df_filtrado[df_filtrado["Zona"].isin(filtro_zona)]
    if filtro_plan:
      df_filtrado = df_filtrado[df_filtrado["Plan"].isin(filtro_plan)]
    if filtro_acceso:
      df_filtrado = df_filtrado[df_filtrado["Tipo Acceso"].isin(filtro_acceso)]

    st.dataframe(
        df_filtrado.style.apply(resaltar_picc, axis=1), use_container_width=True
    )

    # Exportación a Excel con resaltado PICC
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_filtrado.to_excel(writer, index=False, sheet_name="Planilla_SURA")
      ws = writer.sheets["Planilla_SURA"]

      from openpyxl.styles import Font, PatternFill

      fill_picc = PatternFill(
          start_color="FFD6D6", end_color="FFD6D6", fill_type="solid"
      )
      font_picc = Font(color="721C24", bold=True)

      col_acceso_idx = None
      for idx, col in enumerate(df_filtrado.columns):
        if col == "Tipo Acceso":
          col_acceso_idx = idx + 1
          break

      if col_acceso_idx:
        for row_idx in range(2, len(df_filtrado) + 2):
          cell_val = ws.cell(row=row_idx, column=col_acceso_idx).value
          if str(cell_val).upper() == "PICC":
            for c_idx in range(1, len(df_filtrado.columns) + 1):
              cell = ws.cell(row=row_idx, column=c_idx)
              cell.fill = fill_picc
              if c_idx == col_acceso_idx:
                cell.font = font_picc

    excel_data = output.getvalue()
    nombre_archivo = (
        f"Planilla_Agudos_SURA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )

    st.download_button(
        label="📥 Descargar Planilla en Excel (.xlsx)",
        data=excel_data,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.info("No hay pacientes registrados aún en la base de datos.")
