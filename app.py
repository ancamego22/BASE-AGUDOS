import io
import json
import os
import re
import sqlite3
from datetime import datetime, timedelta
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
DB_FILE = "sura_agudos_limpio.db"

# ----------------- ESTILOS CORPORATIVOS -----------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F4F7FB; color: #0F172A; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #002270 0%, #00123D 100%); color: white; padding-top: 1.5rem; }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stRadio label { background: rgba(255, 255, 255, 0.03); padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; display: block; border: 1px solid rgba(255, 255, 255, 0.06); font-weight: 500; font-size: 0.92rem; }
    .hospital-header { background: #FFFFFF; border-radius: 12px; padding: 1.2rem 2rem; box-shadow: 0px 4px 16px rgba(0, 51, 160, 0.04); display: flex; align-items: center; justify-content: space-between; margin-bottom: 1.5rem; border-left: 5px solid #0033A0; border: 1px solid #E2E8F0; }
    .hospital-title { font-size: 1.45rem; font-weight: 700; color: #0033A0; margin: 0; }
    .hospital-sub { font-size: 0.85rem; color: #00AEC7; margin: 0; font-weight: 600; text-transform: uppercase; }
    .footer-autor { text-align: center; padding: 1.4rem 0; margin-top: 3rem; border-top: 1px solid #E2E8F0; color: #64748B; font-size: 0.85rem; background: #FFFFFF; border-radius: 10px; }
    .footer-autor strong { color: #0033A0; }
    </style>
""",
    unsafe_allow_html=True,
)

# ----------------- CATÁLOGOS Y CONEXIÓN -----------------
def load_excel_catalog(filepath, fallback_val):
    if os.path.exists(filepath):
        try:
            df = pd.read_excel(filepath, sheet_name=0, header=None)
            if not df.empty:
                items = [str(val).strip() for val in df.iloc[:, 0].dropna() if str(val).strip().lower() not in ['nan', 'none', 'medicamento', 'gestor', 'nombre']]
                items = sorted(list(set(items)))
                if items:
                    return [""] + items
        except Exception:
            pass
    return ["", fallback_val]

cat_meds_agudos = load_excel_catalog("BASE MEDICAMENTOS AGUDOS.xlsx", "MEDICAMENTO GENÉRICO")
cat_gestores = load_excel_catalog("GESTORES DE AGUDOS.xlsx", "GESTOR GENÉRICO")

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
            frecuencia_horas INTEGER NOT NULL,
            dias INTEGER NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_fin TEXT NOT NULL,
            estado TEXT DEFAULT 'Activo',
            usuario_modificacion TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL,
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
            lugar_cita TEXT DEFAULT '',
            paciente_informado TEXT DEFAULT 'NO',
            nombre_acompanante TEXT DEFAULT '',
            contacto_acompanante TEXT DEFAULT '',
            resuelto_por TEXT DEFAULT ''
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

# ----------------- MENÚ LATERAL -----------------
with st.sidebar:
    st.markdown(f"""
        <div style="background-color: #FFFFFF; padding: 16px 20px; border-radius: 12px; margin-bottom: 18px; text-align: center;">
            <img src="{LOGO_SURA_URL}" style="width: 100%; max-width: 150px; height: auto;">
        </div>
    """, unsafe_allow_html=True)
    menu = st.radio("MENÚ OPERATIVO", [
        "Gestor Pacientes",
        "Ayudas diagnosticas",
        "Censo y Planilla en Vivo",
        "Informes"
    ])

# ----------------- ENCABEZADO -----------------
st.markdown(f"""
    <div class="hospital-header">
        <div>
            <h2 class="hospital-title">Base y Control Agudos</h2>
            <p class="hospital-sub">Programa Salud en Casa · SURA</p>
        </div>
        <div><img src="{LOGO_SURA_URL}" style="height: 42px;"></div>
    </div>
""", unsafe_allow_html=True)

# ----------------- 1. GESTOR PACIENTES -----------------
if menu == "Gestor Pacientes":
    st.subheader("🔍 Registro y Gestión de Pacientes")
    
    col_t, col_d, col_btn = st.columns([1, 2.5, 1])
    tipos_doc = ["CC", "TI", "CE", "PPT", "Pasaporte", "RC"]
    t_doc = col_t.selectbox("Tipo Doc", tipos_docel if 'tipos_docel' in locals() else tipos_doc)
    doc_ing = col_d.text_input("Número de Documento:").strip()

    if doc_ing:
        p_info = run_query("SELECT * FROM pacientes WHERE documento = ?", (doc_ing,))
        
        with st.form("form_paciente_limpio"):
            st.markdown("#### Ficha del Paciente")
            f_col1, f_col2, f_col3 = st.columns(3)
            nombre_p = f_col1.text_input("Nombre y Apellido", value=p_info[0][2] if p_info else "")
            plan_p = f_col2.selectbox("Plan", ["POS", "Póliza", "ARL"], index=0)
            piso_p = f_col3.selectbox("Piso", ["Norte", "Sur", "Larga Estancia", "AMI", "PAS"], index=0)
            
            f_col4, f_col5 = st.columns(2)
            zona_p = f_col4.selectbox("Zona", ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Zona 8"], index=0)
            gestor_p = st.selectbox("Profesional Responsable", cat_gestores)

            submitted = st.form_submit_button("💾 Guardar Paciente")
            if submitted:
                if not nombre_p or not gestor_p:
                    st.error("Por favor complete el nombre y el profesional responsable.")
                else:
                    run_query(
                        """INSERT OR REPLACE INTO pacientes (documento, tipo_documento, nombre, plan, programa, piso, zona, usuario_registro, fecha_ingreso)
                           VALUES (?, ?, ?, ?, 'Agudos', ?, ?, ?, ?)""",
                        (doc_ing, t_doc, nombre_p, plan_p, piso_p, zona_p, gestor_p, datetime.now().date().isoformat()),
                        fetch=False
                    )
                    st.success("✅ Paciente guardado exitosamente en la base limpia.")
                    st.rerun()

# ----------------- 2. AYUDAS DIAGNÓSTICAS -----------------
elif menu == "Ayudas diagnosticas":
    st.subheader("🧬 Solicitud y Gestión de Ayudas Diagnósticas")
    
    with st.form("form_ayuda_dx"):
        c1, c2 = st.columns([1, 2])
        t_doc_dx = c1.selectbox("Tipo Doc", ["CC", "TI", "CE", "PPT", "Pasaporte", "RC"])
        doc_dx = c2.text_input("Número de Documento del Paciente:").strip()
        
        ex_dx = st.text_input("Ayuda Diagnóstica / Examen Solicitado (Ej: TAC, Ecografía):")
        
        orig_dx = st.selectbox("Origen de Solicitud:", ["Línea de Ayudas Diagnósticas", "PGP"])
        prestador_dx = st.text_input("Prestador PGP (Si aplica):") if orig_dx == "PGP" else ""
        
        c3, c4, c5 = st.columns(3)
        peso_dx = c3.text_input("Peso (Kg):")
        o2_dx = c4.selectbox("Oxígeno", ["NO", "SI"])
        cant_o2 = c5.text_input("Flujo O2 (L/min):") if o2_dx == "SI" else ""
        
        trans_dx = st.selectbox("Transporte", ["No requiere transporte", "TAB", "TAM", "MVR"])
        gestor_dx = st.selectbox("Gestor Responsable", cat_gestores)
        
        submitted_dx = st.form_submit_button("💾 Guardar Solicitud de Ayuda Dx")
        if submitted_dx:
            if not doc_dx or not ex_dx or not gestor_dx:
                st.error("Complete el documento, la ayuda solicitada y el gestor.")
            else:
                run_query(
                    """INSERT INTO ayudas_diagnosticas (documento, tipo_documento, piso, zona, ayuda_solicitada, peso, requiere_oxigeno, cantidad_oxigeno, tipo_transporte, origen_solicitud, prestador_pgp, registrado_por, fecha_registro)
                       VALUES (?, ?, 'Norte', 'Zona 1', ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (doc_dx, t_doc_dx, ex_dx, peso_dx, o2_dx, cant_o2, trans_dx, orig_dx, prestador_dx, gestor_dx, datetime.now().strftime("%d/%m/%Y %I:%M %p")),
                    fetch=False
                )
                st.success("✅ Solicitud guardada exitosamente y formulario limpio.")
                st.rerun()

    st.markdown("---")
    st.markdown("#### 📋 Listado de Ayudas Registradas")
    df_ayudas = pd.read_sql_query("SELECT * FROM ayudas_diagnosticas ORDER BY id DESC", sqlite3.connect(DB_FILE))
    if not df_ayudas.empty:
        st.dataframe(df_ayudas, use_container_width=True)
    else:
        st.info("No hay ayudas diagnósticas registradas aún.")

# ----------------- 3. CENSO Y PLANILLA -----------------
elif menu == "Censo y Planilla en Vivo":
    st.subheader("📊 Censo de Pacientes Activos")
    df_pacs = pd.read_sql_query("SELECT * FROM pacientes", sqlite3.connect(DB_FILE))
    if not df_pacs.empty:
        st.dataframe(df_pacs, use_container_width=True)
    else:
        st.info("No hay pacientes registrados en el censo.")

# ----------------- 4. INFORMES -----------------
elif menu == "Informes":
    st.subheader("💾 Exportación de Respaldos")
    if st.button("📊 Descargar Base en Excel"):
        conn = sqlite3.connect(DB_FILE)
        df_exp = pd.read_sql_query("SELECT * FROM pacientes", conn)
        conn.close()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_exp.to_excel(writer, index=False, sheet_name="Pacientes")
        st.download_button("⬇️ Descargar Archivo", data=output.getvalue(), file_name="Reporte_SURA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------- PIE DE PÁGINA -----------------
st.markdown(f"""
    <div class="footer-autor">
        🏥 <strong>Sistema de Gestión y Control de Tratamientos Domiciliarios (SENC)</strong><br>
        Diseñado y desarrollado por <strong>Camilo Andrés Medina</strong> · Salud en Casa SURA Colombia
    </div>
""", unsafe_allow_html=True)
