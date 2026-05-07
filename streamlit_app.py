import streamlit as st
import pandas as pd
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

st.set_page_config(page_title="Control Presupuesto", layout="centered")

# --- LOGO ---
try:
    st.image("logo.png", width=180)
except:
    st.title("💰 Control de Presupuesto")

# --- FORMULARIO ---
with st.form("form_presupuesto"):
    st.subheader("Registrar Nuevo Albarán")
    
    n_albaran = st.text_input("Número de Albarán")
    trabajador = st.text_input("Nombre del Trabajador")
    fecha = st.date_input("Fecha", datetime.now())
    
    partidas = [
        "Material Eléctrico", "Cuadros Eléctricos", "Iluminación", 
        "Mano de Obra", "Maquinaria", "Otros Gastos"
    ]
    partida_sel = st.selectbox("Partida asociada:", partidas)
    gastos = st.number_input("Importe del gasto (€)", min_value=0.0, step=0.01)
    comentarios = st.text_area("Comentarios")
    
    # NOTA EXTRA: Subir foto
    foto = st.file_uploader("📸 Foto del Albarán (opcional)", type=["jpg", "png", "jpeg"])
    
    boton = st.form_submit_button("Añadir Gasto")

# --- LÓGICA ---
if "datos_p" not in st.session_state:
    st.session_state.datos_p = []

if boton:
    if n_albaran and trabajador:
        nuevo_reg = {
            "Albarán": n_albaran,
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "Trabajador": trabajador,
            "Partida": partida_sel,
            "Gasto (€)": gastos,
            "Comentarios": comentarios
        }
        # Guardar foto si existe
        if foto:
            nuevo_reg["Foto_Bytes"] = foto.getvalue()
            nuevo_reg["Foto_Nombre"] = foto.name
        else:
            nuevo_reg["Foto_Bytes"] = None

        st.session_state.datos_p.append(nuevo_reg)
        st.success("Gasto registrado correctamente")
    else:
        st.error("Faltan datos obligatorios")

# --- TABLA Y ENVÍO ---
if st.session_state.datos_p:
    df = pd.DataFrame(st.session_state.datos_p).drop(columns=["Foto_Bytes"], errors="ignore")
    st.table(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()

    st.download_button("📥 Descargar Excel", data=excel_data, file_name="presupuesto.xlsx")

    if st.button("📧 Enviar por Correo a Ana"):
        st.info("Reporte preparado para ana@fundacionmasaveu.com. (Configura los 'Secrets' en Streamlit para el envío real).")
