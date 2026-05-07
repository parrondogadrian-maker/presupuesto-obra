import streamlit as st
import pandas as pd
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Configuración de la página
st.set_page_config(page_title="Control Presupuestario", layout="centered")

# --- LOGO ---
try:
    st.image("logo.png", width=180)
except:
    st.title("💰 Control de Presupuesto")
    st.caption("Seguimiento de Albaranes y Gastos")

# --- FORMULARIO DE PRESUPUESTO ---
with st.form("form_presupuesto"):
    st.subheader("📝 Registro de Gasto (Albarán)")
    
    col1, col2 = st.columns(2)
    with col1:
        n_albaran = st.text_input("Número de Albarán")
        trabajador = st.text_input("Trabajador")
    with col2:
        fecha = st.date_input("Fecha", datetime.now())
        gastos = st.number_input("Importe (€)", min_value=0.0, step=0.01)
    
    partidas = [
        "Material Eléctrico", "Cuadros y Protecciones", "Iluminación", 
        "Canalizaciones", "Mano de Obra", "Maquinaria", "Varios"
    ]
    partida_asociada = st.selectbox("Partida del presupuesto:", partidas)
    comentarios = st.text_area("Comentarios / Notas")
    
    # NOTA EXTRA: Foto del albarán
    foto_albaran = st.file_uploader("📸 Foto del albarán (opcional)", type=["jpg", "jpeg", "png"])
    
    boton_guardar = st.form_submit_button("Guardar en Informe")

# --- GESTIÓN DE DATOS ---
if "datos_presupuesto" not in st.session_state:
    st.session_state.datos_presupuesto = []

if boton_guardar:
    if n_albaran and trabajador:
        registro = {
            "Albarán": n_albaran,
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "Trabajador": trabajador,
            "Partida": partida_asociada,
            "Gasto (€)": gastos,
            "Comentarios": comentarios
        }
        if foto_albaran:
            registro["Foto_Bytes"] = foto_albaran.getvalue()
            registro["Foto_Nombre"] = foto_albaran.name
        else:
            registro["Foto_Bytes"] = None
            
        st.session_state.datos_presupuesto.append(registro)
        st.success("✅ Gasto añadido al informe.")
    else:
        st.error("❌ Por favor, rellena el número de albarán y el trabajador.")

# --- SECCIÓN DE INFORME Y ENVÍO A ANA ---
if st.session_state.datos_presupuesto:
    st.write("---")
    st.subheader("📋 Informe de Gastos")
    
    # Mostrar tabla (sin los datos binarios de la foto)
    df_ver = pd.DataFrame(st.session_state.datos_presupuesto).drop(columns=["Foto_Bytes"], errors="ignore")
    st.table(df_ver)
    st.write(f"**Gasto Total acumulado:** {df_ver['Gasto (€)'].sum():.2f} €")

    # Preparar Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ver.to_excel(writer, index=False)
    excel_data = output.getvalue()

    st.write("### 📤 Finalizar y Enviar")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_data,
            file_name=f"gastos_obra_{datetime.now().strftime('%d_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_b:
        # BOTÓN DE ENVÍO A ANA
        if st.button("📧 Enviar Reporte a Ana"):
            destino = "ana@fundacionmasaveu.com"
            
            if "email_user" in st.secrets:
                try:
                    user = st.secrets["email_user"]
                    password = st.secrets["email_password"]
                    
                    msg = MIMEMultipart()
                    msg['From'] = user
                    msg['To'] = destino
                    msg['Subject'] = f"Presupuesto Obra - Albarán {n_albaran}"
                    
                    # Adjuntar Excel
                    part_ex = MIMEBase('application', "octet-stream")
                    part_ex.set_payload(excel_data)
                    encoders.encode_base64(part_ex)
                    part_ex.add_header('Content-Disposition', 'attachment; filename="gastos.xlsx"')
                    msg.attach(part_ex)
                    
                    # Adjuntar fotos si las hay
                    for r in st.session_state.datos_presupuesto:
                        if r["Foto_Bytes"]:
                            part_img = MIMEBase('image', 'png')
                            part_img.set_payload(r["Foto_Bytes"])
                            encoders.encode_base64(part_img)
                            part_img.add_header('Content-Disposition', f'attachment; filename="{r["Foto_Nombre"]}"')
                            msg.attach(part_img)
                    
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(user, password)
                    server.send_message(msg)
                    server.quit()
                    st.success(f"✅ ¡Enviado a {destino}!")
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.info(f"ℹ️ Reporte listo para enviar a {destino}. (Configura los 'Secrets' en Streamlit para el envío real).")

    st.warning("⚠️ Recuerda descargar o enviar antes de cerrar la pestaña.")
