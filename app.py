import streamlit as st
import pandas as pd
from utils.generar_oficios import generar_oficios
from io import BytesIO
import zipfile

st.title("Generador de Oficios Automáticos")

# Subir Excel
excel_file = st.file_uploader("Sube el archivo Excel con los destinatarios", type=["xlsx"])

# Subir plantilla
plantilla_file = st.file_uploader("Sube la plantilla de Word", type=["docx"])

# Botón para generar
if st.button("Generar oficios") and excel_file and plantilla_file:
    df = pd.read_excel(excel_file)
    zip_buffer = generar_oficios(df, plantilla_file)
    
    st.success("Oficios generados correctamente 🎉")
    st.download_button("Descargar ZIP", zip_buffer, file_name="oficios.zip")
