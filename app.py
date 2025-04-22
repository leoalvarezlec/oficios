import streamlit as st
from docx import Document
from io import BytesIO

# --------------------------
# 1. Lista de destinatarios
# --------------------------
destinatarios = [
    {"nombre": "Juan Pérez", "cargo": "Director General"},
    {"nombre": "Laura Gómez", "cargo": "Jefa de Finanzas"},
    {"nombre": "Carlos Ruiz", "cargo": "Coordinador de Proyectos"}
]

st.title("Generador de Oficios Personalizados")

# --------------------------
# 2. Elegir destinatario
# --------------------------
nombres = [d["nombre"] for d in destinatarios]
nombre_seleccionado = st.selectbox("Selecciona un destinatario", nombres)

# Obtener los datos del destinatario elegido
destinatario = next(d for d in destinatarios if d["nombre"] == nombre_seleccionado)

# --------------------------
# 3. Secciones dinámicas
# --------------------------
if "textos" not in st.session_state:
    st.session_state.textos = []

if "tablas" not in st.session_state:
    st.session_state.tablas = []

if st.button("Agregar texto nuevo"):
    st.session_state.textos.append("")

if st.button("Agregar tabla nueva"):
    st.session_state.tablas.append([["", ""]])  # tabla básica de 2 columnas

# Mostrar campos de texto agregados
for i, texto in enumerate(st.session_state.textos):
    st.session_state.textos[i] = st.text_area(f"Texto #{i+1}", value=texto, height=100)

# Mostrar tablas agregadas
for i, tabla in enumerate(st.session_state.tablas):
    st.markdown(f"**Tabla #{i+1}**")
    cols = st.columns(2)
    fila1 = cols[0].text_input(f"Fila 1, Columna 1 - T{i}", value=tabla[0][0], key=f"t{i}_00")
    fila2 = cols[1].text_input(f"Fila 1, Columna 2 - T{i}", value=tabla[0][1], key=f"t{i}_01")
    st.session_state.tablas[i] = [[fila1, fila2]]

# --------------------------
# 4. Generar el oficio
# --------------------------
if st.button("Generar oficio"):
    doc = Document(plantilla_file.docx)

    # Encabezado básico
    doc.add_paragraph(f"Oficio para: {destinatario['nombre']}, {destinatario['cargo']}")

    # Agregar los textos
    for texto in st.session_state.textos:
        doc.add_paragraph(texto)

    # Agregar las tablas
    for tabla in st.session_state.tablas:
        table = doc.add_table(rows=1, cols=len(tabla[0]))
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(tabla[0]):
            hdr_cells[i].text = col

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("Oficio generado correctamente 🎉")
    st.download_button("Descargar oficio", buffer, file_name="oficio.docx")

