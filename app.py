import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from io import BytesIO
import os

def alinear_parrafo_derecha(parrafo):
    parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = parrafo._element.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '0')
    pPr.append(ind)

# -------------------------
# Ruta de imágenes del encabezado y pie de página
# -------------------------
ENCABEZADO_IMG = "encabezado.png"  # asegúrate que esté en la misma carpeta
PIE_IMG = "pie.png"

st.title("📝 Generador de Oficios")

# 1️⃣ Ingreso del número de oficio
numero_oficio = st.text_input("Número de oficio", placeholder="Ej. OF-123/2025")

# 2️⃣ Selección de destinatario
destinatarios = [
    {"nombre": "Juan Pérez", "cargo": "Director General"},
    {"nombre": "Laura Gómez", "cargo": "Jefa de Finanzas"},
    {"nombre": "Carlos Ruiz", "cargo": "Coordinador de Proyectos"}
]

nombres = [d["nombre"] for d in destinatarios]
nombre_seleccionado = st.selectbox("Selecciona un destinatario", nombres)
destinatario = next(d for d in destinatarios if d["nombre"] == nombre_seleccionado)

# 3️⃣ Inputs dinámicos
if "textos" not in st.session_state:
    st.session_state.textos = []
if "tablas" not in st.session_state:
    st.session_state.tablas = []

if st.button("Agregar texto"):
    st.session_state.textos.append("")

if st.button("Agregar tabla"):
    st.session_state.tablas.append([["", ""]])

# Campos de texto
for i, texto in enumerate(st.session_state.textos):
    st.session_state.textos[i] = st.text_area(f"Texto #{i+1}", value=texto)

# Tablas
for i, tabla in enumerate(st.session_state.tablas):
    st.markdown(f"**Tabla #{i+1}**")
    cols = st.columns(2)
    fila1 = cols[0].text_input(f"T{i}_0", value=tabla[0][0], key=f"t{i}_00")
    fila2 = cols[1].text_input(f"T{i}_1", value=tabla[0][1], key=f"t{i}_01")
    st.session_state.tablas[i] = [[fila1, fila2]]

# 4️⃣ Generar documento
if st.button("Generar oficio"):
    doc = Document()#"plantilla_file.docx")

    # Sección actual
    section = doc.sections[0]

    # Encabezado con imagen
    header_paragraph = header.add_paragraph()
    header_paragraph.add_run().add_picture(ENCABEZADO_IMG, width=Inches(7.5))
    alinear_imagen_a_la_derecha(header_paragraph)



    # Número de oficio en el encabezado, alineado a la derecha
    num_parrafo = header.add_paragraph()
    num_parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = num_parrafo.add_run(numero_oficio)
    run.bold = True
    run.font.size = Pt(12)

    # Pie de página con imagen
    footer_paragraph = footer.add_paragraph()
    footer_paragraph.add_run().add_picture(PIE_IMG, width=Inches(7.5))
    alinear_imagen_a_la_derecha(footer_paragraph)


    # Contenido del oficio
    doc.add_paragraph(f"Destinatario: {destinatario['nombre']}, {destinatario['cargo']}")

    for texto in st.session_state.textos:
        doc.add_paragraph(texto)

    for tabla in st.session_state.tablas:
        table = doc.add_table(rows=1, cols=len(tabla[0]))
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(tabla[0]):
            hdr_cells[i].text = col

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("Oficio generado correctamente 🎉")
    st.download_button("📥 Descargar oficio", buffer, file_name=f"Oficio_{numero_oficio}.docx")
