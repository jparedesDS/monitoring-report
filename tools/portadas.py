from docx import Document

# Carga el documento
doc = Document("Attachment 13  Labels for Vendors Documents_Revised (002)__.docx")

# Busca y reemplaza los campos específicos
reemplazos = {
    "PURCHASE ORDER N°:": "12345",
    "MAT. REQ. Nº:": "67890",
    "ITEM N°:": "001",
    "WOOD DWG. N°:": "A-123"
}

for parrafo in doc.paragraphs:
    for clave, valor in reemplazos.items():
        if clave in parrafo.text:
            parrafo.text = parrafo.text.replace(clave, f"{clave} {valor}")

# Guarda el documento modificado
doc.save("documento_rellenado.docx")
