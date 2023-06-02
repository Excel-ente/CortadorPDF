import PyPDF2
import io
from PIL import Image
import pytesseract

# Abrir el archivo PDF
pdf_file = open('C:/Users/kevin/OneDrive/Escritorio/pruebas_edenor.PDF', 'rb')

# Leer el archivo PDF con PyPDF2
pdf_reader = PyPDF2.PdfFileReader(pdf_file)

# Convertir la página del PDF en una imagen con PIL
page = pdf_reader.getPage(0)
resources = page['/Resources']
xobject = resources['/XObject'].getObject()

for obj in xobject:
    if xobject[obj]['/Subtype'] == '/Image':
        image = xobject[obj].getImage()

img = Image.open(io.BytesIO(image))

# Usar Tesseract OCR para extraer el texto de la imagen
texto = pytesseract.image_to_string(img)

# Procesar el texto extraído
print(texto)

# Cerrar el archivo PDF
pdf_file.close()
