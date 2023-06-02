from PIL import Image
import fitz

def pdf_to_png(pdf_path, png_path):
    with fitz.open(pdf_path) as doc:
        page = doc.loadPage(0)  # carga la primera página del PDF
        pix = page.getPixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(png_path)

pdf_to_png('archivo.pdf', 'archivo.png')