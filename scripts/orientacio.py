from pypdf import PdfReader, PdfWriter
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import sys
import os

def detect_orientation(image):
    """Detecta l’orientació mitjançant Tesseract (OSD). Retorna l’angle de rotació (0, 90, 180, 270)."""
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        return osd.get("rotate", 0)
    except pytesseract.TesseractError as e:
        print(f"⚠️  Error d’OSD de Tesseract: {e}. S’assumeix orientació 0°.")
        return 0

def correct_pdf_orientation(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    doc = fitz.open(input_pdf)

    total_pages = len(reader.pages)
    print(f"📄 Processant {total_pages} pàgines...")

    for page_number in range(total_pages):
        pdf_page = reader.pages[page_number]
        fitz_page = doc[page_number]

        # Renderitzar la pàgina com a imatge
        pix = fitz_page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Detectar orientació
        angle = detect_orientation(img)

        if angle != 0:
            print(f"↻ Pàgina {page_number + 1}: detectada rotació {angle}° → corregint amb {-angle}°")
            pdf_page.rotate(-angle)  # Correcció: gir antihorari per desfer la rotació
        else:
            print(f"✓ Pàgina {page_number + 1}: orientació correcta")

        writer.add_page(pdf_page)  # Afegim sempre la pàgina (corregida o no)

    # Escrivim el PDF corregit UNA SOLA VEGADA
    with open(output_pdf, "wb") as f:
        writer.write(f)

    doc.close()
    print(f"✅ PDF corregit guardat a: {output_pdf}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Ús: python orientation_fix.py input.pdf output.pdf")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"❌ Error: No s’ha trobat el fitxer d’entrada: {input_path}")
        sys.exit(1)

    correct_pdf_orientation(input_path, output_path)