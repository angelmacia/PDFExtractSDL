from pypdf import PdfReader, PdfWriter
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import os
from pathlib import Path

def detect_orientation(image):
    """Detecta orientació amb Tesseract (OSD = Orientation & Script Detection)."""
    osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
    angle = osd.get("rotate", 0)
    return angle


def correct_pdf_orientation(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    doc = fitz.open(input_pdf)

    for page_number, page in enumerate(reader.pages, start=1):
        # Renderitzar pàgina com imatge (per passar a Tesseract)
        pix = doc[page_number-1].get_pixmap(dpi=150)
        
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Detectar orientació
        angle = detect_orientation(img)

        if angle != 0:
            print(f"↻ Pàgina {page_number}: detectada rotació {angle}° -> corregint")
            page.rotate(angle)
            writer.add_page(page)
        else: 
            print(f"↻ Pàgina {page_number}: orientació correcta")
            #writer.add_page(page)

        if os.path.exists(output_pdf):
            output_pdf=output_pdf.replace('.pdf',"_"+str(page_number)+'.pdf')

        with open(output_pdf, "wb") as f:
            writer.write
       

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Ús: python orientation_fix.py input.pdf output.pdf")
    else:
        correct_pdf_orientation(sys.argv[1], sys.argv[2])
