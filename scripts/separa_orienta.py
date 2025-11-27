from pypdf import PdfReader, PdfWriter
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import sys
import os
from pathlib import Path

def detect_orientation(image):
    """Detecta l’orientació mitjançant Tesseract (OSD)."""
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        return osd.get("rotate", 0)
    except pytesseract.TesseractError as e:
        print(f"⚠️  Error d’OSD: {e}. S’assumeix 0°.")
        return 0

def separa_i_orienta(input_pdf):
    """
    Corregeix l’orientació de cada pàgina i la guarda en un fitxer individual.
    Els fitxers tindran el nom: {output_prefix}_1.pdf, {output_prefix}_2.pdf, etc.
    """
    reader = PdfReader(input_pdf)
    doc = fitz.open(input_pdf)
    nomsenseext=input_pdf.replace('.pdf','')
    total_pages = len(reader.pages)
    print(f"📄 Processant {total_pages} pàgines...")

    for page_number in range(total_pages):
        pdf_page = reader.pages[page_number]
        fitz_page = doc[page_number]

        # Renderitzar com a imatge
        pix = fitz_page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Detectar i corregir orientació
        angle = detect_orientation(img)
        if angle != 0:
            print(f"↻ Pàgina {page_number + 1}: rotació detectada {angle}° → corregint")
            pdf_page.rotate(-angle)
        else:
            print(f"✓ Pàgina {page_number + 1}: orientació correcta")

        # Crear fitxer individual
        
        output_file = f"{nomsenseext}_Pag_{page_number + 1}.pdf"
        writer = PdfWriter()
        writer.add_page(pdf_page)
        with open(output_file, "wb") as f:
            writer.write(f)
        print(f"   → Guardat: {output_file}")

    doc.close()
    print(f"✅ Totes les pàgines separades amb prefix: Pag_*.pdf")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Ús: python orientation_fix.py input.pdf output_prefix")
        print("Exemple: python orientation_fix.py document.pdf plana")
        print("Generarà: plana_1.pdf, plana_2.pdf, ...")
        sys.exit(1)

    input_path = sys.argv[1]
    

    if not os.path.exists(input_path):
        print(f"❌ Error: No s’ha trobat el fitxer d’entrada: {input_path}")
        sys.exit(1)

    # Assegurar que el directori de sortida existeix (opcional)
    output_dir = Path(input_path).parent
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    separa_i_orienta(input_path)