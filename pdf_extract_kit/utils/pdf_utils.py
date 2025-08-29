from pdf2image import convert_from_path
from img2pdf import convert as pdf_convert

def load_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    return images

def save_pdf(images, path: str):
    with open(path, 'wb') as f:
        f.write(pdf_convert(images))
