import rarfile
import pytesseract
from PIL import Image
import os
from gensim.summarization import summarize


rarfile.UNRAR_TOOL = 'C:\\Program Files\\WinRAR\\unrar.exe'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def extract_text_from_image(image_path):
    try:
        with Image.open(image_path) as img:
            text = pytesseract.image_to_string(img)
            return text.strip()
    except PermissionError as e:
        print(f"PermissionError: {e}. Skipping file: {image_path}")
        return ""


def generate_title(text):
    title = summarize(text, word_count=10)  
    return title

def extract_titles_from_images_in_rar(rar_file_path, folder_name):
    titles = []
    with rarfile.RarFile(rar_file_path, 'r') as rf:
        
        image_files = [rf.extract(file, path='temp') for file in rf.infolist() if os.path.dirname(file.filename) == folder_name]
        for image_file in image_files:
            text = extract_text_from_image(image_file)
            title = generate_title(text)
            titles.append(title)
    return titles

rar_file_path = 'scientific_publication.rar'
folder_name = 'scientific_publication'  # Specify the folder path relative to the RAR file
titles = extract_titles_from_images_in_rar(rar_file_path, folder_name)


excel_file_path = 'Ouput_format.xlsx'
wb = load_workbook(excel_file_path)
ws = wb.active


column = 'Generated Title'


for i, author in enumerate(authors, start=1):
    ws[f'{column}{i}'] = author

wb.save(excel_file_path)

print(f"Author names saved to {excel_file_path}")
