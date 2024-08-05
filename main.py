import os
import shutil
import pikepdf
import openpyxl
from PIL import Image
from docx2pdf import convert
from PyPDF2 import PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import textwrap

# Increase the decompression bomb size limit
Image.MAX_IMAGE_PIXELS = None

def clear_directory(directory)   : 
    if  not os.path.exists(directory): 
        os.makedirs(directory)
        print(f"Directory {directory} created.")
    for filename in os.listdir(directory): 
        file_path = os.path.join(directory, filename)
        try                                                   : 
            if os.path.isfile(file_path) or os.path.islink(file_path): 
                os.unlink(file_path)
            elif os.path.isdir(file_path): 
                shutil.rmtree(file_path)
        except Exception as e: 
            print(f'Failed to delete {file_path}. Reason: {e}')

def  optimize_pdf(input_pdf, output_pdf): 
    with pikepdf.open(input_pdf) as pdf     : 
        pdf.save(output_pdf, compress_streams=True)

def clear_terminal(): 
    os.system('cls' if os.name == 'nt' else 'clear')

def convert_image_to_pdf(image_path): 
    img      = Image.open(image_path)
    max_size = (2480, 3508)
    img.thumbnail(max_size, Image.LANCZOS)
    pdf_path = image_path.rsplit('.', 1)[0] + '.pdf'
    img.convert('RGB').save(pdf_path)
    return pdf_path

def convert_xlsx_to_pdf(xlsx_path): 
    wb       = openpyxl.load_workbook(xlsx_path)
    sheet    = wb.active
    pdf_path = xlsx_path.rsplit('.', 1)[0] + '.pdf'
    c        = canvas.Canvas(pdf_path, pagesize=letter)
    width, height  = letter
    text     = c.beginText(40, height - 40)
    for row in sheet.iter_rows(values_only=True): 
        line = ', '.join([str(cell) for cell in row if cell is not None])
        text.textLine(line)
    c.drawText(text)
    c.save()
    return pdf_path

def  convert_text_to_pdf(text_path)                : 
    with open(text_path, 'r', encoding='utf-8') as file: 
        content      = file.read()
    pdf_path     = text_path.rsplit('.', 1)[0] + '.pdf'
    c            = canvas.Canvas(pdf_path, pagesize=letter)
    width, height      = letter
    margin       = 40
    text_width   = width - 2 * margin
    text         = c.beginText(margin, height - margin)
    wrapper      = textwrap.TextWrapper(width=80)  # Adjust width according to the font and page size
    wrapped_text = wrapper.wrap(content)
    for line in wrapped_text: 
        text.textLine(line)
    c.drawText(text)
    c.save()
    return pdf_path

def main(folder_path, output_folder, output_filename): 
    clear_directory(output_folder)
    merger               = PdfMerger()
    files                = sorted(os.listdir(folder_path), key=lambda x: x.lower())
    temp_files_to_delete = []
    num_files            = len(files)
    processed_files      = 0
    print(f"Total number of files to process: {num_files}")

    for file in files: 
        file_path = os.path.join(folder_path, file)
        if file.lower().endswith('.pdf'): 
            optimized_pdf_path = file_path.rsplit('.', 1)[0] + '_optimized.pdf'
            optimize_pdf(file_path, optimized_pdf_path)
            merger.append(optimized_pdf_path)
            temp_files_to_delete.append(optimized_pdf_path)
        elif file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff')): 
            pdf_path = convert_image_to_pdf(file_path)
            merger.append(pdf_path)
            temp_files_to_delete.append(pdf_path)
        elif file.lower().endswith('.docx'): 
            pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
            convert(file_path, pdf_path)
            merger.append(pdf_path)
            temp_files_to_delete.append(pdf_path)
        elif file.lower().endswith('.xlsx'): 
            pdf_path = convert_xlsx_to_pdf(file_path)
            merger.append(pdf_path)
            temp_files_to_delete.append(pdf_path)
        elif file.lower().endswith('.txt'): 
            pdf_path = convert_text_to_pdf(file_path)
            merger.append(pdf_path)
            temp_files_to_delete.append(pdf_path)

        processed_files += 1
        print(f"Processed {processed_files}/{num_files} files...")

    output_path = os.path.join(output_folder, output_filename)
    merger.write(output_path)
    merger.close()

    for temp_file in temp_files_to_delete: 
        try                              : 
            os.remove(temp_file)
            print(f"Deleted temporary file: {temp_file}")
        except Exception as e: 
            print(f"Failed to delete {temp_file}: {e}")

    print("PDF building has been completed successfully!")

if __name__ == "__main__":
    clear_terminal()
    location_files_to_import  = r"C:\bot\pdf-builder\files-to-import"
    location_folder_to_export = r"C:\bot\pdf-builder\output"
    filename_exported_file    = "pdf-builder-export.pdf"
    main(location_files_to_import, location_folder_to_export, filename_exported_file)
