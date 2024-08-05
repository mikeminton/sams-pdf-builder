import os
import shutil  # Import shutil module
import pikepdf
import openpyxl
from PIL import Image
from docx2pdf import convert
from PyPDF2 import PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Increase the decompression bomb size limit
Image.MAX_IMAGE_PIXELS = None  # This will disable the limit

def clear_directory(directory):
    # Create the directory if it doesn't exist
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"Directory {directory} created.")
    
    # Deletes all files in the given directory
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

def optimize_pdf(input_pdf, output_pdf):
    with pikepdf.open(input_pdf) as pdf:
        pdf.save(output_pdf, compress_streams=True)

def clear_terminal(): 
    # Check if the operating system is Windows
    if os.name == 'nt':
        os.system('cls')  # Command to clear terminal in Windows
    else:
        os.system('clear')  # Command to clear terminal in Linux/Mac
        
def convert_image_to_pdf(image_path): 
    img = Image.open(image_path)
    
    # Resize the image if it's too large
    max_size = (2480, 3508)  # A4 size in pixels at 300 dpi (approximately)
    img.thumbnail(max_size, Image.LANCZOS)  # Use Image.LANCZOS for high-quality resizing
    
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

def main(folder_path, output_folder, output_filename): 
    # Clear the directory if needed
    clear_directory(output_folder)

    # Create a PdfMerger object
    merger = PdfMerger()

    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    # Filter and sort the files alphanumerically
    files = sorted(files, key=lambda x: x.lower())

    # Keep track of optimized files to delete later
    optimized_files = []

    # Count the number of files to be processed
    num_files       = len(files)
    processed_files = 0

    print(f"Total number of files to process: {num_files}")

    for file in files: 
        file_path = os.path.join(folder_path, file)
        if file.lower().endswith('.pdf'): 
            # Optimize the PDF before appending
            optimized_pdf_path = file_path.rsplit('.', 1)[0] + '_optimized.pdf'
            optimize_pdf(file_path, optimized_pdf_path)
            merger.append(optimized_pdf_path)
            optimized_files.append(optimized_pdf_path)
        elif file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff')): 
            # If it's an image file, convert it to PDF and append
            pdf_path = convert_image_to_pdf(file_path)
            merger.append(pdf_path)
        elif file.lower().endswith('.docx'): 
            # If it's a DOCX file, convert it to PDF and append
            pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
            convert(file_path, pdf_path)
            merger.append(pdf_path)
        elif file.lower().endswith('.xlsx'): 
            # If it's an XLSX file, convert it to PDF and append
            pdf_path = convert_xlsx_to_pdf(file_path)
            merger.append(pdf_path)

        processed_files += 1
        print(f"Processed {processed_files}/{num_files} files...")

    # Write out the merged PDF after all files have been appended
    output_path = os.path.join(output_folder, output_filename)
    merger.write(output_path)
    merger.close()

    # Now that all files are merged, delete the optimized files
    for optimized_file in optimized_files:
        try:
            os.remove(optimized_file)
        except Exception as e:
            print(f"Failed to delete {optimized_file}: {e}")

    print("PDF building has been completed successfully!")

if __name__ == "__main__":
    clear_terminal()
    location_files_to_import   = r"C:\bot\pdf-builder\files-to-import"
    location_folder_to_export  = r"C:\bot\pdf-builder\output"
    filename_exported_file     = "pdf-builder-export.pdf"
    main(location_files_to_import, location_folder_to_export, filename_exported_file)
