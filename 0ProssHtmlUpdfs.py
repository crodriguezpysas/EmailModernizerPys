import os
import pdfkit
import fitz  # PyMuPDF

def html_a_pdf(input_path, output_path, path_wkhtmltopdf=None):
    """Convierte un archivo HTML a PDF usando pdfkit."""
    config = None
    if path_wkhtmltopdf:
        config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    
    options = {
        'no-stop-slow-scripts': '',
        'enable-local-file-access': '',
        'quiet': '',
        'dpi': '300'  # Establecer la resolución a 300 DPI
    }
    
    try:
        pdfkit.from_file(input_path, output_path, configuration=config, options=options)
    except IOError as e:
        print(f'Error al convertir {input_path} a {output_path}: {e}')

def convertir_todos_los_html(base_path, start, end, path_wkhtmltopdf=None):
    """Convierte todos los archivos HTML en el rango de subcarpetas especificado a PDF."""
    for i in range(start, end + 1):
        folder_path = os.path.join(base_path, str(i))
        input_file = os.path.join(folder_path, '0.html')
        output_file = os.path.join(folder_path, '0.pdf')

        if os.path.exists(input_file):
            print(f'Convirtiendo {input_file} a {output_file}...')
            html_a_pdf(input_file, output_file, path_wkhtmltopdf)
        else:
            print(f'{input_file} no existe.')

def combine_pdfs_in_folder(input_folder, output_file):
    """Combina todos los archivos PDF en la carpeta especificada en un solo archivo PDF."""
    output_pdf = fitz.open()
    for item in sorted(os.listdir(input_folder)):
        if item.endswith('.pdf'):
            with fitz.open(os.path.join(input_folder, item)) as pdf:
                for page_num in range(pdf.page_count):
                    output_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)
    output_pdf.save(output_file)
    output_pdf.close()

def main(base_path, start, end, path_wkhtmltopdf):
    """Convierte HTML a PDF y luego combina los PDFs generados en cada subcarpeta."""
    convertir_todos_los_html(base_path, start, end, path_wkhtmltopdf)
    
    for i in range(start, end + 1):
        subfolder = os.path.join(base_path, str(i))
        if os.path.exists(subfolder):
            output_file = os.path.join(base_path, f"PS{os.path.basename(base_path)}{str(i).zfill(4)}.pdf")
            combine_pdfs_in_folder(subfolder, output_file)
            print(f"Combined PDFs in folder {subfolder} into {output_file}")
        else:
            print(f"Subfolder {subfolder} does not exist, skipping.")

if __name__ == "__main__":
    base_path = 'C:/BCS/Final/250409'
    start_subfolder = 1
    end_subfolder = 331
    path_wkhtmltopdf = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
    main(base_path, start_subfolder, end_subfolder, path_wkhtmltopdf)
