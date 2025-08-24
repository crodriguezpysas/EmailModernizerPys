import os
import csv
import pdfkit
from bs4 import BeautifulSoup
import pandas as pd


def extract_email_data(html_path):
    """
    Extrae información específica de un archivo HTML.
    
    Args:
        html_path (str): Ruta del archivo HTML.
        
    Returns:
        dict: Información extraída.
    """
    try:
        with open(html_path, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')

            # Extraer información del contenido
            from_field = soup.find(text="De:").find_next().text.strip() if soup.find(text="De:") else "No encontrado"
            datetime_received = soup.find(text="Enviado el:").find_next().text.strip() if soup.find(text="Enviado el:") else "No encontrado"
            to_field = soup.find(text="Para:").find_next().text.strip() if soup.find(text="Para:") else "No encontrado"
            subject = soup.find(text="Asunto:").find_next().text.strip() if soup.find(text="Asunto:") else "No encontrado"
            attachments = soup.find(text="Datos Adjuntos:").find_next().text.strip() if soup.find(text="Datos Adjuntos:") else "No encontrado"
            
            # Extraer y ajustar campos específicos
            de_adjusted = attachments.split("Enviado el:")[0].strip() if "Enviado el:" in attachments else attachments
            enviado_el_adjusted = from_field.split(";")[0].strip()

            return {
                "From": from_field,
                "DateTimeReceived": datetime_received,
                "To": to_field,
                "Subject": subject,
                "Attachments": attachments,
                "De": de_adjusted,
                "Enviado el": enviado_el_adjusted,
            }
    except Exception as e:
        print(f"Error al procesar {html_path}: {e}")
        return None


def process_html_to_csv_and_pdf(base_path):
    """
    Procesa todos los archivos 0.html, extrae información, guarda en CSV y convierte a PDF.

    Args:
        base_path (str): Ruta base donde se encuentran las subcarpetas YYYYMMDD.
    """
    # Configuración de nombres
    summary_file = os.path.join(base_path, f"Summary{os.path.basename(base_path)}.csv")

    with open(summary_file, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow([
            "From", "DateTimeReceived", "To", "Subject", "Attachments", "Carpeta", 
            "De", "Enviado el"
        ])

        # Obtener carpetas y ordenar numéricamente
        subfolders = sorted([f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))], key=lambda x: int(x))

        for subfolder in subfolders:
            subfolder_path = os.path.join(base_path, subfolder)
            html_file_path = os.path.join(subfolder_path, "0.html")

            if not os.path.exists(html_file_path):
                print(f"Archivo 0.html no encontrado en {subfolder_path}")
                continue

            # Extraer datos del archivo HTML
            email_data = extract_email_data(html_file_path)
            if email_data:
                csv_writer.writerow([
                    email_data.get("De", ""),                  # From
                    email_data.get("From", ""),                # DateTimeReceived
                    email_data.get("DateTimeReceived", ""),    # To
                    email_data.get("To", ""),                  # Subject
                    email_data.get("Subject", ""),             # Attachments
                    subfolder,                                  # Carpeta
                    email_data.get("Attachments", ""),         # De
                    email_data.get("Enviado el", ""),          # Enviado el
                ])
                print(f"Procesado {html_file_path} exitosamente.")
            
            # Convertir a PDF
            pdf_file_path = os.path.join(subfolder_path, "0.pdf")
            try:
                pdfkit.from_file(html_file_path, pdf_file_path)
                print(f"Convertido a PDF: {pdf_file_path}")
            except Exception as e:
                print(f"Error al convertir {html_file_path} a PDF: {e}")


def guardar_resumen_a_excel(csv_path):
    """
    Convierte un archivo CSV de resumen en un archivo Excel.
    
    Args:
        csv_path (str): Ruta del archivo CSV a convertir.
    """
    excel_path = csv_path.replace(".csv", ".xlsx")
    try:
        df = pd.read_csv(csv_path, encoding='utf-8')
        df.to_excel(excel_path, index=False, sheet_name="Resumen")
        print(f"Resumen convertido a Excel: {excel_path}")
    except Exception as e:
        print(f"Error al convertir {csv_path} a Excel: {e}")


if __name__ == "__main__":
    base_path = r'C:\Emails\20250823'  # Cambia esta ruta según tu sistema
    process_html_to_csv_and_pdf(base_path)

    # Convertir el resumen CSV a Excel
    summary_csv = os.path.join(base_path, f"Summary{os.path.basename(base_path)}.csv")
    guardar_resumen_a_excel(summary_csv)