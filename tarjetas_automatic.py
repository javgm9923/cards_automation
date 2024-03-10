import PyPDF2
import io
import os

import pandas as pd

from os import getenv
from dotenv import load_dotenv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

#CONFIGURACIÓN DEL PROGRAMA
#Cargar variables de entorno con credenciales
load_dotenv()

# Definir la ruta al archivo Excel que contiene los datos a ser incorporados en el PDF y leer el archivo
archivo_excel = os.path.join(getenv("DOWNLOAD_DIRECTORY"), 'input', 'clientes.xlsx')
df = pd.read_excel(archivo_excel)

# Definir las coordenadas para cada valor del registro
coordenadas_base = [(182, 641), (182, 609), (182, 575), (182, 542), (182, 512), (380, 542), (380, 512)]
coordenadas_desplazadas = [(x, y - 343) for x, y in coordenadas_base]  # Desplazar las coordenadas para el segundo registro

# Función para añadir texto a un PDF en la misma página
def añadir_texto_a_pdf(pdf_path, output_path, registros, coordenadas1, coordenadas2):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    
    # Dibujar el primer registro usando las primeras coordenadas
    for texto, (x, y) in zip(registros[0], coordenadas1):
        c.drawString(x, y, str(texto))
    
    # Si hay un segundo registro, dibujarlo usando las segundas coordenadas
    if len(registros) > 1:
        for texto, (x, y) in zip(registros[1], coordenadas2):
            c.drawString(x, y, str(texto))
    
    c.save()
    
    packet.seek(0)
    new_pdf = PyPDF2.PdfReader(packet)
    existing_pdf = PyPDF2.PdfReader(open(pdf_path, "rb"))
    output = PyPDF2.PdfWriter()

    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    
    with open(output_path, "wb") as outputStream:
        output.write(outputStream)

# Asegúrate de que el directorio de salida existe
directorio_salida = os.path.join(getenv("DOWNLOAD_DIRECTORY"), 'output')
if not os.path.exists(directorio_salida):
    os.makedirs(directorio_salida)

# Iterar sobre cada par de filas del DataFrame
for indice in range(0, len(df), 2):
    registros_actuales = df.iloc[indice:indice+2].values
    
    nombre_archivo_salida = os.path.join(directorio_salida, f'tarjetas_{indice+1}-{indice+2}.pdf')
    pdf_base_path = os.path.join(getenv("DOWNLOAD_DIRECTORY"), 'input', 'tarjeta_base.pdf')
    añadir_texto_a_pdf(pdf_base_path, nombre_archivo_salida, registros_actuales, coordenadas_base, coordenadas_desplazadas)
