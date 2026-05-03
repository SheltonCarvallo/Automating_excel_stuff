import pandas as pd
from openpyxl import load_workbook
import numpy as np

emails_en_match = []
emails_en_reporte_general = []



columnas_para_colocar_en_base_match = ['Sponsor', 'Estatus Participante','Universidad', 'Carrera', 'Nombre', 'CI', 'Celular', 'Correo', 'Empleador', 'Posicion Laboral']

path_emails_en_match = "emails_match.txt"
path_emails_en_reporte_general = "emails_plataforma.txt"

def find_correos_faltantes(emails_match: list[str], all_emails: list[str]) -> list[str]:
    emails_faltantes =  []
    try:
        print("Iniciando coparación de correos...")
        for email in all_emails:
            if email not in emails_match:
                emails_faltantes.append(email)       
    except Exception as e:
        print(f"El siguiente error se presento mientras se comparaba los emails: {e}")
    else:
        print("Comparación de correos ejecutada con exito.")
        print(f"Total de correo faltantes => {len(emails_faltantes)}")  
    return emails_faltantes

"""Lectura de los archivos planos"""

# Base match
try:
    print("Leyendo emails de la base de match...")
    with open(path_emails_en_match, "r", encoding="utf-8") as file:
        print(f"File descriptor => {file.fileno()}")
        for line in file:
            emails_en_match.append(line.strip().lower())       
except Exception as e:
    print(f"Ocurrió el siguiente error: {e}")
else:
    print(f"{len(emails_en_match)} correos de base match leídos")

# Reporte General
try:
    print("Leyendo emails del reporte general...")
    with open(path_emails_en_reporte_general, "r", encoding="utf-8") as file:
        print(f"File descriptor => {file.fileno()}")
        for line in file:
            emails_en_reporte_general.append(line.strip().lower())
except Exception as e:
    print(f"Ocurrió el siguiente error: {e}")
else:
    print(f"{len(emails_en_reporte_general)} correos de reporte general leídos")


emails_por_agregar_en_base_match = find_correos_faltantes(emails_match=emails_en_match, all_emails=emails_en_reporte_general)


df = pd.read_excel(io="ReporteGeneralSpring2027.xls", usecols=columnas_para_colocar_en_base_match, skiprows=5)
df['Correo'] = df['Correo'].str.strip().str.lower()

mask_email = df['Correo'].isin(emails_por_agregar_en_base_match)
mask_activos = df['Estatus Participante'] == 'Activo'

df_participantes_por_agregar = df.loc[(mask_activos & mask_email), columnas_para_colocar_en_base_match].sort_values(by='Sponsor')

with pd.ExcelWriter("Participantes_por_agregar.xlsx") as writer:
    df_participantes_por_agregar.to_excel(writer)
    print("Excel guardado")

wb = load_workbook("Participantes_por_agregar.xlsx")
ws = wb.active


for col in ws.columns:
   max_lenght = max(len(str(cell.value)) for cell in col if cell.value)
   ws.column_dimensions[col[0].column_letter].width = max_lenght + 2

wb.save('Participantes_por_agregar.xlsx')