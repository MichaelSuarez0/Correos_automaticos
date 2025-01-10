## 0. Variables
from correos_automaticos.classes.sharepoint_manager import Sharepoint
from correos_automaticos.classes.file_manager import FileManager
from dotenv import load_dotenv
import os
import re
import pandas as pd
from icecream import ic
import urllib.parse

script_dir = os.path.dirname(__file__) # for .py files
#script_dir = os.getcwd()  # for jupyter

# Cargar variables de entorno
load_dotenv()

# Variables globales
DOWNLOAD_PATH = os.path.join(script_dir, "..", "descargas")  # Carpeta de descargas
UPLOAD_PATH = os.path.join(script_dir, "..", "descargas", "clasificados")  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = os.path.join(script_dir, "..", "email_templates") # Carpeta desde la que se obtendrán los email templates

# Credenciales segundo usuario
SHAREPOINT_URL_SITE = os.getenv("SHAREPOINT_URL_SITE") # Ruta fija (Enlace)
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_FOLDER = os.getenv("SHAREPOINT_FOLDER") # Ruta del canal (execu compartidos/AOI Tendencias)
SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC") # Ruta específica del folder (Prueba)
SHAREPOINT_USERNAME = os.getenv("SHAREPOINT_USERNAME") # no usar os.path.join()



## 1. Iniciar variables
file_manager = FileManager(search_directory= UPLOAD_PATH)
session = Sharepoint()
session._auth()
excel_path = os.path.join(script_dir, "..", "docs", "Registro de Participación con adjuntos_v4.xlsx")
df_merged = pd.read_excel(excel_path)

meses = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}


## 2. Descargar los adjuntos del folder
file_list = file_manager.list_files()
#session.download_files_from_folder()


## 3. Definir función principal
def allocate_files_from_folder(data, file_name: str):
    code = ""
    if file_name in data["name"].values:
        row_index = data.index[data["name"] == file_name][0]

        # Actividad operativa
        AOI = data.loc[row_index, "Seleccione la actividad operativa o tema relacionado"]
        if AOI == "Asistencia técnica (Políticas y planes)":
            code = 'ATECNICA'
        elif AOI == "Espacios de difusión (Estudios/plataformas)":
            AOI = "Espacios de difusión (Estudios y plataformas)" # Para que no haya problema con los paths
            code = "DIFUSION"
        elif AOI == "Instrumentos técnicos en prospectiva":
            code = "INSTRUME"
        elif AOI == "Convenios":
            code = "CONSULTA"

        # Fecha 
        fecha = data.loc[row_index, "Fecha de ejecución de la actividad"]
        y,m,d = str(fecha.year), str(fecha.month).zfill(2), str(fecha.day).zfill(2)
        nombre_mes = meses[m]
        code = f'{code}-{y}-{m}-{d}'

        # Nivel de Gobierno
        nivel_gob = data.loc[row_index, "Nivel de Gobierno"]
        if nivel_gob == "Gobierno Nacional":
            code = f'{code}-GN'
        elif nivel_gob == "Gobierno Regional":
            code = f'{code}-GR'
        elif nivel_gob == "Gobierno Local":
            code = f'{code}-GL'
        else:
            code = f'{code}-NA'

        # Naturaleza del trabajo
        naturaleza = data.loc[row_index, "Naturaleza del trabajo"]
        if naturaleza == "Revisión de entregables":
            code = f'{code}-ENTREG'
        elif naturaleza in  ["Talleres", "Talleres de capacitación"]:
            code = f'{code}-TALLER'
        elif naturaleza == "Webinar":
            code = f'{code}-WEBINR'
        elif naturaleza == "Convenios":
            code = f'{code}-CONVEN'

        # Iniciales del autor
        autor = data.loc[row_index, "Especialista de la DNPE a cargo"]
        if autor in ["Enrique Del Águila", "Alberto Del Aguila"]:
            code = f'{code}-AA'
        else:
            autor = autor.split()
            inicial, segundo = autor[0], autor[1]
            code = f'{code}-{inicial[:1]}{segundo[:1]}'
        
        constructed_url = f'Documentos compartidos/AOI Asistencia técnica/Prueba/{AOI}/{nombre_mes}/{code}'
        #constructed_url = f'Documentos compartidos/AOI Tendencias/Prueba/{AOI}/{naturaleza}/{code}'
        #print(f' - URL: {constructed_url}, code: {code}')
        try:
            session.upload_file(file_name=file_name, custom_folder_path=constructed_url, create_folder=True)
        except Exception as e:
            print(f'Hubo un problema con la subida del archivo "{file_name}"')


def main():
    for file in file_list:
        allocate_files_from_folder(data=df_merged, file_name=file)

if __name__ == "__main__":
    main()


