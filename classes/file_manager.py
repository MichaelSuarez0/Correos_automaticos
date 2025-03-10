import enum
from dotenv import load_dotenv
import os
import pandas as pd
import re
from dataclasses import dataclass
from typing import List
import logging
from pprint import pprint

script_dir = os.path.dirname(__file__)

# Configuración del logging para guardar en el archivo con ruta personalizada
log_file_path = os.path.join(script_dir, "..", "scripts", "file_manager_log.txt")

logging.basicConfig(
    level=logging.INFO,  # Nivel de registro
    format="%(asctime)s - %(levelname)s - %(message)s",  # Formato del mensaje
    datefmt="%Y-%m-%d %H:%M:%S",  # Formato de fecha y hora
    handlers=[
        logging.FileHandler(log_file_path),  # Guardar en la ruta especificada
        logging.StreamHandler()  # También mostrar en la consola
    ]
)

# Cargar variables de entorno
load_dotenv()

# Variables globales
DOWNLOAD_PATH = os.path.join(script_dir, "..", "descargas")  # Carpeta de descargas
UPLOAD_PATH = os.path.join(script_dir, "..", "descargas", "clasificados")  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = os.path.join(script_dir, "..", "email_templates") # Carpeta desde la que se obtendrán los email templates

# TODO: Add function to copy files with a specific extension
class FileManager():
    def __init__(self, search_directory):
        """
        Initializes File Manager for managing files in the specified directory.

        Args:
            search_directory (str): Directory to search for files. Should be a valid path string.
        """
        self.search_directory = search_directory
        #self.target_directories = target_directories

    def list_files(self, extension: str = None, with_extension: bool = True) -> list:
        """
        Lista los nombres de archivos presentes en el directorio de búsqueda, 
        permitiendo filtrar por extensión y seleccionar si se devuelve el nombre
        con o sin la extensión.

        Args:
            extension (str, optional): Filtro de extensión (por ejemplo, '.txt'). 
                Si se deja en None, se listan todos los archivos.
            with_extension (bool, optional): Si es True, devuelve los nombres con extensión;
                si es False, devuelve los nombres sin la extensión. Defaults to True.

        Returns:
            list: Lista de nombres de archivos filtrados y formateados según los parámetros.
        """
        file_paths_list = []
        file_names_list = []
        #print(f'Archivos presentes en la carpeta {self.search_directory}:')
        for item in os.listdir(self.search_directory):
            item_full_path = os.path.join(self.search_directory, item)
            if os.path.isfile(item_full_path):
                file_paths_list.append([item, item_full_path])
                file_names_list.append(item)
                #print(f" - {item}")

        #print(file_paths_list)
        if not file_paths_list:
            logging.info(" - No se encontraron archivos en la carpeta")
            return None

        if extension:
            file_names_list = [file for file in file_names_list if extension == os.path.splitext(file)[1]]
        
        if not with_extension:
            file_names_list = [os.path.splitext(file)[0] for file in file_names_list]
        return file_names_list


    @staticmethod
    def sort_files_by_number(file_list: list) -> list:
        def extract_number(text: str) -> int:   
            # This finds all numbers in the string and returns the first occurrence.
            numbers = re.findall(r'\d+', text)
            return int(numbers[0]) if numbers else float('inf')
        return sorted(file_list, key=extract_number)
    

    def rename_files(self, diccionario: dict, lowercase = True) -> list[dict]:
        """
        Renombrar archivos en una carpeta a partir de valores de un diccionario
        
        Args:
            carpeta (str): Ruta de la carpeta donde están los archivos descargados.
            diccionario (dict): Diccionario para buscar los nombres en los keys y cambiarlos por sus values.
            lowercase (bool): Si es true, antes de renombrar según el dict, se convierte a minúsculas
        Returns: 
            renamed_files_map (list): List of dicts containing new_name and original_name as keys for every file.
        """
        renamed_files_map = []

         # Seleccionar subcarpetas para clasificados y no clasificados
        carpeta_clasificados = os.path.join(self.search_directory, "clasificados")
        carpeta_no_clasificados = os.path.join(self.search_directory, "no_clasificados")
        os.makedirs(carpeta_clasificados, exist_ok=True)
        os.makedirs(carpeta_no_clasificados, exist_ok=True)

        # Iterar sobre los archivos
        for archivo in os.listdir(self.search_directory):
            archivo_path = os.path.join(self.search_directory, archivo)

            # Asegurar que es un archivo y no una carpeta
            if not os.path.isfile(archivo_path):
                continue

            # Extraer patron del nombre del archivo (si está presente)
            nombre_original, extension = os.path.splitext(archivo)
            if lowercase:
                nombre_modificado = nombre_original.lower()
            else:
                nombre_modificado = nombre_original
            patron = re.split(r'[\s-]', nombre_modificado)[0].strip()  # Se asume que el patrón está antes del primer espacio o guion
            #print(f"Este es el código: {patron}")

            if patron in diccionario:
                # Si el código está en los datos, mover a 'clasificados'
                titulo_nuevo = diccionario[patron].get("titulo_largo", "sin_título").replace("/", "-").strip() # cambiar esta parte para que sea modular
                nuevo_nombre = f"{patron} - {titulo_nuevo}{extension}"
                nuevo_path = os.path.join(self.search_directory, "clasificados", nuevo_nombre)
                clasificacion = "clasificados"
            else:
                # Si no se reconoce el código, mover a 'no_clasificados'
                nuevo_path = os.path.join(self.search_directory, "no_clasificados", archivo)
                clasificacion = "no_clasificados"
            
            # Verificar si el archivo ya existe en el destino y eliminarlo si es necesario
            if os.path.exists(nuevo_path):
                logging.info(f"El archivo '{archivo}' ya existe. Reemplazando...")
                os.remove(nuevo_path)  # Eliminar archivo existente

            # Mover el archivo al destino correspondiente
            os.rename(archivo_path, nuevo_path)
            logging.debug(f" Archivo '{archivo}' -> movido a: {clasificacion}")

            # Agregar el mapeo de archivos renombrados
            try:
                renamed_files_map.append({
                    "original_name" : archivo,
                    "new_name" : nuevo_nombre
                })
            except UnboundLocalError:
                logging.error(f" Código del archivo '{archivo} no se encontró en info_obs'")

        return renamed_files_map

    # def rename_files2(self, format = "" , diccionario = {}):
    #     """
    #     Renombrar archivos según un formato específico
    #     """
    #     self.search_directory

# @dataclass
# class EmailMetadata:
#     id: str
#     _from: str
#     to: str
#     subject: str
#     body: str
#     attachments: List[str]

# @dataclass
# class EmailList:
#     email_list: List[EmailMetadata]

#     def convert_to_dataframe(self):
#     # Convierte el diccionario de datos de emails en un DataFrame
#         return pd.DataFrame.from_records(self.email_list)
    
#     def filter_by_sender(self):
#         return None
#         #return [email from email in self.email_list]
        

# {'from_name': 'Gabriela Sthefany Pozo Bornas',
#   'from_email': 'gpozo@ceplan.gob.pe',
#   'sent': '',
#   'to': 'Consulta Técnica <consultatecnica@ceplan.gob.pe>',
#   'subject': 'Sistematizar Tg39',
#   'body': '',
#   'attachments': ['Tg39-Variabilidad de las precipitaciones 13.12.docx',
#    'Tg39-Anexo 7B Variabilidad de las precipitaciones 13.12.xlsx']},