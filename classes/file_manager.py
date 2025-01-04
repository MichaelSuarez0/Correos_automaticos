from dotenv import load_dotenv
import os
import pandas as pd
import re
import imaplib

script_dir = os.path.dirname(__file__)

# Cargar variables de entorno
load_dotenv()


# Variables globales
SUBJECT_FILTER = os.getenv("SUBJECT_FILTER")
DOWNLOAD_PATH = os.path.join(script_dir, "..", "descargas")  # Carpeta de descargas
UPLOAD_PATH = os.path.join(script_dir, "..", "descargas", "clasificados")  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = os.path.join(script_dir, "..", "email_templates") # Carpeta desde la que se obtendrán los email templates


class FileManager():
    def __init__(self, search_directory):
        """
        Inicializa el gestor de archivos.

        :param search_directory: Directorio donde se buscarán los archivos.
        :param target_directories: Diccionario con extensiones como claves y carpetas destino como valores.
        """
        self.search_directory = search_directory
        #self.target_directories = target_directories

    def list_files(self):
        file_paths_list = []
        file_names_list = []
        print(f'Archivos presentes en la carpeta {self.search_directory}:')
        for item in os.listdir(self.search_directory):
            item_full_path = os.path.join(self.search_directory, item)
            if os.path.isfile(item_full_path):
                file_paths_list.append([item, item_full_path])
                file_names_list.append(item)
                print(f" - {item}")

        #print(file_paths_list)
        if not file_paths_list:
            print(" - No se encontraron archivos en la carpeta")
        return file_names_list
        #return file_paths_list

    def rename_files(self, diccionario, lowercase = True):
        """
        Renombrar archivos en una carpeta a partir de valores de un diccionario
        
        Args:
            carpeta (str): Ruta de la carpeta donde están los archivos descargados.
            diccionario (dict): Diccionario para buscar los nombres en los keys y cambiarlos por sus values.
            lowercase (bool): Si es true, antes de renombrar según el dict, se convierte a minúsculas
        Returns: 
            renamed_files_map (list): List of renamed files
        """
        renamed_files_map = {}
         # Crear subcarpetas para clasificados y no clasificados
        carpeta_clasificados = os.path.join(self.search_directory, "clasificados")
        carpeta_no_clasificados = os.path.join(self.search_directory, "no_clasificados")
        os.makedirs(carpeta_clasificados, exist_ok=True)
        os.makedirs(carpeta_no_clasificados, exist_ok=True)

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
            patron = re.split(r'[\s-]', nombre_modificado)[0].strip()  # Se asume que el patrón está antes del primer espacio o del primero guion
            #print(f"Este es el código: {patron}")

            if patron in diccionario:
                # Si el código está en los datos, mover a 'clasificados'
                titulo_nuevo = diccionario[patron].get("titulo_largo", "sin_título").replace("/", "-").strip() # cambiar esta parte para que sea modular
                nuevo_nombre = f"{patron} - {titulo_nuevo}{extension}"
                nuevo_path = os.path.join(self.search_directory, "clasificados", nuevo_nombre)
                renamed_files_map[archivo] = nuevo_nombre
                clasificacion = "clasificados"
            else:
                # Si no se reconoce el código, mover a 'no_clasificados'
                nuevo_path = os.path.join(self.search_directory, "no_clasificados", archivo)
                clasificacion = "no_clasificados"
            
            # Verificar si el archivo ya existe en el destino y eliminarlo si es necesario
            if os.path.exists(nuevo_path):
                print(f"El archivo '{nuevo_path}' ya existe. Reemplazando...")
                os.remove(nuevo_path)  # Eliminar archivo existente

            # Mover el archivo al destino correspondiente
            os.rename(archivo_path, nuevo_path)
            print(f"Archivo '{archivo}' -> movido a: {clasificacion}")
        return renamed_files_map

    def rename_files2(self, format = "" , diccionario = {}):
        """
        Renombrar archivos según un formato específico
        """
        self.search_directory

