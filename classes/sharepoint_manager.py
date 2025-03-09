from dotenv import load_dotenv
import os
import re
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from email.header import decode_header
from datetime import date, datetime, timedelta
import time
import logging

script_dir = os.path.dirname(__file__)

log_file_path = os.path.join(script_dir, "..", "scripts", "sharepoint_log.txt")

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

# Sharepoint credentials from env
SHAREPOINT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")

# Folders Sharepoint
SHAREPOINT_URL_SITE = "https://ceplangobpe.sharepoint.com/sites/DNPE"
SHAREPOINT_FOLDER = "Documentos compartidos/AOI Tendencias/Prueba"  # Ruta del folder de destino

# Folders Sharepoint (usuario personal)
# SHAREPOINT_URL_SITE = "https://ceplangobpe-my.sharepoint.com/personal/msuarez_ceplan_gob_pe" 
# SHAREPOINT_FOLDER = "Documents/Aplicaciones/Microsoft Forms/Registro de Asistencia Técnica Participación de la/Question" # Ruta del folder

# Para manejar diferentes funciones automáticamente
SHAREPOINT_ROOT_FOLDER = SHAREPOINT_URL_SITE.split("/")[-2] # sites 
SHAREPOINT_SITE_NAME = SHAREPOINT_URL_SITE.split("/")[-1] # DNPE
personal = True if SHAREPOINT_ROOT_FOLDER == "sites" else False


# Custom folders siempre deben comenzar con "Documentos compartidos" o su equivalente


class Sharepoint():
    def __init__(self, sharepoint_url: str, sharepoint_folder: str, connect_on_creation = True):
        self.SHAREPOINT_URL_SITE = sharepoint_url
        self.SHAREPOINT_FOLDER = sharepoint_folder
        self.SHAREPOINT_URL_BASE = "/".join(self.SHAREPOINT_URL_SITE.split("/")[:-2]) #.com
        self.SHAREPOINT_ROOT_FOLDER = self.SHAREPOINT_URL_SITE.split("/")[-2] # sites 
        self.SHAREPOINT_SITE_NAME = self.SHAREPOINT_URL_SITE.split("/")[-1] # DNPE
        self.conn = None
        if connect_on_creation:
            self.conn= self.auth()

    def auth(self):
        try:
            self.conn = AuthenticationContext(self.SHAREPOINT_URL_SITE)
            if self.conn.acquire_token_for_user(SHAREPOINT_EMAIL, SHAREPOINT_PASSWORD):
                self.conn = ClientContext(self.SHAREPOINT_URL_SITE, self.conn)
                print(f"Autenticación exitosa. Conexión establecida con SharePoint para {self.SHAREPOINT_URL_SITE}")
                return self.conn
        except Exception as e:
            print(f"Error al autenticar: {e}")
            return None
    
    def logout(self):
        """Log out from the SharePoint session."""
        if self.conn:
            try:
                # Simply clearing the connection object will log out the session
                self.conn = None
                print("Desconexión exitosa de SharePoint.")
            except Exception as e:
                print(f"Error al cerrar sesión: {e}")
        else:
            print("No hay conexión activa para cerrar sesión.")
    
    def _select_folder(self, custom_folder_path = "", folder_name= ""):
        if not custom_folder_path:
           target_folder_url = f'/{self.SHAREPOINT_ROOT_FOLDER}/{self.SHAREPOINT_SITE_NAME}/{self.SHAREPOINT_FOLDER}'
        else:
            target_folder_url = f'/{self.SHAREPOINT_ROOT_FOLDER}/{self.SHAREPOINT_SITE_NAME}/{custom_folder_path}'
        if folder_name:
            target_folder_url = f'{target_folder_url}/{folder_name}'
        return target_folder_url
    
    
    def list_files(self, custom_folder_path="", folder_name="", author = False):
        """
        List files in a specified folder and retrieve metadata, including author information if required (slows down retrieval)

        Args:
            custom_folder_path (str, optional): Custom folder path relative to the root. Default is an empty string.
            folder_name (str, optional): Specific folder name. Default is an empty string.
            author (bool, optional): Whether to include author information for each file. Default is False.

        Returns:
            list: A list of dictionaries containing metadata for each file, including:
                - name (str): File name.
                - server_relative_url (str): URL relative to the server.
                - time_created (str): Creation time in "YYYY-MM-DD HH:MM:SS" format.
                - time_last_modified (str): Last modification time in "YYYY-MM-DD HH:MM:SS" format.
                - author (str, optional): Author's email if `author=True`.
                - editor (str): Editor ID if available.
                - uniqueId (str): Unique identifier for the file.

        Raises:
            Exception: If the specified folder cannot be accessed.
        """
        file_metadata = []

        # Decide the base URL based on whether it's personal or a team site
        target_folder_url = self._select_folder(custom_folder_path, folder_name)

        # Extract folder name from the path
        folder_name = target_folder_url.split("/")[-1]

        try:
            # Get the folder by the server-relative URL
            root_folder = self.conn.web.get_folder_by_server_relative_url(target_folder_url)
            
            # Expand the folder to include files and subfolders
            root_folder.expand(["Files", "Folders"]).get().execute_query()  # No way to include author here
            print(f'Archivos presentes en la carpeta "{folder_name}":')
            print(f"Total de archivos encontrados: {len(root_folder.files)}")
        except Exception as e:
            print(f"No se encontró la carpeta con el path {target_folder_url}")
            return None

        # Iterate over the files and retrieve metadata
        for file in root_folder.files:
            # Access fields
            list_item = file.listItemAllFields
            editor = getattr(list_item, "EditorId", None)  # Editor ID or None if not present
            time_created = file.time_created  # Directly access if it's a datetime object
            time_modified = file.time_last_modified  # Directly access if it's a datetime object

            # Format times as strings
            time_created = time_created.strftime("%Y-%m-%d %H:%M:%S") if time_created else None
            time_modified = time_modified.strftime("%Y-%m-%d %H:%M:%S") if time_modified else None

            # Retrieve author information if requested
            author_data = None
            if author:
                try:
                    file.context.load(file, ["Author"])
                    file.context.execute_query()
                    author_data = file.author.email if hasattr(file.author, "email") else "Unknown"
                except Exception as e:
                    print(f"Error retrieving author information for file {file.name}: {e}")
            
            file_metadata.append({
                "name": file.name,
                "server_relative_url": file.serverRelativeUrl,
                "time_created": time_created,
                "time_last_modified": time_modified,
                "author": author_data,
                "editor": editor,
                "uniqueId": file.unique_id,
            })
            print(f' - {file.name}')  # Print only the name of each file

        return file_metadata
    
    def ensure_folders_exist(self, path: str):
        """
        Ensures all folders in the given path exist on SharePoint.

        Args:
            path (str): The folder path to create (e.g., "Parent/Child/Grandchild").
        
        Returns:
            str: "created" if any folder was created, "exists" if the entire path already existed.
        """
        try:
            folder_path = path.strip("/")  # Remove any leading/trailing slashes
            folder_names = folder_path.split("/")  # Split the path into individual folder names

            # Start at the root of the document library
            parent_folder = self.conn.web.folders.get_by_url(folder_names[0])
            self.conn.load(parent_folder)
            self.conn.execute_query()

            # Traverse and create folders incrementally
            for folder_name in folder_names[1:]:
                sub_folder = parent_folder.folders.add(folder_name)
                self.conn.execute_query()
                parent_folder = sub_folder  # Move to the newly created/ensured folder

            return "created"

        except Exception as e:
            # Handle the case where folders already exist
            if "already exists" in str(e).lower():
                return "exists"
            print(f"Failed to ensure folders exist for path '{path}'. Error: {e}")
            raise

    
    def upload_file(self, file_name: str, custom_folder_path="", create_folder = False):
        """
        Uploads a file to SharePoint. Optionally, a folder can be created with the same name as the file.
        Needs server-relative path (from /sites/)
        
        Args:
            file_name (str): Name of the file to upload, including the extension.
            target_folder_path (str, optional): SharePoint folder path to upload the file to 
                (e.g., "Tendencias/Tendencias Nacionales"). Default is the root folder from env.
            create_folder (bool, optional): If True, creates a folder with the same name as the file 
                (if it doesn't exist) before uploading. Default is False.
        
        Returns:
            bool: True if the upload was successful, otherwise raises an exception.
        """
        file_path = os.path.join(UPLOAD_PATH, file_name)
        with open(file_path, "rb") as file:
            content = file.read()  # Read binary content of the file       

    # Construir la URL del folder en SharePoint
        try:
            target_folder_url = self._select_folder(custom_folder_path)
        except Exception as e:
            print(f"ERROR al construir la URL del folder '{custom_folder_path}': {e}")
            return False

        # Crear carpeta si es necesario
        if create_folder and custom_folder_path:
            try:
                folder_status = self.ensure_folders_exist(custom_folder_path)
                #print(f"Carpeta '{custom_folder_path}' creada o ya existente")
            except Exception as e:
                print(f"ERROR al crear/verificar la carpeta '{custom_folder_path}': {e}")
                return False

        # Subir el archivo al folder de SharePoint
        try:
            target_folder = self.conn.web.get_folder_by_server_relative_path(target_folder_url)
            self.conn.load(target_folder)
            self.conn.execute_query()
        except Exception as e:
            print(f"ERROR al acceder a la carpeta de destino '{target_folder_url}': {e}")
            return False

        try:
            upload_status = target_folder.upload_file(file_name, content).execute_query()
            if upload_status:
                logging.info(f" - Archivo '{file_name}' subido exitosamente a '{self.SHAREPOINT_URL_BASE}{target_folder_url}'.")
                return True
            else:
                logging.error(f"ERROR desconocido al subir el archivo '{file_name}' a '{target_folder_url}'.")
                return False
        except Exception as e:
            logging.error(f"ERROR al subir el archivo '{file_name}' a la carpeta '{target_folder_url}': {e}")
            return False


    # def allocate_files_from_folder(self, dictionary, personal):
    #     """
    #     No tiene mucha utilidad por ahora

    #     Args:
    #         folder_name (str)

    #     Returns:
    #         uploaded_files (list)
    #     """
    #     uploaded_files = []
    #     local_folder_path = UPLOAD_PATH

    #     # Verificar que la carpeta local exista
    #     if not os.path.exists(local_folder_path):
    #         print(f"La carpeta local {local_folder_path} no existe.")
    #         return uploaded_files  # Si la carpeta no existe, retornar lista vacía
        
    #     try:
    #         for file_name in os.listdir(local_folder_path):
    #             file_path = os.path.join(local_folder_path, file_name)

    #             if os.path.isfile(file_path):
    #                 try:
    #                     upload_status = self.upload_file(file_name, )

    #                     if upload_status:
    #                         uploaded_files.append(file_name)
    #                 except Exception as e:
    #                     print(f"Error subiendo el archivo {file_name}: {e}")

    #     except Exception as e:
    #         print(f"Error al acceder a la carpeta local {local_folder_path}: {e}")

    #     print(f"Total de archivos subidos: {len(uploaded_files)}")
    #     return uploaded_files
        
    def download_file(self, file_url: str, file_name: str):
        """
        Descarga un archivo específico de SharePoint.

        Args:
            file_url (str): URL completa del archivo en SharePoint.
            file_name (str): Nombre del archivo a guardar localmente.

        Returns:
            str: Ruta local del archivo descargado.
        """
        try:
            # Descargar el archivo
            file = File.open_binary(self.conn, file_url)

            # Crear la ruta local para guardar el archivo
            local_file_path = os.path.join(DOWNLOAD_PATH, file_name)

            # Escribir el contenido en el archivo local
            with open(local_file_path, 'wb') as local_file:
                local_file.write(file.content)

            print(f"Archivo descargado con éxito: {local_file_path}")
            return local_file_path
        except Exception as e:
            print(f"No se pudo descargar el archivo {file_name}: {e}")
            return None
        
        
    def download_single_file(self, file_name, custom_folder_path = ""):
        """
        Falta modularizar
        """
        target_folder_url = self._select_folder(custom_folder_path)
        file_url = f'{target_folder_url}/{file_name}'

        try:
            local_file_path = self.download_file(file_url, file_name)
        except Exception as e:
            print(f"Error al descargar el archivo '{file_name}': {e}")
        return local_file_path
    
    def download_files_from_folder(self, custom_folder_path="", folder_name="", extension=""):
        """
        Descarga todos los archivos de una carpeta específica de SharePoint.
        
        Args:
            custom_folder_path (str, optional): Ruta personalizada de la carpeta.
            folder_name (str, optional): Nombre de la subcarpeta dentro de la carpeta personalizada.
            personal (bool, optional): Indica si la carpeta es personal o de un sitio de equipo.
            extension (str, optional): Filtra los archivos por extensión (e.g., '.txt', '.csv'). 
                Si está vacío, descarga todos los archivos.

        Returns:
            list: Lista de rutas locales de los archivos descargados.
        """
        # Crear la URL de la carpeta de destino
        target_folder_url = self._select_folder(custom_folder_path, folder_name)

        # Listar los archivos de la carpeta
        files_metadata = self.list_files(custom_folder_path, folder_name)
        if not files_metadata:
            print(f"No se encontraron archivos en la carpeta: {target_folder_url}")
            return []

        # Crear la carpeta local de descarga si no existe
        if not os.path.exists(DOWNLOAD_PATH):
            os.makedirs(DOWNLOAD_PATH)

        downloaded_files = []

        # Descargar cada archivo
        for file_meta in files_metadata:
            file_name = file_meta["name"]

            # Filtrar por extensión si se especifica
            if extension and not file_name.endswith(extension):
                continue

            # Construir la URL completa del archivo en SharePoint
            file_url = file_meta["server_relative_url"]

            try:
                # Descargar el archivo
                file = File.open_binary(self.conn, file_url)

                # Crear la ruta local para guardar el archivo
                local_file_path = os.path.join(DOWNLOAD_PATH, file_name)

                # Escribir el contenido en el archivo local
                with open(local_file_path, 'wb') as local_file:
                    local_file.write(file.content)

                print(f"Archivo descargado con éxito: {local_file_path}")
                downloaded_files.append(local_file_path)
            except Exception as e:
                print(f"No se pudo descargar el archivo {file_name}: {e}")
        print(f'Number of files downloaded: {len(downloaded_files)}')

        return downloaded_files


#sharepoint_session = Sharepoint(SHAREPOINT_URL_SITE, SHAREPOINT_FOLDER)
# sharepoint_session._auth()
# lista_archivos, _ = sharepoint_session.list_files(target_folder="Tendencias/Tendencias Globales")
# print(lista_archivos)
#Sharepoint().download_file(SHAREPOINT_DOC, "t90.docx")
#Sharepoint().upload_file('t75 - recuperación de la solidaridad.docx', "Tendencias/Tendencias Nacionales")


# gestionar subir archivos solo en modo lectura


    
