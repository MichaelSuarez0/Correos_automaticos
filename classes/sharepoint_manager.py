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


script_dir = os.path.dirname(__file__)

# Cargar variables de entorno
load_dotenv()

# Variables globales
SUBJECT_FILTER = os.getenv("SUBJECT_FILTER")
DOWNLOAD_PATH = os.path.join(script_dir, "..", "descargas")  # Carpeta de descargas
UPLOAD_PATH = os.path.join(script_dir, "..", "descargas", "clasificados")  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = os.path.join(script_dir, "..", "email_templates") # Carpeta desde la que se obtendrán los email templates

# Credenciales Sharepoint
SHAREPOINT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
SHAREPOINT_URL_SITE = os.getenv("SHAREPOINT_URL_SITE") # Ruta fija (Enlace)
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME") # Nombre de la ruta fija (DNPE)
SHAREPOINT_FOLDER = os.getenv("SHAREPOINT_FOLDER") # Ruta del folder (Documentos compartidos/AOI Tendencias/Prueba)
SHAREPOINT_USERNAME = os.getenv("SHAREPOINT_USERNAME") # Nombre del usuario (msuarez_ceplan_gob_pe)


class Sharepoint():
    def __init__(self):
        self.conn= None

    def _auth(self):
        try:
            self.conn = AuthenticationContext(SHAREPOINT_URL_SITE)
            if self.conn.acquire_token_for_user(SHAREPOINT_EMAIL, SHAREPOINT_PASSWORD):
                self.conn = ClientContext(SHAREPOINT_URL_SITE, self.conn)
                print("Autenticación exitosa. Conexión establecida con SharePoint")
        except Exception as e:
            print(f"Error al autenticar: {e}")
            return None
    
    @staticmethod
    def _select_folder(custom_folder_path = "", folder_name= "", personal = False):
        # Decide the base URL based on whether it's personal or a team site
        if not custom_folder_path:
            if personal:
                # If 'personal' is True, the URL will start with '/personal/'
                target_folder_url = f'/personal/{SHAREPOINT_USERNAME}/{SHAREPOINT_FOLDER}'
            else:
                # If 'persona' is False, the URL will start with '/sites/'
                target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_FOLDER}'
        else:
            if personal:
                target_folder_url = f'/personal/{custom_folder_path}'
            else:
                target_folder_url = f'/sites/{custom_folder_path}'
        if folder_name:
            target_folder_url = f'{target_folder_url}/{folder_name}'
        return target_folder_url
    
    
    def list_files(self, custom_folder_path="", folder_name="", personal=False):
        """_summary_

        Args:
            target_folder (str, optional): _description_. Defaults to "".
            personal (bool, optional): _description_. Defaults to False.

        Returns:
            _type_: _description_
        """
        file_metadata = []

        # Decide the base URL based on whether it's personal or a team site
        target_folder_url = self._select_folder(custom_folder_path, folder_name, personal)

        # Extract folder name from the path
        folder_name = target_folder_url.split("/")[-1]

        try:
            # Get the folder by the server-relative URL
            root_folder = self.conn.web.get_folder_by_server_relative_url(target_folder_url)
            
            # Expand the folder to include files and subfolders (this is why we use 'expand')
            root_folder.expand(["Files", "Folders"]).get().execute_query()  # No way to include author here
            print(f'Archivos presentes en la carpeta "{folder_name}":')
            print(f"Total de archivos encontrados: {len(root_folder.files)}")
        except Exception as e:
            print(f"No se encontró la carpeta con el path {target_folder_url}")
            return None

        # Iterate over the files and retrieve metadata
        for file in root_folder.files:
            # Ensure ListItemAllFields is loaded
            list_item = file.listItemAllFields

            # Safely access fields
            editor = getattr(list_item, "EditorId", None)  # Editor ID or None if not present
            time_created = file.time_created  # Directly access if it's a datetime object
            time_modified = file.time_last_modified  # Directly access if it's a datetime object

            # Format times as strings
            time_created = time_created.strftime("%Y-%m-%d %H:%M:%S") if time_created else None
            time_modified = time_modified.strftime("%Y-%m-%d %H:%M:%S") if time_modified else None

            #Ensure that the author property is loaded (slows down retrieval)
            # file.context.load(file, ["Author"])  # Explicitly load the author property 
            # file.context.execute_query()
            #Access the author data (returns email)
            # author_data = file.author
            # print(author_data)
            
            file_metadata.append({
                "name": file.name,
                "server_relative_url": file.serverRelativeUrl,
                "time_created": time_created,
                "time_last_modified": time_modified,
                "author": file.author,
                "editor": editor,
                "uniqueId": file.unique_id,
            })
            print(f' - {file.name}')  # Print only the name of each file

        return file_metadata
    
    
    def upload_file(self, file_name, custom_folder_path="", folder_name="", create_folder = False, personal = False):
        """
        Uploads a file to SharePoint. Optionally, a folder can be created with the same name as the file.

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

        # Construct the target SharePoint folder URL
        target_folder_url = self._select_folder(custom_folder_path, personal)

        try:
            # Folder creation logic
            if create_folder:
                folder_name = os.path.splitext(file_name)[0]
                folder_list = self.list_files(target_folder_url)
                if folder_name not in folder_list:
                    # Create folder if it doesn't exist
                    new_folder_url = f"{target_folder_url}/{folder_name}"
                    self.conn.web.ensure_folder_path(new_folder_url).execute_query()
                    print(f"Folder '{folder_name}' created successfully.")
                    target_folder_url = new_folder_url

            # Get the target_folder object
            target_folder = self.conn.web.get_folder_by_server_relative_path(target_folder_url)

            # Upload the file            
            upload_status = target_folder.upload_file(file_name, content).execute_query()
            if upload_status:
                print(f"File '{file_name}' uploaded successfully to '{target_folder_url}'.")
                return True
        except Exception as e:
            print(f'No se pudo subir el archivo {file_name} a la carpeta "{target_folder_url}')
            return False


    def allocate_files_from_folder(self, dictionary, personal):
        """
        No tiene mucha utilidad por ahora

        Args:
            folder_name (str)

        Returns:
            uploaded_files (list)
        """
        uploaded_files = []
        local_folder_path = UPLOAD_PATH

        # Verificar que la carpeta local exista
        if not os.path.exists(local_folder_path):
            print(f"La carpeta local {local_folder_path} no existe.")
            return uploaded_files  # Si la carpeta no existe, retornar lista vacía
        
        try:
            for file_name in os.listdir(local_folder_path):
                file_path = os.path.join(local_folder_path, file_name)

                if os.path.isfile(file_path):
                    try:
                        upload_status = self.upload_file(file_name, )

                        if upload_status:
                            uploaded_files.append(file_name)
                    except Exception as e:
                        print(f"Error subiendo el archivo {file_name}: {e}")

        except Exception as e:
            print(f"Error al acceder a la carpeta local {local_folder_path}: {e}")

        print(f"Total de archivos subidos: {len(uploaded_files)}")
        return uploaded_files
        

    def download_file(self, file_url, file_name):
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
        
        
    def download_single_file(self, file_name, custom_folder_path = "", personal = False):
        """
        Falta modularizar
        """
        target_folder_url = self._select_folder(custom_folder_path, personal)
        file_url = f'{target_folder_url}/{file_name}'

        try:
            local_file_path = self.download_file(file_url, file_name)
        except Exception as e:
            print(f"Error al descargar el archivo '{file_name}': {e}")
        return local_file_path
    
    def download_files_from_folder(self, custom_folder_path="", folder_name="", personal=False, extension=""):
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
        target_folder_url = self._select_folder(custom_folder_path, folder_name, personal)

        # Listar los archivos de la carpeta
        files_metadata = self.list_files(custom_folder_path, folder_name, personal)
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

    

    
