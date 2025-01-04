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
SHAREPOINT_FOLDER = os.getenv("SHAREPOINT_FOLDER") # Ruta del canal (Documentos compartidos/AOI Tendencias)
SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC") # Ruta específica del folder (Prueba)


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
    
    def list_files(self, target_folder = "", target_folder_url=""):
        files_list = []
        if not target_folder_url:
            if target_folder:
                target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}{SHAREPOINT_FOLDER}/{SHAREPOINT_DOC}/{target_folder}' 
            else:
                target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}{SHAREPOINT_FOLDER}/{SHAREPOINT_DOC}' 
        folder_name = target_folder_url.split("/")[-1]
        try:
            root_folder = self.conn.web.get_folder_by_server_relative_url(target_folder_url)   # Concatenates and formats the path
            root_folder.expand(["Files", "Folders"]).get().execute_query()  # Preguntar qué hace esto y por qué ese expand
            print(f'Archivos presentes en la carpeta "{folder_name}":')
        except Exception as e:
            print(f"No se encontró la carpeta con el path {folder_name}")
        #print(root_folder.files)
        for file in root_folder.files:
            print(f' - {file.name}')  # Imprime solo el nombre de cada archivo
            files_list.append(file.name)
            #print(file.content)
        return files_list, root_folder.files
    
    
    def upload_file(self, file_name, target_folder_path="", create_folder = False):
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
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}{SHAREPOINT_FOLDER}/{SHAREPOINT_DOC}'
        if target_folder_path:
            target_folder_url = f'{target_folder_url}/{target_folder_path}'     

        try:
            # Folder creation logic
            if create_folder:
                folder_name = os.path.splitext(file_name)[0]
                folder_list = self.list_files(target_folder_url=target_folder_url)
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


    def upload_files_from_folder(self, folder_name):
        """
        No tiene mucha utilidad por ahora

        Args:
            folder_name (str)

        Returns:
            uploaded_files (list)
        """
        uploaded_files = []
        if folder_name:
            local_folder_path = os.path.join(UPLOAD_PATH, folder_name) 

        # Verificar que la carpeta local exista
        if not os.path.exists(local_folder_path):
            print(f"La carpeta local {local_folder_path} no existe.")
            return uploaded_files  # Si la carpeta no existe, retornar lista vacía
        
        try:
            for file_name in os.listdir(local_folder_path):
                file_path = os.path.join(local_folder_path, file_name)

                if os.path.isfile(file_path):
                    try:
                        upload_status = self.upload_file(file_path, folder_name)

                        if upload_status:
                            uploaded_files.append(file_name)
                    except Exception as e:
                        print(f"Error subiendo el archivo {file_name}: {e}")

        except Exception as e:
            print(f"Error al acceder a la carpeta local {local_folder_path}: {e}")

        print(f"Total de archivos subidos: {len(uploaded_files)}")
        return uploaded_files
    
    def download_file(self, file_name):
        """
        Falta modularizar
        """
        download_path = DOWNLOAD_PATH
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}{SHAREPOINT_FOLDER}/{SHAREPOINT_DOC}/{file_name}' 
        file = File.open_binary(self.conn, file_url) # Preguntar qué hace esto
        # Ruta local donde quieres guardar el archivo
        local_file_path = os.path.join(download_path, file_name)
    
        # Escribir el contenido binario en el archivo local
        with open(local_file_path, 'wb') as local_file:
            local_file.write(file.content)
    
        print(f"El archivo {file_name} ha sido descargado con éxito en {local_file_path}.")
        return local_file_path
    
    def download_files(self, target_folder, extension=""):
        return None
    

    
