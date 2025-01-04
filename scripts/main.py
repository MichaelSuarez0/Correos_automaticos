# # 0. Inicializar variables de entorno
import os
import re
import json
import sys
from pathlib import Path, PurePath
from shutil import move
from datetime import date, timedelta
from dotenv import load_dotenv
from exchangelib import Account, Credentials, DELEGATE
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import imaplib
import email
from email.header import decode_header
from correos_automaticos.classes.outlook_manager import OutlookRetriever, OutlookSender
from correos_automaticos.classes.file_manager import FileManager
from correos_automaticos.classes.sharepoint_manager import Sharepoint

import pprint
from collections import defaultdict

# 1: Root directory to upload files
#ROOT_DIR = sys.argv[1]

# 2: Sharepoint folder name
#SHAREPOINT_FOLDER_NAME = sys.argv[2]

# 3: dictionary to classify files

load_dotenv()

script_dir = os.path.dirname(__file__)

# Configuración global
IMAP_SERVER = os.getenv("IMAP_SERVER")  
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
SUBJECT_FILTER = os.getenv("SUBJECT_FILTER")
DOWNLOAD_PATH = os.path.join(script_dir, "..", 'descargas')
IMAP_PORT = os.getenv("IMAP_PORT")  # Puerto para conexión segura


# Diccionario para clasificar códigos de fichas
ruta_json_1 = os.path.join(script_dir, "..", "..", 'datasets', "rubros_subrubros.json")
with open(ruta_json_1, "r", encoding='utf-8') as file:
    rubros_subrubros = json.load(file)


# Diccionario para buscar metadata de las fichas y renombrar archivos a partir de "Título largo"
ruta_json_2 = os.path.join(script_dir, "..", "..", 'datasets', "info_obs.json")
with open(ruta_json_2, "r", encoding='utf-8') as file:
    info_obs = json.load(file)

ruta_renombrados= os.path.join(DOWNLOAD_PATH, "clasificados")


# 1. Definir funciones subordinadas
def find_file_path(file_name, regex_dict):
    """
    Clasifica un archivo según su nombre a partir de un diccionario de categorías.

    Args:
        file_name (str): El nombre del archivo a clasificar.
        diccionario (dict): Diccionario que define las categorías y sus patrones regex.

    Returns:
        str: El path de clasificación en formato 'rubro/subrubro' o 'rubro/subrubro/departamento' si es territorial. 
    """
    patron = file_name.split(" ")[0]  # Se asume que el patrón o código está antes del espacio
    
    for rubro, subdict in regex_dict.items():
        for subrubro, regex in subdict.items():
            
            # Caso 1: nivel simple (nacional, global)
            if isinstance(regex, str) and re.match(regex, patron, re.IGNORECASE):  # Añadí re.IGNORECASE para no diferenciar entre mayúsculas y minúsculas
                return f'{rubro}/{subrubro}'
            
            # Caso 2: nivel complejo (es territorial)
            if isinstance(regex, dict):
                for departamento, true_regex in regex.items():
                    if isinstance(true_regex, str) and re.match(true_regex, patron, re.IGNORECASE):
                        return f'{rubro}/{subrubro}/{departamento}'

    print(f"  No se encontró coincidencia para: {file_name}")
    return ""


def construct_log_dict(email_data, renamed_files_map):
    """
    Reconstructs email_data dict to another dict suitable for sending confirmation emails to users based on uploaded attachments

    Args:
        email_data (dict): Dict obtained from get_emails.
        renamed_files_map (dict): Dict mapping original names as keys and new names as values

    Returns:
        dict: A dictionary with senders as keys, attachments as subkeys, and details as values.
    """
    user_attachments_log = {}
    
    # Iterar sobre los valores en email_data
    for details in email_data.values():
        sender = details.get("from_email")

        if not sender:
            print("No sender found for email details:", details)  # Depuración
            continue

        # Inicializar diccionario para el remitente si no existe
        if sender not in user_attachments_log:
            user_attachments_log[sender] = {}
        
        # Obtener lista de adjuntos
        attachments = details.get("attachments", [])
        if not attachments:
            print(f"No attachments found for sender {sender}")  # Depuración
            continue
        
        # Iterar sobre los nombres originales de los archivos adjuntos
        for old_file_name in attachments:
            # Obtener el nuevo nombre del archivo del mapa de renombrados
            new_file_name = renamed_files_map.get(old_file_name)
            if not new_file_name:
                print(f"File {old_file_name} not found in renamed_files_map for sender {sender}")  # Depuración
                continue
            
            # Generar la ruta del nuevo archivo
            new_file_path = find_file_path(new_file_name, rubros_subrubros)
            
            # Construir la entrada del archivo
            user_attachments_log[sender][old_file_name] = {
                "original_name": old_file_name,
                "new_name": new_file_name,
                "path": new_file_path
            }

    return user_attachments_log

# Imprimir de manera legible
def print_dict(d):
    """
    Imprime un diccionario de manera más legible.
    """
    for sender, attachments in d.items():
        print(f"\nRemitente: {sender}")
        if attachments:
            for file_name, path in attachments.items():
                print(f"  Archivo: {file_name} -> Ruta: {path}")
        else:
            print("  No hay archivos clasificados.")

#print(find_file_category("tg1 ejemplo.txt", rubros_subrubros))  # Debería imprimir: "Tendencias/Tendencias Globales"
print(find_file_path("r23_madre ejemplo.txt", rubros_subrubros))  # Debería imprimir: "Tendencias/Tendencias Territoriales/Madre de Dios"


# ## 2. Definir funciones principales
def obtener_archivos(start_date: str):
    """_summary_

    Args:
        start_date (str, optional): _description_. 

    Returns:
        email_data (dict)
    """
    outlook_session = OutlookRetriever()
    outlook_session._auth()
    email_data = outlook_session.get_emails(start_date=start_date)
    outlook_session.download_attachments(email_data)
    return email_data


def renombrar_y_clasificar(search_directory = DOWNLOAD_PATH, email_data = {}):
    """_summary_

    Args:
        search_directory (str): Path, please change it from DOWNLOAD_PATH.
        dict_to_rename (dict): Defaults to info_obs.
        email_data (dict): 

    Returns:
        user_attachments_final (dict): Diccionario con senders como keys, attachments como subkeys y los paths como valores.
    """
    file_manager = FileManager(search_directory=search_directory)
    renamed_files_map = file_manager.rename_files(info_obs)
    #lista_nombres_archivos = FileManager(ruta_renombrados).list_files()
    user_attachments_log = construct_log_dict(email_data=email_data, renamed_files_map=renamed_files_map)
    return renamed_files_map, user_attachments_log



def upload_files_to_sharepoint(user_attachments_log):
    sharepoint_session = Sharepoint()
    sharepoint_session._auth()

    for sender, file_dict in user_attachments_log.items():
        for old_file_name, attachment_details in file_dict.items():
            path = attachment_details.get("path")
            new_file_name  = attachment_details.get("new_name")
            try:
                # Subir archivo a Sharepoint
                sharepoint_session.upload_file(new_file_name, path)
                # Actualizar estado en el log
                attachment_details["sharepoint_status"] = "uploaded"
            except Exception as e:
                attachment_details["sharepoint_status"] = "error"


def send_confirmation_emails(user_attachments_log):
    outlook_sender_session = OutlookSender()
    outlook_sender_session._auth()
    outlook_sender_session.send_emails_with_template(user_attachments_log, "sharepoint_success.html")
    outlook_sender_session.logout()


def main(start_date: str):
    email_data = obtener_archivos(start_date=start_date) # OutlookRetriever
    renamed_files_map, user_attachments_log = renombrar_y_clasificar(email_data=email_data) # FileManager
    upload_files_to_sharepoint(user_attachments_log) # Sharepoint
    send_confirmation_emails(user_attachments_log) # OutlookSender
    return email_data, user_attachments_log, renamed_files_map

if __name__ == "__main__":
    main("1-Jan-2025")
