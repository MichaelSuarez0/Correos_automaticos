# # 0. Inicializar variables de entorno
import os
import re
import json
from shutil import move
from datetime import date, timedelta
from dotenv import load_dotenv
from correos_automaticos.classes.outlook_manager import OutlookRetriever, OutlookSender
from correos_automaticos.classes.file_manager import FileManager
from correos_automaticos.classes.sharepoint_manager import Sharepoint
from pprint import pprint
from collections import defaultdict
from icecream import ic
import logging


# -------------------------------------------------------------
# --------------- 0. Definir variables globales ---------------
# -------------------------------------------------------------
# Configuración básica del logging
logging.basicConfig(
    level=logging.INFO,  # Nivel de registro (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    format="%(asctime)s - %(levelname)s - %(message)s",  # Formato del mensaje
    datefmt="%Y-%m-%d %H:%M:%S"  # Formato de fecha y hora
)

# Configuración global

load_dotenv()
script_dir = os.path.dirname(__file__)
SHAREPOINT_URL = "https://ceplangobpe.sharepoint.com/sites/DNPE"
SHAREPOINT_FOLDER = "Documentos compartidos"
SHAREPOINT_FOLDER_TENDENCIAS = f'{SHAREPOINT_FOLDER}/AOI Tendencias'
SHAREPOINT_FOLDER_RYO = f'{SHAREPOINT_FOLDER}/AOI Riesgos y oportunidades'
DOWNLOAD_PATH = os.path.join(script_dir, "..", 'descargas')
SUBJECT_FILTER = "Sistematizar"

# Diccionario para clasificar códigos de fichas
ruta_json_1 = os.path.join(script_dir, "..", "..", 'datasets', "rubros_subrubros.json")
with open(ruta_json_1, "r", encoding='utf-8') as file:
    rubros_subrubros = json.load(file)

# Diccionario para renombrar archivos a partir de "Título largo"
ruta_json_2 = os.path.join(script_dir, "..", "..", 'datasets', "info_obs.json")
with open(ruta_json_2, "r", encoding='utf-8') as file:
    info_obs = json.load(file)

# --------------------------------------------------------------
# ------------- 1. Definir funciones subordinadas --------------
# --------------------------------------------------------------
def find_file_path(file_name: str, regex_dict: dict = rubros_subrubros) -> str:
    """
    Clasifica una ficha según su nombre a partir del diccionario rubros_subrubros.

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


def construct_user_attachments(email_data: dict, renamed_files_map: list)-> dict:
    """
    Reconstructs email_data dict to another dict suitable for sending confirmation emails to users based on uploaded attachments

    Args:
        email_data (dict): Dict obtained from get_emails.
        renamed_files_map (dict): Dict obtained from rename_files

    Returns:
        dict: A dictionary with senders as keys, attachments as subkeys, and details as values.
    """
    user_attachments_log = {}
    
    # Iterar sobre los valores en email_data
    for details in email_data.values():
        sender = details.get("from_email")

        if not sender:
            logging.info("No sender found for email details:", details) 
            continue

        # Inicializar diccionario para el remitente si no existe
        if sender not in user_attachments_log:
            user_attachments_log[sender] = []
        
        # Obtener lista de adjuntos
        attachments = details.get("attachments", [])
        if not attachments:
            logging.info(f"No attachments found for sender {sender}")  # Depuración
            continue
        
        # Iterar sobre los nombres originales de los archivos adjuntos
        for old_file_name in attachments:
            new_file_name = ""

            # Obtener el nuevo nombre del archivo del mapa de renombrados
            for file_dict in renamed_files_map:
                original_name = file_dict.get('original_name', '')
                if original_name == old_file_name:
                    new_file_name = file_dict.get('new_name')
                
                    # Construir la entrada del archivo
                    user_attachments_log[sender].append({
                        "original_name": old_file_name,
                        "new_name": file_dict.get('new_name'),
                        "path": find_file_path(new_file_name)
                    })
                    break
            else:
                logging.info(f"File {old_file_name} not found in renamed_files_map for sender {sender}")  # Depuración

    return user_attachments_log


#print(find_file_category("tg1 ejemplo.txt", rubros_subrubros))  # Debería imprimir: "Tendencias/Tendencias Globales"
#print(find_file_path("r23_madre ejemplo.txt", rubros_subrubros))  # Debería imprimir: "Tendencias/Tendencias Territoriales/Madre de Dios"

# -------------------------------------------------------------
# ------------- 2. Definir funciones principales --------------
# -------------------------------------------------------------
def obtener_archivos(start_date: str):
    """_summary_

    Args:
        start_date (str, optional): _description_. 

    Returns:
        email_data (dict)
    """
    outlook_session = OutlookRetriever()
    outlook_session._auth()
    email_data = outlook_session.get_emails(start_date=start_date, subject_filter=SUBJECT_FILTER)
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
    user_attachments_log = construct_user_attachments(email_data, renamed_files_map)
    return user_attachments_log


def upload_files_to_sharepoint(user_attachments_log: dict):
    sharepoint_tendencias = Sharepoint(SHAREPOINT_URL, SHAREPOINT_FOLDER_TENDENCIAS, connect_on_creation=False)
    sharepoint_ryo = Sharepoint(SHAREPOINT_URL, SHAREPOINT_FOLDER_RYO, connect_on_creation=False)

    for sender, file_list in user_attachments_log.items():
        for attachment_details in file_list:
            path = attachment_details.get("path")
            path_folders = path.split("/")

            # Determinar en qué SharePoint subir el archivo
            if path_folders[0].upper() == "TENDENCIAS":
                sharepoint_session = sharepoint_tendencias
                custom_folder_path = f'{SHAREPOINT_FOLDER}/AOI Tendencias/{path[1:]}'
            else:
                sharepoint_session = sharepoint_ryo
                custom_folder_path = f'{SHAREPOINT_FOLDER}/AOI Riesgos y oportunidades/{path[1:]}'

            # Construir la ruta de la carpeta según la condición
            if len(path_folders) < 3:
                # OBTENER EL DEPARTAMENTO DEL CÓDIGO
                dpto = ""
                custom_folder_path = f'{custom_folder_path}/{dpto}'

            new_file_name = attachment_details.get("new_name")

            try:
                # Subir archivo a SharePoint
                sharepoint_session.upload_file(new_file_name, custom_folder_path, create_folder=True)
                # Actualizar estado en el log
                attachment_details["sharepoint_status"] = "uploaded"
            except Exception as e:
                attachment_details["sharepoint_status"] = "error"

    return user_attachments_log

def send_confirmation_emails(user_attachments_log):
    outlook_sender_session = OutlookSender()
    outlook_sender_session.send_emails_with_template(user_attachments_log, "sharepoint_success.html")
    outlook_sender_session.logout()



# -------------------------------------------------------------
# ------------------------- 3. MAIN ---------------------------
# -------------------------------------------------------------
def main(start_date: str):
    email_data = obtener_archivos(start_date)                                    # OutlookRetriever
    user_attachments_log = renombrar_y_clasificar(DOWNLOAD_PATH, email_data)     # FileManager
    user_attachments_log= upload_files_to_sharepoint(user_attachments_log)       # Sharepoint
    send_confirmation_emails(user_attachments_log)                               # OutlookSender
    ic(email_data)
    ic(user_attachments_log)

if __name__ == "__main__":
    main("21-Feb-2025")
