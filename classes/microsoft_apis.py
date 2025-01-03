from dotenv import load_dotenv
import os
import pandas as pd
from exchangelib import Account, Credentials, DELEGATE
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import imaplib
import email
from email.header import decode_header
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date, datetime, timedelta
from pathlib import PurePath
import json
import re
from collections import defaultdict


# Cargar variables de entorno
load_dotenv()

# Variables globales
SUBJECT_FILTER = os.getenv("SUBJECT_FILTER")
DOWNLOAD_PATH = r'C:\Users\SALVADOR\OneDrive\CEPLAN\CeplanPythonCode\microsoft\descargas'  # Carpeta de descargas
UPLOAD_PATH = r'C:\Users\SALVADOR\OneDrive\CEPLAN\CeplanPythonCode\microsoft\descargas\clasificados'  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = r'C:\Users\SALVADOR\OneDrive\CEPLAN\CeplanPythonCode\microsoft\email_templates' # Carpeta desde la que se obtendrán los email templates

# Credenciales Outlook
IMAP_PORT = os.getenv("IMAP_PORT")  # Puerto para conexión segura
IMAP_SERVER = os.getenv("IMAP_SERVER")  
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_SENDER_EMAIL = os.getenv("OUTLOOK_SENDER_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = os.getenv("SMTP_PORT")

# Credenciales Sharepoint
SHAREPOINT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
SHAREPOINT_URL_SITE = os.getenv("SHAREPOINT_URL_SITE") # Ruta fija (Enlace)
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME") # Nombre de la ruta fija (DNPE)
SHAREPOINT_FOLDER = os.getenv("SHAREPOINT_FOLDER") # Ruta del canal (Documentos compartidos/AOI Tendencias)
SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC") # Ruta específica del folder (Prueba)



# Clases
class OutlookRetriever:
    def __init__(self):
        self.mail=None
        
    def _auth(self):
        print("- Intentando establecer conexión con Outlook...")
        try:
            self.mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            self.mail.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
            print(f'- Conexión IMAP exitosa para el correo {OUTLOOK_EMAIL}')
        except imaplib.IMAP4.error as e:
            raise ValueError(f"- Error de conexión IMAP: {e}")
        except Exception as e:
            raise ValueError(f"- Unexpected error during IMAP authentication: {e}")
            

    @staticmethod
    def safe_decode(value, encoding='utf-8'):
        """
        Decodes a value while handling unknown or invalid encodings gracefully.
        """
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        try:
            # Attempt to decode with the provided encoding
            decoded_value = value.decode(encoding, errors='replace')
        except (LookupError, UnicodeDecodeError):
            try:
                # Fallback to latin1 as a last resort
                decoded_value = value.decode('latin1', errors='replace')
            except Exception:
                # If all fails, return a string representation of the bytes
                decoded_value = str(value)
        
        return decoded_value
        
    def decode_text(self, text):
        """
        Decodifica un texto codificado en formato MIME usando safe_decode.
        Maneja tanto asuntos como nombres de archivos.
        """
        if not text:
            return ""
        
        try:
            # Decodificar los headers MIME
            decoded_parts = decode_header(text)
            result = ""
            
            for bytes_or_str, charset in decoded_parts:
                # Si es bytes o string, usar safe_decode
                if charset == 'unknown-8bit':
                    charset = 'iso-8859-1'
                decoded_part = self.safe_decode(bytes_or_str, charset or 'utf-8')
                result += decoded_part
                
            return result.strip()
            
        except Exception as e:
            print(f"Error decoding MIME text: {e}")
            # Si falla la decodificación MIME, intentar safe_decode directamente
            return self.safe_decode(text)
                
    def get_emails(self, start_date=None, subject_filter=SUBJECT_FILTER, parameter="ALL"):
        """
        Filters emails by date or subject (optional). If no date is provided, retrieves all emails.

        Args:
            start_date (str): Date from which to retrieve emails in the format 'DD-Mon-YYYY' (e.g., "01-Jan-2023").
            subject_filter (str): Text that must be present in the email subject.
            limit (int): Maximum number of emails to retrieve.
            parameter (str): Additional search parameter (e.g., 'ALL', 'RECENT').

        Returns:
            dictionary: a dict containing relevant data of every email fetched
        """
        if not self.mail:
            raise ValueError("Debes autenticarte usando el método `_auth`")
        email_data = {}
        try:
            self.mail.select("INBOX")

            if start_date:
                try:
                    datetime.strptime(start_date, "%d-%b-%Y")  # Validate the date format
                    search_criteria = f"SINCE {start_date.upper()}"
                except ValueError:
                    print("The date format is incorrect. It should be 'DD-Mon-YYYY' (e.g., 01-Jan-2023).")
                    return {}
            else:
                search_criteria = parameter

            # Retrieving emails
            status, messages = self.mail.search(None, search_criteria)
            if status != "OK":
                print("Failed to retrieve emails from the inbox.")
                return {}

            message_ids = messages[0].split()

            # Obtain filtered emails by id
            for msg_id in message_ids:
                try:
                    status, msg_data = self.mail.fetch(msg_id, "(RFC822)")
                    if status != "OK":
                        print(f"Error retrieving the info of email with ID {msg_id}.")
                        continue

                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1]) # contains all email parts in binary

                            # Extract details
                            subject = self.decode_text(msg.get("Subject", ""))
                            if subject_filter and subject_filter.lower() not in subject.lower():
                                    continue # Skip emails that don't match subject filter before continuing
                            sender = self.decode_text(msg.get("From", ""))  # Use safe_decode to decode the sender
                            recipient = self.decode_text(msg.get("To", ""))
                            date = self.decode_text(msg.get("Sent", ""))
                            body = self.decode_text(msg.get("Body", ""))   
                            attachments = []

                            # Extract attachments
                            for part in msg.walk():  # part meansemail part; walk() iterates over multiple attachments as well
                                if part.get_content_disposition() == "attachment":
                                    attachment_name = part.get_filename()
                                    if attachment_name:
                                        # Decode file name
                                        try:
                                            attachment_name = self.decode_text(attachment_name)
                                            # Limpiar caracteres no válidos en el nombre del archivo
                                            attachment_name = re.sub(r'[<>:"/\\|?*\r\n]', '', attachment_name)
                                        except Exception as decode_err:
                                            print(f"Error decoding the attachment name '{attachment_name}' from email with ID {msg_id}: {decode_err}")
                                            continue  # Saltar este archivo si hay un error en el nombre

                                        attachments.append(attachment_name)

                            # Extract name and email from "From" field
                            sender_match = re.match(r"([a-zA-Z\s]+) <(.+)>", sender)
                            if sender_match:
                                from_name = sender_match.group(1).strip()  # Extract name before <email>
                                from_email = sender_match.group(2)  # Extract the email inside <>

                            # Store in dictionary with updated format
                            email_data[msg_id.decode()] = {
                                "from_name": from_name,
                                "from_email": from_email,
                                "sent": date,
                                "to": recipient,
                                "subject": subject,
                                "body": body,
                                "attachments": attachments
                            }
                                      
                except Exception as e:
                    print(f"Error processing email with ID {msg_id}: {e}")

            print(f'- Se han obtenido {len(message_ids)} IDs de correos luego de aplicar el filtro')
            return email_data

        except Exception as e:
            print(f"Error retrieving emails: {e}")
            return {}


    def download_attachments(self, email_data):
        """
        Downloads attachments from email data dictionary

        Args:
            email_data (dict)

        Returns:
            email_data (dict): 
        """
        attachment_names = []
        for msg_id, msg_dict in email_data.items():
            try:
                self.mail.select("INBOX")
                status, msg_data = self.mail.fetch(msg_id.encode(), "(RFC822)")
                if status != "OK":
                    print(f"Error retrieving email with ID {msg_id}.")
                    continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        
                        for part in msg.walk():   # part means email part; walk() iterates over multiple attachments as well
                            if part.get_content_disposition() == "attachment":
                                attachment_name = part.get_filename()
                                if attachment_name:
                                    # Decodificar el nombre del archivo
                                    attachment_name = self.decode_text(attachment_name)
                                    # Limpiar caracteres no válidos en el nombre del archivo
                                    attachment_name = re.sub(r'[<>:"/\\|?*\r\n]', '', attachment_name)
                                    
                                    try:
                                        # Crear el path y guardar el archivo
                                        file_path = os.path.join(DOWNLOAD_PATH, attachment_name)
                                        with open(file_path, "wb") as f:
                                            f.write(part.get_payload(decode=True))

                                        # Añadir el nombre del archivo original a la lista
                                        attachment_names.append(attachment_name)
                                        
                                    except Exception as e:
                                        print(f"Error saving attachment '{attachment_name}' from email {msg_id}: {e}")

            except Exception as e:
                print(f"Error processing email with ID {msg_id} when downloading: {e}")
        
        print(f"- Total files downloaded: {len(attachment_names)}")
        return attachment_names
        

    @staticmethod
    def get_user_attachments(email_data):
        """
        Creates a new dictionary with senders as keys and attachment

        Args:
            email_data (dict): dictionary of email data as obtained from get_emails()_

        Returns:
            user_attachments: new dict
        """
        user_attachments = defaultdict(list)
        for single_email_data in email_data.values():
            sender = single_email_data["from"]
            user_attachments[sender].extend(single_email_data["attachments"])
        return user_attachments


   
class OutlookSender:
    def __init__(self):
        self.smtp_server = None

    def _auth(self):
        try:
            print("- Intentando establecer conexión con el servidor SMTP...")
            self.smtp_server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            self.smtp_server.starttls()  # Asegura la conexión con TLS
            self.smtp_server.login(OUTLOOK_SENDER_EMAIL, OUTLOOK_PASSWORD)
            print("- Autenticación SMTP exitosa.")
        except smtplib.SMTPAuthenticationError as e:
            if "5.7.57" in str(e):
                raise ValueError("- Error de autenticación SMTP: Verifica si necesitas una contraseña de aplicación o permisos para SMTP.")
            else:
                raise ValueError(f"- Error de autenticación SMTP: {e}")
        except Exception as e:
            raise ValueError(f"- Error inesperado al autenticar: {e}")
    

    def send_email(self, recipient, subject, body, sender_name="Outlook Bot", body_type = "plain"):
        """Enviar un correo. Requiere que el usuario esté autenticado."""
        if not self.smtp_server:
            raise ValueError("Debes autenticarte antes de enviar correos usando el método `auth`.")

        try:
            # Crear el mensaje
            msg = MIMEMultipart()
            msg["From"] = f"{sender_name} <{OUTLOOK_SENDER_EMAIL}>"
            msg["To"] = recipient
            msg["Subject"] = subject
            msg.attach(MIMEText(body, body_type))

            # Enviar el correo
            self.smtp_server.send_message(msg)
            print(f"- Correo enviado a {recipient}.")
        except Exception as e:
            print(f"- Error al enviar el correo: {e}")

    def send_emails_with_template(self, user_attachments_log, template_name, templates_path=TEMPLATES_PATH, sender_name= "Outlook Bot"):
        if not self.smtp_server:
            raise ValueError("Debes autenticarte antes de enviar correos usando el método `auth`.")
        if template_name.endswith(".html") or template_name.endswith(".htm"):
            body_type = "html"
        else:
            body_type = "txt"
        template_full_path= os.path.join(templates_path, template_name)

        if not os.path.exists(template_full_path):
            raise FileNotFoundError(f"La plantilla {template_full_path} no existe.")

        
        with open(template_full_path, "r", encoding= 'utf-8') as template_file:
            template = ' '.join(template_file.read().split())
            
        try:
            for sender, subdict in user_attachments_log.items():
                try:
                    attachments_body_details = []
                    for attachment, attachments_details in subdict.items():
                        original_name = attachments_details.get("original_name")
                        new_name = attachments_details.get("new_name")
                        path = attachments_details.get("path")

                        # Agregar detalles del archivo a la lista
                        attachments_body_details.append(
                            f'<li><strong>{original_name}</strong><ul><li>Nuevo nombre: {new_name}</li><li>Ubicación: {path}</li></ul></li>'
                        )
                    # Generar el cuerpo del correo uniendo los detalles
                    attachments_body_details_str = "".join(attachments_body_details)
                    body = template.format(attachments_details_body=attachments_body_details_str)
                    body = body.replace("\n", "").replace("\r", "")
                    self.send_email(sender, "Notificación de archivos subidos", body, body_type=body_type)
                except Exception as e:
                    print(f"No se pudo enviar un correo automático para {sender}: {e}")        
        except Exception as e:
            print(f"Hubo un error al leer el diccionario user_attachments_log: {e}")
        # Crear la lista de detalles de los archivos
                
    def logout(self):
        """Cerrar la conexión SMTP."""
        if self.smtp_server:
            self.smtp_server.quit()
            print("- Conexión SMTP cerrada.")
    



# outlook_session = OutlookSender()
# outlook_session._auth()
# outlook_session.send_email(
#     recipient="drios@ceplan.gob.pe",
#     subject="Holi",
#     body = "Este correo fue enviado de forma automática"
# )
# outlook_session.logout()




class EmailTemplate:
    def __init__(self, template_name, template_folder="templates"):
        """
        Inicializa la plantilla cargándola desde un archivo.

        :param template_name: Nombre del archivo de la plantilla (ejemplo: 'ficha_subida.txt')
        :param template_folder: Carpeta donde se almacenan las plantillas
        """
        self.template_path = os.path.join(template_folder, template_name)
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"La plantilla '{template_name}' no existe en '{template_folder}'.")
        
        with open(self.template_path, "r", encoding="utf-8") as f:
            self.template = f.read()
    
    def render(self, **kwargs):
        """
        Rellena la plantilla con los valores dinámicos.

        :param kwargs: Diccionario con valores a reemplazar (placeholders)
        :return: Texto final con los marcadores reemplazados
        """
        return self.template.format(**kwargs)
    
    def get_placeholders(self):
        """
        Identifica los marcadores dinámicos en la plantilla.

        :return: Lista de marcadores encontrados en la plantilla
        """
        return re.findall(r"{(.*?)}", self.template)



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



# sharepoint_session = Sharepoint()
# sharepoint_session._auth()
# lista_archivos, _ = sharepoint_session.list_files(target_folder="Tendencias/Tendencias Globales")
# print(lista_archivos)
#Sharepoint().download_file(SHAREPOINT_DOC, "t90.docx")
#Sharepoint().upload_file('t75 - recuperación de la solidaridad.docx', "Tendencias/Tendencias Nacionales")


# gestionar subir archivos solo en modo lectura


# outlook_session = OutlookRetriever()
# outlook_session._auth()
# emails = outlook_session.get_emails(start_date="6-Dec-2024")
# print(outlook_session.download_attachments(emails))


def convert_to_dataframe(emails_data):
    # Convierte el diccionario de datos de emails en un DataFrame
    df = pd.DataFrame.from_dict(emails_data, orient='index')
    
    # Asegurarse de que los archivos adjuntos se gestionen correctamente
    # (Si las listas de adjuntos no están vacías, convierte en una cadena, o la mantienes como lista).
    df['attachments'] = df['attachments'].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)
    
    return df

### CREAR UNA FUNCIÓN PARA CAMBIAR EL DICCIONARIO A UNO DE SENDERS_ATTACHMENTS PARA ENVIAR EL CORREO


#ruta = os.path.join(DOWNLOAD_PATH)
#FileManager(DOWNLOAD_PATH).list_files()
#ruta_clasificados= os.path.join(DOWNLOAD_PATH, "clasificados")
#print(FileManager(ruta_clasificados).list_files())


# ruta_json = r'C:\Users\SALVADOR\OneDrive\CEPLAN\CeplanPythonCode\datasets\info_obs.json'
# with open(ruta_json, "r", encoding='utf-8') as file:
#     info_obs = json.load(file)
# file_manager = FileManager(DOWNLOAD_PATH)
# print(DOWNLOAD_PATH)
# file_manager.classify_files(info_obs)
