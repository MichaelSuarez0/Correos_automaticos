from dotenv import load_dotenv
import os
import pandas as pd
import imaplib
import email
from email.header import decode_header
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date, datetime, timedelta
import re
from collections import defaultdict
import imaplib
import socket
from tenacity import retry, stop_after_attempt, wait_fixed
from correos_automaticos.classes.storing_models import EmailData

script_dir = os.path.dirname(__file__)

# Cargar variables de entorno
load_dotenv()

# Variables globales
SUBJECT_FILTER = os.getenv("SUBJECT_FILTER")
DOWNLOAD_PATH = os.path.join(script_dir, "..", "descargas") 
UPLOAD_PATH = os.path.join(script_dir, "..", "descargas", "clasificados")  # Carpeta desde donde se subirán archivos
TEMPLATES_PATH = os.path.join(script_dir, "..", "email_templates") # Carpeta desde la que se obtendrán los email templates

# Credenciales Outlook
IMAP_PORT = os.getenv("IMAP_PORT") 
IMAP_SERVER = os.getenv("IMAP_SERVER")  
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_SENDER_EMAIL = os.getenv("OUTLOOK_SENDER_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = os.getenv("SMTP_PORT")


class OutlookRetriever:
    def __init__(self):
        self.mail = None

    @retry(stop=stop_after_attempt(2), wait=wait_fixed(1))    
    def _auth(self):
        print("- Intentando establecer conexión con Outlook...")
        try:
            # Verify the server is reachable first
            socket.gethostbyname(IMAP_SERVER)
            
            self.mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            self.mail.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
            print(f'- Conexión IMAP exitosa para el correo {OUTLOOK_EMAIL}')
            
        except socket.gaierror as e:
            raise ValueError(f"- Error de resolución DNS: No se puede conectar a {IMAP_SERVER}. "
                           f"Verifique su conexión a internet y las variables IMAP_SERVER/PORT: {e}")
            
        except imaplib.IMAP4.error as e:
            raise ValueError(f"- Error de autenticación IMAP: Verifique sus credenciales: {e}")
            
        except Exception as e:
            raise ValueError(f"- Error inesperado durante la autenticación IMAP: {str(e)}")
            

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
                
    def get_emails(self, start_date=None, subject_filter=SUBJECT_FILTER, parameter="ALL") -> list[EmailData]:
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
        emails_data = []
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

                            # Store in a list with updated format
                            emails_data.append(EmailData(
                                msg_id = msg_id.decode(),
                                from_name= from_name,
                                from_email = from_email,
                                sent = date,
                                to = recipient,
                                subject = subject,
                                body = body,
                                attachments = attachments        
                            ))
                        
                                      
                except Exception as e:
                    print(f"Error processing email with ID {msg_id}: {e}")

            print(f'- Se han obtenido {len(message_ids)} IDs de correos luego de aplicar el filtro')
            return emails_data

        except Exception as e:
            print(f"Error retrieving emails: {e}")
            return []


    def download_attachments(self, emails_data: list[EmailData])-> list:
        """
        Downloads attachments from email data dictionary

        Args:
            email_data (dict)

        Returns:
            None 
        """
        attachment_names = []
        for email_data in emails_data:
            msg_id = email_data.msg_id
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
        Creates a new dictionary with senders as keys and attachments as values

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

    def send_emails_with_template(self, user_attachments_log: dict, template_name: str, templates_path=TEMPLATES_PATH, sender_name= "Outlook Bot"):
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
            for sender, file_list in user_attachments_log.items():
                try:
                    attachments_body_details = []
                    for attachments_details in file_list:
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



# outlook_session = OutlookRetriever()
# outlook_session._auth()
# emails = outlook_session.get_emails(start_date="6-Dec-2024")
# print(outlook_session.download_attachments(emails))


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
