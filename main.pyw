import win32com.client
import tempfile
import os
import shutil
import re
import time

from utiles.log_config import setup_logging
from clases.IniciarSesionRequest import IniciarSesionRequest
from clases.CorreoRecibidoRequest import CorreoRecibidoRequest
from clases.NotificacionRequest import NotificacionRequest
from clases.ResponseGenericBE import ResponseGenericBE
from utiles.api import token_api
from utiles.api import upload_file
from utiles.api import save_email
from utiles.api import notificar_errores

from utiles.common import comprimir_file
from utiles.common import load_key
from utiles.common import generate_key
from utiles.common import decrypt_text
from utiles.common import leer_settings
from utiles.common import is_numeric


logger = setup_logging()


# Credenciales 
email = ""
patronAPM = r"https://portal\.efacturacion\.pe/visorComprobante/sat/vista/descarga\.jsf\?code=[\w/]+"

# Generar la clave (solo la primera vez)
# generate_key()

# Cargar la clave desde el archivo
key = load_key()

# Obtener ajustes
lineas = leer_settings()

# Obtener variables
username = decrypt_text(lineas[0], key) 
password = decrypt_text(lineas[1], key)
urlApi = decrypt_text(lineas[2], key)
urlApi = "http://localhost:5194"

# Conectar con la aplicación de Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtener la carpeta de Bandeja de Entrada (puede ser cualquier carpeta válida)
inbox = namespace.GetDefaultFolder(6)  # 6 es la Bandeja de Entrada

# Obtener el correo SMTP del usuario actual
recipient = inbox.FolderPath.split("\\")[1]  # Primera carpeta del usuario
address_entry = namespace.CurrentUser.AddressEntry

if address_entry.Type == "EX":  # Si es un usuario de Exchange
    smtp_address = address_entry.GetExchangeUser().PrimarySmtpAddress
    email = smtp_address
else:
    email = address_entry.Address

# email = email.split('@')[0]
# email = "aaviles"

# logger.info(f"Carpeta principal: {inbox.Name}")

try:
    subfolder = inbox.Folders["FacturacionElectronica"]
except KeyError:
    exit()


messages = subfolder.Items

# Filtrar los correos no leídos
unread_messages = [msg for msg in messages if msg.UnRead]
# unread_messages = [msg for msg in messages if msg.Class == 43 and msg.UnRead]

if len(unread_messages) == 0:    
    logger.info(f"No se encontraron correos no leídos '{subfolder.Name}")
    time.sleep(3)
    exit()
else:
    logger.info(f"Correos no leídos '{len(unread_messages)}")


logger.info(email)

iniciarSesionRequest = IniciarSesionRequest(username, password)
iniciarSesionResponse = token_api(urlApi, iniciarSesionRequest)

extensiones_compresion = [
    ".zip", 
    ".rar", 
    ".tar", 
    ".gz", 
    ".7z", 
    ".bz2", 
    ".xz", 
    ".iso", 
    ".cab", 
    ".zipx", 
    ".tar.xz", 
    ".z", 
    ".lzh"
]

if iniciarSesionResponse is None or iniciarSesionResponse.token == "":
    logger.info("Token no disponible")
    exit()

try:
    
    notificaciones = []
    
    for i, message in enumerate(unread_messages, 0):

        message.UnRead = False

        remitente_correo = message.SenderEmailAddress  
        if message.SenderEmailType == "EX":
            remitente_correo = message.Sender.GetExchangeUser().PrimarySmtpAddress

        correoRecibidoRequest = CorreoRecibidoRequest(remitente=remitente_correo, asunto=message.Subject, cuerpoMensaje=message.Body, usuario=email)
        respuesta = save_email(urlApi, iniciarSesionResponse.token, correoRecibidoRequest)

        if not is_numeric(respuesta):
            notificaciones.append(NotificacionRequest(asunto=message.Subject, mensaje=respuesta))            
            continue
        else:
            logger.info(f"Correo guardado: {message.Subject}")

        for attachment in message.Attachments:
                                       
            temp_dir = tempfile.mkdtemp()

            file_path = os.path.join(temp_dir, attachment.FileName)
            attachment.SaveAsFile(file_path)

            _, ext = os.path.splitext(attachment.FileName)

            descomprimir = ""
            file_zip = ""

            if ext.lower() in extensiones_compresion:
                descomprimir = "NO"
                file_zip = file_path
            else:
                descomprimir = "SI"
                file_zip = comprimir_file(file_path)

            respuestaAdjunto = upload_file(urlApi, iniciarSesionResponse.token, file_zip, respuesta, descomprimir)
            if respuestaAdjunto != "OK":
                notificaciones.append(NotificacionRequest(asunto=message.Subject, mensaje=respuestaAdjunto))
                logger.error(f"Error subiendo el archivo: {file_path}")
            else:
                logger.info(f"Archivo adjunto guardado: {file_path}")

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    if len(notificaciones) > 0:
        notificar_errores(urlApi, iniciarSesionResponse.token, email, notificaciones)

    time.sleep(5)

except Exception as e:
    logger.error(f"Error: {e}")