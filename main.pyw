import win32com.client
import tempfile
import os
import shutil
import re
import time

from utiles.log_config import setup_logging
from clases.IniciarSesionRequest import IniciarSesionRequest
from clases.NotificacionRequest import NotificacionRequest
from clases.ResponseGenericBE import ResponseGenericBE
from utiles.api import token_api
from utiles.api import upload_file
from utiles.api import notificar_errores
# from utiles.common import generar_nombre_unico
from utiles.common import comprimir_file
# from utiles.common import descomprimir_zip
# from utiles.common import guardar_archivo_outlook
# from utiles.common import extraer_todos_archivos_unSoloDirectorio
# from utiles.common import descargar_archivo_web
from utiles.common import load_key
from utiles.common import generate_key
# from utiles.common import encrypt_text
from utiles.common import decrypt_text
from utiles.common import leer_settings

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
    # logger.info(f"Subcarpeta seleccionada: {subfolder.Name}")
except KeyError:
    # logger.error("La carpeta 'FacturacionElectronica' no existe dentro de la Bandeja de Entrada.")
    exit()


# Obtener correos de la subcarpeta
messages = subfolder.Items
# logger.info(f"Cantidad de correos en '{subfolder.Name}': {len(messages)}")

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

if iniciarSesionResponse is None or iniciarSesionResponse.token == "":
    logger.info("Token no disponible")
    exit()

try:
    
    notificaciones = []
    temp_dir = tempfile.mkdtemp()

    for i, message in enumerate(unread_messages, 0):

        logger.info(f"Procesando correo {i}: {message.Subject}")

        nombre_archivo_correo = f"Correo_{i}.eml"
        ruta_archivo_correo = os.path.join(temp_dir, nombre_archivo_correo)
        message.SaveAs(ruta_archivo_correo, 1) 
        logger.debug(f"Correo guardado en {ruta_archivo_correo}")

        zip_nombre = f"Correo_{i}.zip"
        zip_ruta = os.path.join(temp_dir, zip_nombre)
        ruta_archivo_correo = comprimir_file(ruta_archivo_correo)

        if not ruta_archivo_correo:
            notificaciones.append(NotificacionRequest(asunto=message.Subject, mensaje=f"No se pudo comprimir el archivo {ruta_archivo_correo}"))                        
            continue

        respuesta = upload_file(urlApi, iniciarSesionResponse.token, email, ruta_archivo_correo)
        message.UnRead = False

        if respuesta == "OK":
            logger.info(f"Archivo ZIP subido exitosamente: {zip_ruta}")            
        else:
            notificaciones.append(NotificacionRequest(asunto=message.Subject, mensaje=respuesta))            
            logger.error(f"Error subiendo el archivo ZIP: {zip_ruta}")


    if len(notificaciones) > 0:
        notificar_errores(urlApi, iniciarSesionResponse.token, email, notificaciones)

    time.sleep(5)

except Exception as e:
    logger.error(f"Error: {e}")