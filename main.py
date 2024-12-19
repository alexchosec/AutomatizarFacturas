import win32com.client
import tempfile
import os
import shutil
import re

from utiles.log_config import setup_logging
from clases.IniciarSesionRequest import IniciarSesionRequest
from clases.ResponseGenericBE import ResponseGenericBE
from utiles.api import token_api
from utiles.api import upload_xml
from utiles.api import upload_file
from utiles.common import generar_nombre_unico
from utiles.common import comprimir_file
from utiles.common import descomprimir_zip
from utiles.common import guardar_archivo_outlook
from utiles.common import extraer_todos_archivos_unSoloDirectorio
from utiles.common import descargar_archivo_web
from utiles.common import load_key
from utiles.common import generate_key
from utiles.common import encrypt_text
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

logger.info(email)

email = email.split('@')[0]
email = "aaviles"

logger.info(f"Carpeta principal: {inbox.Name}")

try:
    subfolder = inbox.Folders["FacturacionElectronica"]
    logger.info(f"Subcarpeta seleccionada: {subfolder.Name}")
except KeyError:
    logger.error("La carpeta 'FacturacionElectronica' no existe dentro de la Bandeja de Entrada.")
    exit()


# Obtener correos de la subcarpeta
messages = subfolder.Items
logger.info(f"Cantidad de correos en '{subfolder.Name}': {len(messages)}")

# Filtrar los correos no leídos
unread_messages = [msg for msg in messages if msg.UnRead]

if len(unread_messages) == 0:    
    logger.info(f"No se encontraron correos no leídos '{subfolder.Name}")
    exit()

iniciarSesionRequest = IniciarSesionRequest(username, password)
iniciarSesionResponse = token_api(urlApi, iniciarSesionRequest)

if iniciarSesionResponse is None or iniciarSesionResponse.token == "":
    logger.info("Token no disponible")
    exit()

try:

    for i, message in enumerate(unread_messages, 0):

        DocumentoID = int(0)
        pdfValido = False
        urls_completas = []

        logger.error(f"Correo {message.Subject}")

        # Buscamos el XML del comprobante
        for j, attachment in enumerate(message.Attachments, 0):

            filename = attachment.Filename.lower()

            if not (filename.endswith(".xml") or filename.endswith(".zip") or filename.endswith(".rar")):         
                continue     
            
            logger.info(f"Adjunto encontrado: {filename}")

            # Guardar el archivo en una ruta temporal
            ruta_archivo_outlook = guardar_archivo_outlook(attachment)
            logger.info(f"Archivo guardado en: {ruta_archivo_outlook}")

            if filename.endswith(".zip") or filename.endswith(".rar"):    

                ruta_archivo_outlook_subdirectorio = extraer_todos_archivos_unSoloDirectorio(ruta_archivo_outlook)
                for nombre_archivo in os.listdir(ruta_archivo_outlook_subdirectorio):
           
                    if not nombre_archivo.lower().endswith(".xml"):
                        continue

                    # Ruta de archivo
                    ruta_archivo_outlook_subdirectorio_archivo = os.path.join(ruta_archivo_outlook_subdirectorio, nombre_archivo)

                    # Generar archivo zip
                    ruta_archivo_outlook_subdirectorio_archivo_zip = comprimir_file(ruta_archivo_outlook_subdirectorio_archivo)

                   # Opcional: eliminar el archivo zip después de enviarlo
                    os.remove(ruta_archivo_outlook_subdirectorio_archivo)
                    logger.info(f"Archivo zip eliminado: {ruta_archivo_outlook_subdirectorio_archivo}")
 
                    # Enviar archivo API
                    responseXml = upload_xml(urlApi, iniciarSesionResponse.token, email, ruta_archivo_outlook_subdirectorio_archivo_zip)   
    
                    # Opcional: eliminar el archivo zip después de enviarlo
                    os.remove(ruta_archivo_outlook_subdirectorio_archivo_zip)
                    logger.info(f"Archivo zip eliminado: {ruta_archivo_outlook_subdirectorio_archivo_zip}")
 
                    if responseXml.respuesta:
                        DocumentoID = int(responseXml.param2)                        
                        continue

                shutil.rmtree(ruta_archivo_outlook_subdirectorio)
                logger.info(f"Directorio zip eliminado: {ruta_archivo_outlook_subdirectorio}")
                
            else:
                        
                # Comprimir el archivo en un archivo ZIP              
                ruta_archivo_zip = comprimir_file(ruta_archivo_outlook) 
                logger.info(f"Archivo ZIP guardado en: {ruta_archivo_zip}")
                    
                # Enviar archivo API
                responseXml = upload_xml(urlApi, iniciarSesionResponse.token, email, ruta_archivo_zip)   

                # Opcional: eliminar el archivo zip después de enviarlo
                os.remove(ruta_archivo_zip)
                logger.info(f"Archivo zip eliminado: {ruta_archivo_zip}")

                if responseXml.respuesta:
                    DocumentoID = int(responseXml.param2)

                       
            # Opcional: eliminar el archivo original después de comprimirlo
            os.remove(ruta_archivo_outlook)
            logger.info(f"Archivo original eliminado: {ruta_archivo_outlook}")

            # En caso ya encontro y registro salir del bucle
            if DocumentoID > 0:
                break

        # XML - APM
        if DocumentoID == 0:

            # Solo para APM Terminals           
            urls_completas = re.findall(patronAPM, message.Body)

            if len(urls_completas) > 0:

                # En posicion 1 esta el XML
                ruta_archivo_zip = descargar_archivo_web(urls_completas[1])
                logger.info(f"Archivo ZIP guardado en: {ruta_archivo_zip}")
                   
                # Enviar archivo API
                responseXml = upload_xml(urlApi, iniciarSesionResponse.token, email, ruta_archivo_zip)   

                # Opcional: eliminar el archivo zip después de enviarlo
                os.remove(ruta_archivo_zip)
                logger.info(f"Archivo zip eliminado: {ruta_archivo_zip}")

                if responseXml.respuesta:
                    DocumentoID = int(responseXml.param2)
        

        if DocumentoID == 0:
            logger.error(f"El correo {message.Subject} no tiene un archivo XML valido")
            continue

        '''
            SUBIR TODOS LOS ARJUNTOS
        '''

        # Adjuntos del correo
        for j, attachment in enumerate(message.Attachments):
    
            filename = attachment.Filename.lower()
            logger.info(f"Adjunto encontrado: {filename}")

            if "image" in filename:  # Aquí filtramos imágenes en base al nombre del archivo
                logger.info(f"Imagen ignorada: {filename}")
                continue

            # Guardar el archivo en una ruta temporal
            ruta_archivo_outlook = guardar_archivo_outlook(attachment)
            logger.info(f"Archivo guardado en: {ruta_archivo_outlook}")

            isZip = ""
            ruta_archivo_zip = ""

            if filename.endswith(".zip") or filename.endswith(".rar"):                        
                isZip = "Y"  
                ruta_archivo_zip =  ruta_archivo_outlook            
            else:
                # Comprimir el archivo en un archivo ZIP              
                ruta_archivo_zip = comprimir_file(ruta_archivo_outlook) 
                logger.info(f"Archivo ZIP guardado en: {ruta_archivo_zip}")

                # Opcional: eliminar el archivo original 
                os.remove(ruta_archivo_outlook)
                logger.info(f"Archivo original eliminado: {ruta_archivo_outlook}")
                   
            # Enviar archivo API
            responseFile = upload_file(urlApi, iniciarSesionResponse.token, email, DocumentoID, isZip, ruta_archivo_zip)   

            if not responseFile.respuesta:                    
                logger.error(f"Error al subir : {attachment.Filename} : {responseFile.mensaje}")
            else:               
                logger.info(f"Exito al subir: {attachment.Filename}")

            # Opcional: eliminar el archivo zip después de enviarlo
            os.remove(ruta_archivo_zip)
            logger.info(f"Archivo zip eliminado: {ruta_archivo_zip}")
      
        
        # PDF - APM 
        if len(urls_completas) > 0:

            # En posicion 0 esta el PDF
            ruta_archivo_pdf = descargar_archivo_web(urls_completas[0])
            logger.info(f"Archivo PDF guardado en: {ruta_archivo_pdf}")
            
            # Comprimir archivo ZIP
            ruta_archivo_zip = comprimir_file(ruta_archivo_pdf)
            logger.info(f"Archivo ZIP guardado en: {ruta_archivo_zip}")
                
            # Opcional: eliminar el archivo pdf después de enviarlo
            os.remove(ruta_archivo_pdf)
            logger.info(f"Archivo PDF eliminado: {ruta_archivo_pdf}")

            # Enviar archivo API
            responseFile = upload_file(urlApi, iniciarSesionResponse.token, email, DocumentoID, "", ruta_archivo_zip)   

            # Opcional: eliminar el archivo zip después de enviarlo
            os.remove(ruta_archivo_zip)
            logger.info(f"Archivo ZIP eliminado: {ruta_archivo_zip}")



        if DocumentoID > 0:
            message.Unread = False
            message.Save() 
            logger.info(f"El correo {message.Subject} fue enviado correctamente") 


except Exception as e:
    logger.error(f"Error: {e}")