import os
import time
import zipfile
import win32com.client
import tempfile
import requests
import re
from cryptography.fernet import Fernet


def generar_nombre_unico(ruta):
    
    if not os.path.exists(ruta):
        return ruta
    
    nombre, extension = os.path.splitext(ruta)
    sufijo = time.strftime("%Y%m%d%H%M%S")  
    nuevo_nombre = f"{nombre}_{sufijo}{extension}"
    
    return nuevo_nombre

def comprimir_file(path_file):
    
    temp_dir = tempfile.mkdtemp()    
    nombre_archivo = os.path.splitext(os.path.basename(path_file))[0]

    zip_filename = os.path.join(temp_dir, f"{nombre_archivo}.zip")
    zip_filename = generar_nombre_unico(zip_filename)

    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(path_file, os.path.basename(path_file))  
    
    return zip_filename
            

def descomprimir_zip(path_zip):
    extract_to = None
    try:
        
        if not os.path.exists(path_zip):
            print(f"El archivo ZIP '{path_zip}' no existe.")
            return

        temp_dir = tempfile.mkdtemp()   
        nombre_directorio = os.path.splitext(os.path.basename(path_zip))[0]
        extract_to = os.path.join(temp_dir, nombre_directorio)  

        if not os.path.exists(extract_to):
            os.makedirs(extract_to)

        with zipfile.ZipFile(path_zip, 'r') as zipf:
            zipf.extractall(extract_to)
            print(f"Archivos descomprimidos en: {extract_to}")
   
    except zipfile.BadZipFile:
        print(f"El archivo '{path_zip}' no es un archivo ZIP válido.")
    except Exception as e:
        print(f"Error al descomprimir el archivo: {e}")
    return extract_to


def guardar_archivo_outlook(attachment):

    temp_dir = tempfile.mkdtemp()            
    temp_path = os.path.join(temp_dir, attachment.Filename)
    temp_path = generar_nombre_unico(temp_path) 
    attachment.SaveAsFile(temp_path)
    return temp_path

def descargar_archivo_web(url):
    ruta_archivo = None
    try:
        
        ruta_base = tempfile.mkdtemp()  

        response = requests.get(url)
        response.raise_for_status()  

        nombre_archivo = None
        url2 = response.url

        if url2:
            nombre_archivo = url2.split("/")[-1]
            ruta_archivo = os.path.join(ruta_base, nombre_archivo)
            with open(ruta_archivo, "wb") as archivo:
                archivo.write(response.content)
            print(f"Archivo guardado: {ruta_archivo}")
            
    except requests.RequestException as e:
        print(f"Error al descargar {url}: {e}")

    return ruta_archivo


def extraer_todos_archivos_unSoloDirectorio(zip_path):
    """
    Extrae todos los archivos de un ZIP a un solo nivel, sin respetar la estructura de directorios.   
    :param zip_path: Ruta del archivo ZIP.
    """
    extract_to = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if not file_info.is_dir():  # Ignorar directorios
                    # Generar una ruta temporal para el archivo extraído
                    file_name = os.path.basename(file_info.filename)
                    extracted_path = os.path.join(extract_to, file_name)
                    
                    with zip_ref.open(file_info.filename) as source:
                        # Guardar el archivo extraído temporalmente
                        with open(extracted_path, 'wb') as target:
                            target.write(source.read())
                    
                    # Verificar si el archivo extraído es otro ZIP
                    if zipfile.is_zipfile(extracted_path):
                        print(f"Descomprimiendo archivo ZIP anidado: {extracted_path}")
                        # Llamada recursiva para procesar el ZIP anidado
                        extract_all_to_flat_dir(extracted_path, extract_to)
                        # Eliminar el ZIP anidado después de extraer su contenido
                        os.remove(extracted_path)
                        
        print(f"Archivos extraídos a un solo nivel en: {extract_to}")
    except FileNotFoundError:
        print(f"El archivo ZIP no se encontró: {zip_path}")
    except zipfile.BadZipFile:
        print(f"El archivo especificado no es un ZIP válido: {zip_path}")
    except Exception as e:
        print(f"Error al extraer el ZIP: {e}")
    return extract_to




# Función para generar una clave (solo se genera una vez)
def generate_key():
    key = Fernet.generate_key()
    with open("key.key", "wb") as key_file:
        key_file.write(key)
    print("Clave generada y guardada en 'key.key'.")

# Función para cargar la clave
def load_key():
    with open("key.key", "rb") as key_file:
        return key_file.read()

# Función para encriptar texto
def encrypt_text(text, key):
    fernet = Fernet(key)
    encrypted_text = fernet.encrypt(text.encode())
    return encrypted_text.decode()

# Función para desencriptar texto
def decrypt_text(encrypted_text, key):
    fernet = Fernet(key)
    decrypted_text = fernet.decrypt(encrypted_text.encode())
    return decrypted_text.decode()

def leer_settings():
    lineas = None
    archivo = "C:\LOG\settings.ini"
    with open(archivo, "r") as file:
        lineas = file.readlines()
    return lineas
