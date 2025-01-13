import requests
import json
import os
import mimetypes

from typing import Union, List
from clases.IniciarSesionResponse import IniciarSesionResponse
from clases.ResponseGenericBE import ResponseGenericBE
from clases.NotificacionRequest import NotificacionRequest

def token_api(url_api, request):
    response = None
    try:
        
        request_data = request.to_dict()

        json_data = json.dumps(request_data)

        headers = {'Content-Type': 'application/json'}

        url = f"{url_api}/api/Autenticar/IniciarSesion"
        
        respuesta = requests.post(url, data=json_data, headers=headers)

        if respuesta.status_code == 200:
            
            data = respuesta.json()
            
            response = IniciarSesionResponse(
                expires=data['expires'],
                username=data['username'],
                token=data['token']
            )

        else:
            print(f"Error al autenticar: {respuesta.status_code}")

    except Exception as e:
        print(f"Error en la solicitud API: {e}")

    return response

def save_email(url_api, token, request):
    response = None
    try:
        
        request_data = request.to_dict()

        json_data = json.dumps(request_data)

        headers = {
            'Content-Type': 'application/json',
            "Authorization": f"Bearer {token}"
        }

        url = f"{url_api}/api/FacturaProveedor/CorreoRecibidoRegistrar"
        
        respuesta = requests.post(url, data=json_data, headers=headers)

        if respuesta.status_code == 200:
            
            response = str(respuesta.json())  

        else:
            
            response = f"Error al registrar correo: {respuesta.status_code} - {respuesta.text}"

    except Exception as e:
        response = f"Error en la solicitud API: {e}"

    return response


def update_email(url_api, token, request):
    response = None
    try:
        
        request_data = request.to_dict()

        json_data = json.dumps(request_data)

        headers = {
            'Content-Type': 'application/json',
            "Authorization": f"Bearer {token}"
        }

        url = f"{url_api}/api/FacturaProveedor/CorreoRecibidoActualizar"        
        respuesta = requests.post(url, data=json_data, headers=headers)

        if respuesta.status_code == 200:           
            response = "OK" 
        else:           
            response = f"Error al registrar correo: {respuesta.status_code} - {respuesta.text}"

    except Exception as e:
        response = f"Error en la solicitud API: {e}"

    return response




def upload_file(url_api, token, file_path, id_email, unzip):
    response = ""
    try:
        
        headers = {
            "IdCorreo": str(id_email),
            "Authorization": f"Bearer {token}",
            "Descomprimir": unzip
        }

        if not os.path.exists(file_path):
            logger.error(f"Archivo no encontrado: {file_path}")
            return "Archivo no encontrado"

        mime_type, _ = mimetypes.guess_type(file_path)
        if mime_type is None:
            mime_type = 'application/octet-stream'  
            
        with open(file_path, 'rb') as file:
            
            files = {'file': (os.path.basename(file_path), file, mime_type)}

            respuesta = requests.post(f"{url_api}/api/FacturaProveedor/CorreoRecibidoArchivoRegistrar", headers=headers, files=files)
            
            if respuesta.status_code == 200:
                response = "OK"
            else:     
                try:
                    response = respuesta.json()  
                except ValueError:
                    response = respuesta.text 

    except Exception as e:        
        response = str(e)  

    return response



def notificar_errores(url_api: str, token: str, user: str, request: Union['NotificacionRequest', List['NotificacionRequest'], dict]):
    """
    Envía notificaciones de errores a una API.

    Args:
        url_api (str): La URL base de la API a la que se enviarán las notificaciones.
        token (str): Token de autenticación necesario para la autorización en la API.
        user (str): El nombre del usuario al que se le notificará el error.
        request (NotificacionRequest | List[NotificacionRequest] | dict): 
            Un objeto `NotificacionRequest`, una lista de objetos `NotificacionRequest`, 
            o un diccionario con los datos a enviar. Si es un diccionario, debe seguir el formato 
            esperado por la API.

    Raises:
        ValueError: Si el argumento `request` no es del tipo esperado.
    """
    try:

        # Verificar si request es una lista de NotificacionRequest
        if isinstance(request, list):

            # Convertir toda la lista de objetos NotificacionRequest en diccionarios
            request_data = [req.to_dict() for req in request if isinstance(req, NotificacionRequest)]
            
            # Asegurarse de que todos los elementos sean NotificacionRequest
            if len(request_data) != len(request):
                raise ValueError("Todos los elementos de la lista deben ser instancias de NotificacionRequest.")
            
            enviar_notificacion(url_api, token, user, request_data)

        # Verificar si request es una instancia de NotificacionRequest
        elif isinstance(request, NotificacionRequest):
            request_data = [request.to_dict()]  # Crear una lista con un solo diccionario
            enviar_notificacion(url_api, token, user, request_data)

        # Verificar si request es un diccionario
        elif isinstance(request, dict):
            request_data = [request]  # Crear una lista con un solo diccionario
            enviar_notificacion(url_api, token, user, request_data)

        else:
            raise ValueError("El argumento 'request' debe ser una instancia de NotificacionRequest, una lista de NotificacionRequest, o un diccionario.")

    except requests.exceptions.RequestException as e:
        print(f"Error en la solicitud API: {e}")
    except ValueError as e:
        print(f"Error de valor: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

def enviar_notificacion(url_api: str, token: str, user: str, request_data: List[dict]):
    """Envía la solicitud POST a la API con una lista de notificaciones."""
    try:

        headers = {
            "Content-Type": "application/json",  
            "Authorization": f"Bearer {token}",
            "Usarname": user,
        }

        url = f"{url_api}/api/FacturaProveedor/NotificarError"
        respuesta = requests.post(url, json=request_data, headers=headers)

        if respuesta.status_code == 200:
            print("Las notificaciones se registraron correctamente.")
        else:
            print(f"Error al notificar errores. Código de estado: {respuesta.status_code}")
            try:
                detalles = respuesta.json()
                print(f"Detalles del error: {detalles}")
            except ValueError:
                print(f"Detalles del error: {respuesta.text}")

    except requests.exceptions.RequestException as e:
        print(f"Error en la solicitud API: {e}")