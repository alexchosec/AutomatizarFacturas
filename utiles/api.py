import requests
import json
import os

from clases.IniciarSesionResponse import IniciarSesionResponse
from clases.ResponseGenericBE import ResponseGenericBE

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



def upload_xml(url_api, token, user, file_path):
    response = None
    try:
        headers = {
            "Usuario": user,
            "Authorization": f"Bearer {token}"
        }

        with open(file_path, 'rb') as file:

            files = {'file': (os.path.basename(file_path), file, 'application/octet-stream')}
            
            respuesta = requests.post(f"{url_api}/api/Comprobante/SubirXml", headers=headers, files=files)
            
            if respuesta.status_code == 200:

                data = respuesta.json()
                
                response = ResponseGenericBE(
                    respuesta=data['respuesta'],
                    mensaje=data['mensaje'],
                    param1=data['param1'],
                    param2=data['param2']
                )
                
            else:
                response = ResponseGenericBE(f"Error al subir el archivo. Código de respuesta: {respuesta.status_code}", False)

    except Exception as ex:        
        response = ResponseGenericBE(f"Ocurrió un error al intentar subir el archivo: {ex}", False)
 
    return response



def upload_file(url_api, token, user, id, isZip, file_path):
    response = None
    try:
        
        headers = {
            "Usuario": user,
            "DocumentoID": f"{id}",
            "EstaComprimido": isZip,
            "Authorization": f"Bearer {token}"
        }

        with open(file_path, 'rb') as file:
            
            files = {'file': (os.path.basename(file_path), file, 'application/octet-stream')}
            
            respuesta = requests.post(f"{url_api}/api/Comprobante/SubirArchivo", headers=headers, files=files)
            
            if respuesta.status_code == 200:

                data = respuesta.json()
                
                response = ResponseGenericBE(
                    respuesta=data['respuesta'],
                    mensaje=data['mensaje'],
                    param1=data['param1'],
                    param2=data['param2']
                )
                
            else:
                response = ResponseGenericBE(f"Error al subir el archivo. Código de respuesta: {respuesta.status_code}", False)

    except Exception as ex:        
        response = ResponseGenericBE(f"Ocurrió un error al intentar subir el archivo: {ex}", False)
 
    return response
