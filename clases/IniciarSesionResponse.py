from datetime import datetime

class IniciarSesionResponse:
    def __init__(self, expires=None, username=None, token=None):
        # Asegúrate de manejar 'expires' si es None o una cadena
        if isinstance(expires, str):
            self.expires = datetime.strptime(expires, "%Y-%m-%dT%H:%M:%S")
        else:
            self.expires = expires  # Puede ser None o un objeto datetime
        
        # Manejo de username y token como opcionales
        self.username = username if username is not None else ""
        self.token = token if token is not None else ""

    def __repr__(self):
        # Para facilitar la impresión de la instancia
        return f"IniciarSesionResponse(expires={self.expires}, username={self.username}, token={self.token})"
