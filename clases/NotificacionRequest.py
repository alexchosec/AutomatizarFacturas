class NotificacionRequest:
    def __init__(self, asunto: str, mensaje: str):
        if not isinstance(asunto, str):
            raise TypeError("El asunto debe ser una cadena de texto.")
        if not isinstance(mensaje, str):
            raise TypeError("El mensaje debe ser una cadena de texto.")
        self.asunto = asunto
        self.mensaje = mensaje

    def __str__(self):
        return f"Asunto: {self.asunto}, Mensaje: {self.mensaje}"
 
    def to_dict(self):
        return {
            "asunto": self.asunto,
            "mensaje": self.mensaje
        }