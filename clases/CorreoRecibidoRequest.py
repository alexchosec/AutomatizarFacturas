class CorreoRecibidoRequest:
    def __init__(self, remitente: str, asunto: str, cuerpoMensaje: str, usuario: str):
        self.remitente = remitente
        self.asunto = asunto
        self.cuerpoMensaje = cuerpoMensaje
        self.usuario = usuario
        
    def to_dict(self):
        return {
            "remitente": self.remitente,
            "asunto": self.asunto,
            "cuerpoMensaje": self.cuerpoMensaje,
            "usuario": self.usuario,
        }
