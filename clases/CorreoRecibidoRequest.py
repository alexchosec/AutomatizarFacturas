class CorreoRecibidoRequest:
    def __init__(self, remitente: str, asunto: str, usuario: str):
        self.remitente = remitente
        self.asunto = asunto
        self.usuario = usuario
        
    def to_dict(self):
        return {
            "remitente": self.remitente,
            "asunto": self.asunto,
            "usuario": self.usuario,
        }
