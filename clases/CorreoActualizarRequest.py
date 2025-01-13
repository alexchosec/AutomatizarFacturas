class CorreoActualizarRequest:
    def __init__(self, id: int):
        self.id = id
        
    def to_dict(self):
        return {
            "id": self.id
        }
