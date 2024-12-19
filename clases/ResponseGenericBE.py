class ResponseGenericBE:
    def __init__(self, mensaje, respuesta, param1=None, param2=None):
        self.mensaje = mensaje
        self.respuesta = respuesta
        self.param1 = param1 if param1 is not None else ""  
        self.param2 = param2 if param2 is not None else ""  

    def __repr__(self):
        return f"ResponseGenericBE(mensaje={self.mensaje}, respuesta={self.respuesta}, param1={self.param1}, param2={self.param2})"
