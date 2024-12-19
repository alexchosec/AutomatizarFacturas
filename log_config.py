import logging
import os
from datetime import datetime

def setup_logging():
    # Verificar si la carpeta de logs existe, y si no, crearla
    log_dir = r'C:\LOG'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # Limpiar los handlers anteriores
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    # Obtener la fecha actual en formato YYYY-MM-DD
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Crear un handler para escribir en el archivo de log con la fecha actual
    log_filename = os.path.join(log_dir, f'logfile-{current_date}.log')
    log_handler = logging.FileHandler(log_filename, mode='a')  # 'a' es para append
    log_handler.setLevel(logging.INFO)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    log_handler.setFormatter(formatter)

    # Crear un handler para la consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # Configurar el logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)  # Nivel de log
    logger.addHandler(log_handler)  # Agregar el handler para el archivo
    logger.addHandler(console_handler)  # Agregar el handler para la consola

    return logger