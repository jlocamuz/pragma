"""
Archivo principal de la aplicación
Generador de Reportes de Asistencia - Humand.co
"""

import sys
import os

# Agregar el directorio src al path para imports absolutos
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importar la ventana principal desde el módulo ui
from ui.main_window import main

if __name__ == "__main__":
    main()
