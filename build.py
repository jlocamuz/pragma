"""
Script para compilar la aplicación en un ejecutable
Usa PyInstaller para crear un archivo .exe/.app/.bin
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def build_executable():
    """Compila la aplicación en un ejecutable"""
    
    print("🚀 Iniciando compilación de la aplicación...")
    
    # Verificar que PyInstaller esté instalado
    try:
        import PyInstaller
        print(f"✅ PyInstaller encontrado: {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller no está instalado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Configuración de PyInstaller
    app_name = "Generador_Reportes_Asistencia"
    main_script = "src/main.py"
    
    # Argumentos para PyInstaller
    pyinstaller_args = [
        "pyinstaller",
        "--onefile",  # Un solo archivo ejecutable
        "--windowed",  # Sin ventana de consola (solo en Windows)
        "--name", app_name,
        "--paths", "src",  # Agregar src al path de Python
        "--hidden-import", "PyQt5.QtCore",
        "--hidden-import", "PyQt5.QtWidgets",
        "--hidden-import", "PyQt5.QtGui",
        "--hidden-import", "openpyxl",
        "--hidden-import", "requests",
        "--hidden-import", "ui",
        "--hidden-import", "ui.main_window",
        "--hidden-import", "core",
        "--hidden-import", "core.api_client",
        "--hidden-import", "core.data_processor",
        "--hidden-import", "core.excel_generator",
        "--hidden-import", "core.hours_calculator",
        "--hidden-import", "config",
        "--hidden-import", "config.default_config",
        "--clean",  # Limpiar cache antes de compilar
        main_script
    ]
    
    # Agregar icono solo si existe
    if os.path.exists("src/resources/icon.ico"):
        pyinstaller_args.insert(-1, "--icon")
        pyinstaller_args.insert(-1, "src/resources/icon.ico")
    
    print(f"📦 Compilando {app_name}...")
    print(f"📄 Script principal: {main_script}")
    
    try:
        # Ejecutar PyInstaller
        result = subprocess.run(pyinstaller_args, check=True, capture_output=True, text=True)
        print("✅ Compilación exitosa!")
        
        # Mostrar información del ejecutable
        dist_dir = Path("dist")
        if dist_dir.exists():
            executables = list(dist_dir.glob("*"))
            if executables:
                exe_path = executables[0]
                exe_size = exe_path.stat().st_size / (1024 * 1024)  # MB
                print(f"📁 Ejecutable creado: {exe_path}")
                print(f"📏 Tamaño: {exe_size:.1f} MB")
                
                # Instrucciones de uso
                print("\n🎉 ¡Compilación completada!")
                print(f"📂 El ejecutable está en: {exe_path.absolute()}")
                print("\n📋 Instrucciones:")
                print("1. Puedes mover el ejecutable a cualquier carpeta")
                print("2. No necesita instalación adicional")
                print("3. Ejecuta directamente haciendo doble clic")
                print("4. Los reportes se guardarán en ~/Downloads")
                
                return str(exe_path.absolute())
            else:
                print("❌ No se encontró el ejecutable en la carpeta dist/")
                return None
        else:
            print("❌ No se creó la carpeta dist/")
            return None
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante la compilación:")
        print(f"Código de salida: {e.returncode}")
        print(f"Error: {e.stderr}")
        return None
    except Exception as e:
        print(f"❌ Error inesperado: {str(e)}")
        return None

def clean_build_files():
    """Limpia archivos temporales de compilación"""
    print("🧹 Limpiando archivos temporales...")
    
    dirs_to_clean = ["build", "__pycache__"]
    files_to_clean = ["*.spec"]
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"🗑️ Eliminado: {dir_name}/")
    
    # Limpiar archivos .spec
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"🗑️ Eliminado: {spec_file}")
    
    print("✅ Limpieza completada")

def main():
    """Función principal"""
    print("=" * 60)
    print("📦 COMPILADOR DE APLICACIÓN DESKTOP")
    print("   Generador de Reportes de Asistencia")
    print("=" * 60)
    
    # Verificar que estamos en el directorio correcto
    if not os.path.exists("src/main.py"):
        print("❌ Error: No se encontró src/main.py")
        print("   Asegúrate de ejecutar este script desde la carpeta raíz del proyecto")
        return
    
    # Verificar dependencias
    print("🔍 Verificando dependencias...")
    try:
        import PyQt5
        import openpyxl
        import requests
        print("✅ Todas las dependencias están instaladas")
    except ImportError as e:
        print(f"❌ Dependencia faltante: {e}")
        print("   Ejecuta: pip install -r requirements.txt")
        return
    
    # Preguntar si limpiar archivos anteriores
    if os.path.exists("dist") or os.path.exists("build"):
        response = input("🤔 ¿Limpiar compilaciones anteriores? (s/N): ").lower()
        if response in ['s', 'si', 'sí', 'y', 'yes']:
            clean_build_files()
    
    # Compilar
    exe_path = build_executable()
    
    if exe_path:
        print("\n" + "=" * 60)
        print("🎉 ¡COMPILACIÓN EXITOSA!")
        print(f"📂 Ejecutable: {exe_path}")
        print("=" * 60)
        
        # Preguntar si probar el ejecutable
        response = input("🚀 ¿Probar el ejecutable ahora? (s/N): ").lower()
        if response in ['s', 'si', 'sí', 'y', 'yes']:
            try:
                if sys.platform == "win32":
                    os.startfile(exe_path)
                elif sys.platform == "darwin":  # macOS
                    os.system(f"open '{exe_path}'")
                else:  # Linux
                    os.system(f"'{exe_path}' &")
                print("🚀 Ejecutable iniciado!")
            except Exception as e:
                print(f"❌ Error al iniciar ejecutable: {e}")
    else:
        print("\n" + "=" * 60)
        print("❌ COMPILACIÓN FALLIDA")
        print("   Revisa los errores anteriores")
        print("=" * 60)

if __name__ == "__main__":
    main()
