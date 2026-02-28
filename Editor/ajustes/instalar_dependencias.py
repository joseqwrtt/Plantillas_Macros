import subprocess
import sys
import os
from datetime import datetime

LOG_FILE = os.path.join(os.path.dirname(__file__), "log.txt")

def log(msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")

def instalar():
    paquetes = ["flask", "pypandoc", "python-docx"]
    for pkg in paquetes:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
            log(f"Instalado correctamente: {pkg}")
        except Exception as e:
            log(f"❌ Error al instalar {pkg}: {e}")

if __name__ == "__main__":
    log("=== Iniciando instalación de dependencias ===")
    instalar()
    log("=== Dependencias instaladas correctamente ===")
