import pypandoc
import traceback
from datetime import datetime
import os

LOG_FILE = os.path.join(os.path.dirname(__file__), "ajustes", "log.txt")

def escribir_log(msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")

def generar_docx_func(html_path, docx_path):
    try:
        escribir_log(f"Iniciando conversión HTML→DOCX desde {html_path}")
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        pypandoc.convert_text(
            html,
            "docx",
            format="html",
            outputfile=docx_path,
            extra_args=['--standalone']
        )
        escribir_log(f"Conversión completada correctamente → {docx_path}")
    except Exception as e:
        trace = traceback.format_exc()
        escribir_log(f"❌ ERROR durante conversión: {e}\n{trace}")
        raise
