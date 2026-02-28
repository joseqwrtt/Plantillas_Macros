from flask import Flask, render_template, request, jsonify
import os
import threading
from datetime import datetime
import traceback
import re
from generar_docx import generar_docx_func  # ahora está en la raíz de Editor

# === RUTAS ===
BASE_DIR = os.path.dirname(__file__)
AJUSTES_DIR = os.path.join(BASE_DIR, "ajustes")
TEMP_DIR = os.path.join(AJUSTES_DIR, "temp")
PLANTILLAS_DIR = os.path.join(os.path.dirname(BASE_DIR), "Plantillas")
LOG_FILE = os.path.join(AJUSTES_DIR, "log.txt")

for d in [AJUSTES_DIR, TEMP_DIR, PLANTILLAS_DIR]:
    os.makedirs(d, exist_ok=True)

# === LOG GLOBAL ===
def escribir_log(mensaje):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {mensaje}\n")

# === FLASK APP ===
app = Flask(__name__, template_folder='.')

@app.route('/')
def index():
    escribir_log("Página principal cargada (editor.html)")
    return render_template('editor.html')

@app.route('/guardar', methods=['POST'])
def guardar():
    try:
        escribir_log("Petición /guardar recibida")

        data = request.get_json(force=True)
        contenido = data.get('html', '')
        titulo = data.get('titulo', '').strip()
        escribir_log(f"Datos recibidos → Título: {titulo}, Tamaño HTML: {len(contenido)} caracteres")

        if not titulo:
            escribir_log("Error: no se proporcionó título")
            return jsonify({"success": False, "msg": "Debes introducir un título"}), 400

        # Sanitizar nombre de archivo
        titulo_seguro = re.sub(r'[\\/*?:"<>|]', "_", titulo)
        html_path = os.path.join(TEMP_DIR, "content_temp.html")

        with open(html_path, "w", encoding="utf-8") as f:
            f.write(contenido)
        escribir_log(f"HTML temporal guardado en: {html_path}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        docx_name = f"{titulo_seguro}.docx"
        docx_path = os.path.join(PLANTILLAS_DIR, docx_name)

        generar_docx_func(html_path, docx_path)
        escribir_log(f"DOCX generado exitosamente: {docx_path}")

        return jsonify({"success": True, "msg": f"Plantilla generada: {docx_name}"})
    except Exception as e:
        trace = traceback.format_exc()
        escribir_log(f"❌ ERROR en /guardar: {e}\n{trace}")
        return jsonify({"success": False, "msg": "Error interno al guardar. Revisa log.txt."}), 500

@app.route('/cerrar', methods=['POST'])
def cerrar():
    escribir_log("Solicitud de cierre recibida (/cerrar)")
    def shutdown():
        func = request.environ.get('werkzeug.server.shutdown')
        if func:
            func()
            escribir_log("Servidor Flask cerrado correctamente")
    threading.Thread(target=shutdown).start()
    return jsonify({"success": True, "msg": "Servidor cerrado"})

if __name__ == '__main__':
    escribir_log("=== Servidor Flask iniciado ===")
    app.run(port=5000, debug=True)
