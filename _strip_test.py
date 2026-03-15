"""
_strip_test.py — elimina el bloque TEST_FORZAR_MODO de licencia.py
Se llama desde build.bat antes de compilar.
Uso: python _strip_test.py ruta\licencia.py
"""
import sys, re

if len(sys.argv) < 2:
    print("Uso: python _strip_test.py <ruta_licencia.py>")
    sys.exit(1)

path = sys.argv[1]

try:
    content = open(path, encoding="utf-8").read()

    # 1. Eliminar el bloque de comentario + variable TEST_FORZAR_MODO
    content = re.sub(
        r"# ── MODO TEST ─+\n.*?TEST_FORZAR_MODO\s*=\s*None[^\n]*\n",
        "",
        content,
        flags=re.DOTALL
    )

    # 2. Eliminar el bloque if TEST_FORZAR_MODO en comprobar_licencia
    content = re.sub(
        r"\s*# ── Modo test.*?return ResultadoLicencia\(\"offline\".*?\)\n",
        "\n",
        content,
        flags=re.DOTALL
    )

    # 3. Por si quedó alguna referencia suelta
    content = re.sub(r"\s*if TEST_FORZAR_MODO is not None:.*?_iniciar_trial_si_no_existe\(\)",
                     "\n    _iniciar_trial_si_no_existe()",
                     content, flags=re.DOTALL)

    open(path, "w", encoding="utf-8").write(content)
    print(f"OK - TEST_FORZAR_MODO eliminado de {path}")
    sys.exit(0)

except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
