# main.py  ─ v3.0
# CORRECCIÓN CRÍTICA:
#   suppress=True en keyboard.hook bloqueaba TODA escritura globalmente.
#   Solución: dos capas separadas —
#     · Capa 1 (siempre): on_press sin supresión, gestiona el buffer.
#     · Capa 2 (solo popup visible): add_hotkey con suppress=True para
#       ↑ ↓ Enter Esc; se activan al mostrar el popup y se eliminan al cerrarlo.
import os
import sys
import traceback
import time
import threading
import json

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from licencia import comprobar_licencia, ResultadoLicencia
    from ventana_licencia import pedir_licencia
    _LICENCIA_DISPONIBLE = True
except ImportError:
    _LICENCIA_DISPONIBLE = False
    class ResultadoLicencia:
        modo="completo"; ok=True; es_completo=True
        detector_ok=True; max_docx=None; max_odt=None; max_total=None
        dias_trial=0

try:
    from core import GestorPlantillas
    from core_odt import GestorPlantillasODT
except ImportError as e:
    print(f"❌ Error al importar módulos: {e}")
    input("Presiona Enter para salir...")
    sys.exit(1)


def crear_preferencias_si_no_existen(ruta_ajustes, ruta_plantillas):
    preferencias_default = {
        "detonante": "/",
        "ruta_plantillas": ruta_plantillas,
        "debug": False,
        "privacidad_logs": True,
        "detector_activado": True,
        "tiempo_reinicio_buffer": 2,
        "ancho_minimo_popup": 15,
        "max_resultados_popup": 100,
        "mostrar_iconos": True,
        "pegar_automaticamente": True,
        "atajos_activados": True,
        "tema": "dark",
        "fuente_html": {
            "modo": "personalizado",
            "familia": "Arial",
            "tamaño": 10,
            "color": "#000000"
        }
    }
    pref_file = os.path.join(ruta_ajustes, "preferencias.json")
    if not os.path.exists(pref_file):
        with open(pref_file, "w", encoding="utf-8") as f:
            json.dump(preferencias_default, f, ensure_ascii=False, indent=4)
        print(f"✅ Creado archivo de preferencias: {pref_file}")
    else:
        print("📋 Archivo de preferencias ya existe")


# ══════════════════════════════════════════════════════════════
class GestorUnificado:
    """Unifica gestores DOCX + ODT con detector de dos capas."""

    def __init__(self):
        self.gestor_docx = GestorPlantillas()
        self.gestor_odt  = GestorPlantillasODT()
        self.preferencias = self.gestor_docx.preferencias.copy()
        self.gestor_odt.preferencias = self.preferencias.copy()
        ruta = self.preferencias.get("ruta_plantillas", "")
        self.gestor_docx.preferencias["ruta_plantillas"] = ruta
        self.gestor_odt.preferencias["ruta_plantillas"]  = ruta

        self.popup_activo      = None
        self.buffer_teclas     = []
        self.ultimo_tiempo     = time.time()
        self._detector_thread  = None
        self.interfaz          = None
        self._hotkeys_popup    = []   # hotkeys suprimidas activas
        self.licencia          = None
        print("✅ Gestor unificado iniciado (DOCX + ODT)")

    # ── Preferencias ──────────────────────────────────────────
    def cargar_preferencias(self):
        self.preferencias = self.gestor_docx.cargar_preferencias()
        self.gestor_odt.preferencias = self.preferencias.copy()
        return self.preferencias

    def guardar_preferencias(self, nuevas):
        try:
            if self.gestor_docx.guardar_preferencias(nuevas):
                self.preferencias = nuevas.copy()
                self.gestor_odt.preferencias = nuevas.copy()
                ruta = nuevas.get("ruta_plantillas", "")
                self.gestor_docx.preferencias["ruta_plantillas"] = ruta
                self.gestor_odt.preferencias["ruta_plantillas"]  = ruta
                if self.interfaz and hasattr(self.interfaz, "recargar_preferencias"):
                    self.interfaz.recargar_preferencias()
                return True
        except Exception as e:
            print(f"Error guardando preferencias: {e}")
        return False

    # ── Logging ───────────────────────────────────────────────
    def log(self, msg, nivel="debug", datos_sensibles=False):
        return self.gestor_docx.log(msg, nivel, datos_sensibles)

    def log_excepcion(self, e, ctx=""):
        return self.gestor_docx.log_excepcion(e, ctx)

    # ── Plantillas ────────────────────────────────────────────
    def procesar_plantilla(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".docx":
            return self.gestor_docx.procesar_plantilla(path)
        elif ext == ".odt":
            return self.gestor_odt.procesar_plantilla(path)
        print(f"❌ Formato no soportado: {ext}")
        return False

    def buscar_plantillas(self, texto_filtro):
        ruta = self.preferencias.get("ruta_plantillas", "")
        if not os.path.exists(ruta):
            ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Plantillas")
        resultados = []
        filtro_limpio = self.gestor_docx.eliminar_acentos(texto_filtro)
        max_res = self.preferencias.get("max_resultados_popup", 100)
        for raiz, _, archivos in os.walk(ruta):
            for f in archivos:
                ext = os.path.splitext(f)[1].lower()
                if ext in (".docx", ".odt"):
                    nombre = os.path.splitext(f)[0]
                    if filtro_limpio in self.gestor_docx.eliminar_acentos(nombre):
                        resultados.append((nombre, os.path.join(raiz, f)))
                        if len(resultados) >= max_res:
                            return sorted(resultados)
        return sorted(resultados)

    def eliminar_acentos(self, texto):
        return self.gestor_docx.eliminar_acentos(texto)

    def limpiar_buffer(self):
        self.buffer_teclas = []

    def set_popup_activo(self, popup):
        self.popup_activo = popup
        self.gestor_docx.set_popup_activo(popup)
        self.gestor_odt.set_popup_activo(popup)

    def get_popup_activo(self):
        return self.popup_activo

    # ── Hotkeys de navegación (Capa 2) ───────────────────────
    def activar_hotkeys_popup(self):
        """
        Registra ↑ ↓ Enter Esc con suppress=True.
        Llamado SOLO cuando el popup se hace visible.
        """
        import keyboard
        self.desactivar_hotkeys_popup()

        def _up():
            if self.popup_activo and self.popup_activo.esta_activo():
                self.popup_activo.mover_sel(-1)

        def _down():
            if self.popup_activo and self.popup_activo.esta_activo():
                self.popup_activo.mover_sel(1)

        def _enter():
            if self.popup_activo and self.popup_activo.esta_activo():
                self.popup_activo.confirmar()
                self.desactivar_hotkeys_popup()

        def _esc():
            if self.popup_activo and self.popup_activo.esta_activo():
                self.popup_activo.cerrar()
                self.buffer_teclas = []
                self.desactivar_hotkeys_popup()

        try:
            self._hotkeys_popup = [
                keyboard.add_hotkey("up",    _up,    suppress=True, trigger_on_release=False),
                keyboard.add_hotkey("down",  _down,  suppress=True, trigger_on_release=False),
                keyboard.add_hotkey("enter", _enter, suppress=True, trigger_on_release=False),
                keyboard.add_hotkey("esc",   _esc,   suppress=True, trigger_on_release=False),
            ]
            print("🔒 Hotkeys popup activados (↑↓ Enter Esc bloqueados)")
        except Exception as ex:
            print(f"⚠️ Error activando hotkeys: {ex}")

    def desactivar_hotkeys_popup(self):
        """Elimina los hotkeys de navegación."""
        import keyboard
        for hk in self._hotkeys_popup:
            try:
                keyboard.remove_hotkey(hk)
            except Exception:
                pass
        self._hotkeys_popup = []

    # ── Detector principal (Capa 1) ───────────────────────────
    def iniciar_detector(self):
        if not self.preferencias.get("detector_activado", True):
            print("Detector desactivado.")
            return

        def _detector():
            import keyboard

            def on_key(event):
                if not self.preferencias.get("detector_activado", True):
                    return

                detonante  = self.preferencias.get("detonante", "/")
                t_reinicio = self.preferencias.get("tiempo_reinicio_buffer", 2)
                nombre     = event.name

                popup = self.popup_activo
                popup_visible = (popup is not None
                                 and hasattr(popup, "esta_activo")
                                 and popup.esta_activo())

                # ── Con popup visible: solo actualizar filtro ──
                if popup_visible:
                    if nombre == "backspace":
                        if self.buffer_teclas:
                            self.buffer_teclas.pop()
                        if not self.buffer_teclas:
                            popup.cerrar()
                            self.desactivar_hotkeys_popup()
                        elif self.interfaz:
                            filtro = "".join(self.buffer_teclas)[len(detonante):]
                            self.interfaz.actualizar_popup_comandos(filtro)
                    elif len(nombre) == 1 or nombre == "space":
                        self.buffer_teclas.append(nombre if nombre != "space" else " ")
                        if self.interfaz:
                            filtro = "".join(self.buffer_teclas)[len(detonante):]
                            self.interfaz.actualizar_popup_comandos(filtro)
                    return   # ← dejar pasar TODAS las teclas normales

                # ── Sin popup: gestión del buffer ──
                # El detonante puede tener 1 o más caracteres (ej: "/" o "//").
                # Acumulamos cada tecla y comprobamos si el buffer TERMINA con
                # el detonante completo para activar el modo comando.

                if time.time() - self.ultimo_tiempo > t_reinicio:
                    self.buffer_teclas = []
                self.ultimo_tiempo = time.time()

                if nombre == "backspace":
                    if self.buffer_teclas:
                        self.buffer_teclas.pop()
                    return

                if len(nombre) == 1 or nombre == "space":
                    self.buffer_teclas.append(nombre if nombre != "space" else " ")
                    buffer_actual = "".join(self.buffer_teclas)

                    # El buffer acaba de completar el detonante → reiniciar en modo comando
                    if buffer_actual.endswith(detonante) and len(buffer_actual) == len(detonante):
                        # Exactamente el detonante, nada más: listo para recibir filtro
                        self.buffer_teclas = list(detonante)
                        return

                    # Ya tenemos detonante + algo de filtro → mostrar popup
                    if (buffer_actual.startswith(detonante)
                            and len(buffer_actual) > len(detonante)):
                        filtro = buffer_actual[len(detonante):]
                        try:
                            import ctypes
                            import ctypes.wintypes as wt
                            pt = wt.POINT()
                            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                            x, y = pt.x, pt.y + 22
                        except Exception:
                            x, y = None, None
                        if self.interfaz:
                            self.interfaz.mostrar_popup_desde_detector(filtro, x, y)
                    elif not buffer_actual.startswith(detonante):
                        # Si los primeros caracteres no van camino del detonante, limpiar
                        # Pero si el buffer podría ser el inicio del detonante, conservar
                        if not detonante.startswith(buffer_actual):
                            self.buffer_teclas = []

            # ← suppress=False: NO bloquea NADA — solo escucha
            keyboard.on_press(on_key)
            keyboard.wait()

        self._detector_thread = threading.Thread(target=_detector, daemon=True)
        self._detector_thread.start()
        print(f"🎯 Detector '{self.preferencias.get('detonante', '/')}' activado")


# ══════════════════════════════════════════════════════════════
def main():
    try:
        if getattr(sys, "frozen", False):
            current_dir = os.path.dirname(sys.executable)
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))

        # Plantillas: junto al exe (fácil de gestionar por el usuario)
        ruta_plantillas = os.path.join(current_dir, "Plantillas")

        # Ajustes: en AppData del usuario (siempre tiene permisos, estándar Windows)
        # Fallback al directorio del exe si AppData no está disponible
        appdata = os.environ.get("APPDATA", "")
        if appdata and os.path.isdir(appdata):
            ruta_ajustes = os.path.join(appdata, "PlantillasMacro", "Ajustes")
        else:
            ruta_ajustes = os.path.join(current_dir, "Ajustes")

        os.makedirs(ruta_plantillas, exist_ok=True)
        os.makedirs(ruta_ajustes,    exist_ok=True)
        print(f"📁 Ajustes en: {ruta_ajustes}")

        crear_preferencias_si_no_existen(ruta_ajustes, ruta_plantillas)

        import core, core_odt
        core.AJUSTES_DIR        = ruta_ajustes
        core.PLANTILLAS_DIR     = ruta_plantillas
        core.CONFIG_FILE        = os.path.join(ruta_ajustes, "config.json")
        core.PREFERENCIAS_FILE  = os.path.join(ruta_ajustes, "preferencias.json")
        core_odt.AJUSTES_DIR       = ruta_ajustes
        core_odt.PLANTILLAS_DIR    = ruta_plantillas
        core_odt.CONFIG_FILE       = os.path.join(ruta_ajustes, "config.json")
        core_odt.PREFERENCIAS_FILE = os.path.join(ruta_ajustes, "preferencias.json")

        gestor = GestorUnificado()

        if not gestor.preferencias.get("ruta_plantillas"):
            for g in (gestor, gestor.gestor_docx, gestor.gestor_odt):
                g.preferencias["ruta_plantillas"] = ruta_plantillas

        print(f"✅ Iniciando desde: {current_dir}")
        print(f"📁 Plantillas: {gestor.preferencias.get('ruta_plantillas')}")

        # ── Comprobación de licencia ──────────────────────────────────────
        if _LICENCIA_DISPONIBLE:
            lic = comprobar_licencia()
            gestor.licencia = lic
            print(f"📋 Licencia: {lic.modo} — {lic.mensaje}")

            if lic.modo == "limitado":
                pedir_licencia(lic)
                lic = comprobar_licencia()
                gestor.licencia = lic
        else:
            lic = ResultadoLicencia()
            gestor.licencia = lic

        from interfaz_ctk import InterfazCTK
        interfaz = InterfazCTK()

        # Detector solo si la licencia lo permite
        detector_pref      = gestor.preferencias.get("detector_activado", True)
        detector_permitido = getattr(gestor.licencia, "detector_ok", True)
        if detector_pref and detector_permitido:
            try:
                gestor.iniciar_detector()
            except Exception as e:
                print(f"⚠️  Detector no pudo iniciarse: {e}")
                gestor.detector_bloqueado = True
        elif detector_pref and not detector_permitido:
            print("⚠️  Detector desactivado — versión limitada")

        interfaz.ejecutar(gestor)

    except Exception:
        print("🔴 ERROR CRÍTICO AL INICIAR:")
        traceback.print_exc()
        input("\nPresiona Enter para cerrar...")


if __name__ == "__main__":
    main()