# licencia.py  ─ v2.0
# Modos:
#   "trial"    → primeros 14 días, acceso completo (detector incluido)
#   "limitado" → trial expirado sin licencia (12 plantillas, sin detector)
#   "completo" → licencia válida activada, todo ilimitado
#   "offline"  → licencia en caché, sin internet (hasta DIAS_GRACIA_RED días)

import os, sys, json, hashlib, hashlib, base64, datetime, platform, uuid

try:
    from cryptography.fernet import Fernet
    _CRYPTO_OK = True
except ImportError:
    _CRYPTO_OK = False

# ══════════════════════════════════════════════════
#  CONFIGURACIÓN ← rellena con tus datos
# ══════════════════════════════════════════════════
LS_API_KEY    = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiI5NGQ1OWNlZi1kYmI4LTRlYTUtYjE3OC1kMjU0MGZjZDY5MTkiLCJqdGkiOiIwMjA5NzBlOTNmZjQyOTVmYzlmZjY0MzYzZGM0NjViNDk5NDNkMjlmODY4NjI0YjY4NDRkZjgwNTcyN2FhMjU4NzZjODgxYWNmNTdkOGE1OCIsImlhdCI6MTc3MzUyOTc1Ni43OTU3NzMsIm5iZiI6MTc3MzUyOTc1Ni43OTU3NzUsImV4cCI6MjIzMjIzMDQwMC4wMjcxMzMsInN1YiI6IjY3MDQ2MDEiLCJzY29wZXMiOltdfQ.QrnTuEjF5saWZDVNrZKgtm8IL6hqoob0OFVETdMNDX6uU8RvcjSYQszf4yJ90jqiTBgIfrBuCJwjekZbQ7ZhksUdCOjc7hNhj2pHGY2tJO1Cr2ybQl0k6cW7tzjr-qQ84Eu9x-eFnwd7qS5-L4NG9GlzTv1DE7IDu5K5SEF3cDN4ftNcStqdB8W8BmlW228gIPJ8GdzJ6H4dK86uCd5TYfeSM6032tLoBVu3eXcGa-yjEKpw1kS3CwGoNtqvoQD5t_8z9PmsvjdQNv0s6hCfqY7qR9ZH5vMgRHOvqvShoXMKLXiJmYsLi9c9pQDSxROuGJP9OrovM7HzK2Cbdpb2JK8HwO6KCkbqqBJz4DZK2eP0NLGeb3LqhPPG0_qTz5zk-1bu4ztJc6crmtUc-iRJWDcx_TUnUoVkbQEFo-K9b9XHZ1Y8RiCW4TT7uVbKae1_QK1WDKhoL3IB4DLfnTiy3eNZGxkaXfCFzgi1NfVaZTMh-QPAPiI3sAnyVtuXFf_-j9Qa__IJhKtv4GuvEvg5rc-GvfN_lQu6eJNBQqGDK0Sh2hBJ4fzP6b7ZEuyx1NK2onlbG-GWNbN5fYmYLneb_Nf91R9Ulu2WRSFlEe45BgtXkTI90TbNoY_va6uDhPvrmaViVU3vUoCnOnASlkPbOk1vtb4uVaG_ZZGYd60FvTU"
LS_STORE_ID   = "315972"
LS_PRODUCT_ID = "TU_PRODUCT_ID_AQUI"
URL_TIENDA    = "https://srjosef9.lemonsqueezy.com"

DIAS_TRIAL        = 14
DIAS_GRACIA_RED   = 7

# ── MODO TEST ────────────────────────────────────
# Cambia a "limitado", "trial" o "completo" para simular
# Pon None para comportamiento real
TEST_FORZAR_MODO  = None   # ← cambia aquí para probar

# Límites versión gratuita (tras expirar trial)
LIMITE_DOCX       = 6
LIMITE_ODT        = 6
LIMITE_TOTAL      = 12
LIMITE_DETECTOR   = False   # sin detector global en versión limitada
# ══════════════════════════════════════════════════


# ── Rutas ────────────────────────────────────────
def _ruta_ajustes() -> str:
    """Misma lógica que main.py — AppData si está disponible, sino junto al exe."""
    appdata = os.environ.get("APPDATA", "")
    if appdata and os.path.isdir(appdata):
        ruta = os.path.join(appdata, "PlantillasMacro", "Ajustes")
    else:
        base = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                               else os.path.abspath(__file__))
        ruta = os.path.join(base, "Ajustes")
    os.makedirs(ruta, exist_ok=True)
    return ruta

def _cache_file() -> str:
    return os.path.join(_ruta_ajustes(), "licencia.dat")

def _trial_file() -> str:
    return os.path.join(_ruta_ajustes(), "trial.dat")


# ── Huella de dispositivo ────────────────────────
def _huella() -> str:
    raw = "|".join([platform.node(), platform.machine(), str(uuid.getnode())])
    return hashlib.sha256(raw.encode()).hexdigest()[:32]


# ── Cifrado ──────────────────────────────────────
def _fkey() -> bytes:
    return base64.urlsafe_b64encode(hashlib.sha256(_huella().encode()).digest())

def _cifrar(datos: dict) -> bytes:
    raw = json.dumps(datos, ensure_ascii=False).encode()
    if _CRYPTO_OK:
        return Fernet(_fkey()).encrypt(raw)
    k = _fkey()
    return base64.b64encode(bytes(b ^ k[i % len(k)] for i, b in enumerate(raw)))

def _descifrar(enc: bytes) -> dict | None:
    try:
        if _CRYPTO_OK:
            raw = Fernet(_fkey()).decrypt(enc)
        else:
            k     = _fkey()
            xored = base64.b64decode(enc)
            raw   = bytes(b ^ k[i % len(k)] for i, b in enumerate(xored))
        return json.loads(raw.decode())
    except Exception:
        return None

def _leer(path: str) -> dict | None:
    if not os.path.exists(path):
        return None
    try:
        return _descifrar(open(path, "rb").read())
    except Exception:
        return None

def _guardar(path: str, datos: dict):
    open(path, "wb").write(_cifrar(datos))

def _borrar(path: str):
    if os.path.exists(path):
        os.remove(path)


# ── Trial ────────────────────────────────────────
def _firma_trial(inicio_iso: str) -> str:
    """HMAC de la fecha de inicio — detecta si se modificó el archivo."""
    msg = (inicio_iso + _huella()).encode()
    return hashlib.sha256(msg).hexdigest()

# ── Registro de Windows como respaldo anti-borrado ───────────────────────
# Ruta camuflada como configuración de fuentes de Microsoft Office
# Cuatro ubicaciones de respaldo camufladas como entradas legítimas
# de software conocido. Ninguna de estas claves/valores existe realmente
# en Windows ni en las apps mencionadas — no causan conflicto.
#
#  Ruta 1: Office UserInfo  → metadatos de plantilla reciente
#  Ruta 2: .NET Framework   → config de rendimiento de app
#  Ruta 3: Windows Shell    → preferencias de carpetas del usuario
#  Ruta 4: Visual C++ Redist → versión de runtime instalada
#
_REGISTROS = [
    # (ruta, clave_inicio, clave_dia)
    (r"Software\Microsoft\Office\16.0\Common\UserInfo",
     "RecentTemplateTS", "RecentTemplateID"),

    (r"Software\Microsoft\.NETFramework\Policy\AppPerf",
     "InitStamp", "PerfIndex"),

    (r"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\Config",
     "SessionToken", "SessionSeq"),

    (r"Software\Microsoft\VisualStudio\14.0\Setup\VS",
     "BuildTimestamp", "BuildRevision"),
]

def _reg_leer() -> str | None:
    """Lee la fecha de inicio del trial desde el primer registro disponible."""
    try:
        import winreg
        for ruta, k_ini, _ in _REGISTROS:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, ruta)
                val, _ = winreg.QueryValueEx(key, k_ini)
                winreg.CloseKey(key)
                if val:
                    return val
            except Exception:
                continue
    except Exception:
        pass
    return None

def _reg_escribir(valor: str):
    """Guarda la fecha de inicio en todos los registros de respaldo."""
    try:
        import winreg
        for ruta, k_ini, _ in _REGISTROS:
            try:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, ruta)
                winreg.SetValueEx(key, k_ini, 0, winreg.REG_SZ, valor)
                winreg.CloseKey(key)
            except Exception:
                continue
    except Exception:
        pass

def _reg_leer_dia() -> int | None:
    """Lee el ultimo_dia desde el primer registro disponible."""
    try:
        import winreg
        for ruta, _, k_dia in _REGISTROS:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, ruta)
                val, _ = winreg.QueryValueEx(key, k_dia)
                winreg.CloseKey(key)
                return int(val)
            except Exception:
                continue
    except Exception:
        pass
    return None

def _reg_escribir_dia(dia: int):
    """Guarda el ultimo_dia en todos los registros de respaldo."""
    try:
        import winreg
        for ruta, _, k_dia in _REGISTROS:
            try:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, ruta)
                winreg.SetValueEx(key, k_dia, 0, winreg.REG_SZ, str(dia))
                winreg.CloseKey(key)
            except Exception:
                continue
    except Exception:
        pass

def _iniciar_trial_si_no_existe():
    path      = _trial_file()
    reg_valor = _reg_leer()   # marca en registro

    # ── Caso 1: archivo existe → normal ──────────────────────────────────
    if os.path.exists(path):
        # Si el registro no tiene nada (borrado manualmente), restaurarlo
        if reg_valor is None:
            datos = _leer(path)
            if datos:
                _reg_escribir(datos.get("inicio", ""))
        return

    # ── Caso 2: archivo NO existe pero registro SÍ → fue borrado ─────────
    if reg_valor is not None:
        ahora_orig = reg_valor
        try:
            datetime.datetime.fromisoformat(ahora_orig)
        except Exception:
            ahora_orig = datetime.datetime.now().isoformat()

        # Recuperar ultimo_dia del registro; si no está usar hoy
        # pero nunca menos que los días transcurridos reales
        dia_reg = _reg_leer_dia()
        hoy_ord = datetime.datetime.now().toordinal()
        # El ultimo_dia recuperado debe ser al menos el día actual
        # (para no perder días transcurridos)
        ultimo_recuperado = max(dia_reg or hoy_ord, hoy_ord)

        _guardar(path, {
            "inicio":      ahora_orig,
            "dispositivo": _huella(),
            "firma":       _firma_trial(ahora_orig),
            "ultimo_dia":  ultimo_recuperado,
        })
        return

    # ── Caso 3: ni archivo ni registro → primera vez real ─────────────────
    ahora = datetime.datetime.now().isoformat()
    _guardar(path, {
        "inicio":      ahora,
        "dispositivo": _huella(),
        "firma":       _firma_trial(ahora),
        "ultimo_dia":  datetime.datetime.now().toordinal(),
    })
    _reg_escribir(ahora)   # guardar también en registro

def dias_trial_restantes() -> int:
    datos = _leer(_trial_file())

    # Sin datos o dispositivo distinto → expirado
    if not datos or datos.get("dispositivo") != _huella():
        return 0

    try:
        inicio_iso = datos.get("inicio", "")
        inicio = datetime.datetime.fromisoformat(inicio_iso)

        # ── Anti-manipulación 1: firma HMAC ──────────────────────────────
        # Si alguien edita trial.dat para cambiar la fecha, la firma no cuadra
        firma_esperada = _firma_trial(inicio_iso)
        if datos.get("firma") != firma_esperada:
            return 0   # archivo manipulado → expirado

        # ── Anti-manipulación 2: fecha hacia atrás ───────────────────────
        # Si el usuario retrasó el reloj del sistema, el día actual
        # será MENOR que el último día registrado → manipulación detectada
        hoy = datetime.datetime.now().toordinal()
        ultimo = datos.get("ultimo_dia", hoy)
        if hoy < ultimo - 1:   # tolerancia de 1 día por zonas horarias
            return 0   # reloj manipulado → expirado

        # Actualizar "ultimo_dia" si avanzó (guardar el mayor visto)
        if hoy > ultimo:
            datos["ultimo_dia"] = hoy
            _guardar(_trial_file(), datos)

        # ── Cálculo normal ───────────────────────────────────────────────
        transcurridos = (datetime.datetime.now() - inicio).days
        return max(0, DIAS_TRIAL - transcurridos)

    except Exception:
        return 0

def trial_activo() -> bool:
    return dias_trial_restantes() > 0


# ── Validación online LemonSqueezy ───────────────
def validar_online(clave: str) -> dict:
    """Devuelve dict: ok, mensaje, instancia, datos"""
    try:
        import urllib.request, urllib.parse

        payload = urllib.parse.urlencode({
            "license_key":   clave.strip().upper(),
            "instance_name": platform.node(),
        }).encode()

        req = urllib.request.Request(
            "https://api.lemonsqueezy.com/v1/licenses/validate",
            data    = payload,
            method  = "POST",
            headers = {
                "Accept":       "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
            }
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            body = json.loads(r.read().decode())

        valida    = body.get("valid", False)
        lk        = body.get("license_key", {})
        estado    = lk.get("status", "")
        instancia = body.get("instance", {}).get("id", "")

        if not valida:
            msg = {
                "inactive": "La licencia está desactivada.",
                "expired":  "La licencia ha expirado.",
                "disabled": "La licencia ha sido deshabilitada.",
            }.get(estado, f"Licencia no válida (estado: {estado})")
            return {"ok": False, "mensaje": msg, "instancia": "", "datos": body}

        # Verificar que es de 1 dispositivo: si ya hay instancia y no es la nuestra
        instancias = body.get("license_key", {}).get("activation_usage", 0)
        limite     = body.get("license_key", {}).get("activation_limit", 1)
        if instancias >= limite and not instancia:
            return {"ok": False,
                    "mensaje": "Esta licencia ya está activada en otro dispositivo.\n"
                               "Desactívala primero desde el otro equipo.",
                    "instancia": "", "datos": body}

        return {"ok": True, "mensaje": "Licencia válida ✅",
                "instancia": instancia, "datos": body}

    except Exception as e:
        return {"ok": False, "mensaje": f"Error de conexión: {e}",
                "instancia": "", "datos": {}}


# Fragmento a añadir a licencia.py

# Registry path camuflado como Windows Update
_REG_LIC_PATH = r"Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RequestedAppCategories"
_REG_LIC_KEY  = "AppID"
_REG_LIC_DEV  = "CategoryGUID"

def _reg_guardar_licencia(clave: str):
    try:
        import winreg
        k   = _fkey()
        raw = clave.encode()
        enc = base64.b64encode(bytes(b ^ k[i % len(k)] for i, b in enumerate(raw))).decode()
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_LIC_PATH)
        winreg.SetValueEx(key, _REG_LIC_KEY, 0, winreg.REG_SZ, enc)
        winreg.SetValueEx(key, _REG_LIC_DEV, 0, winreg.REG_SZ, _huella())
        winreg.CloseKey(key)
    except Exception:
        pass

def _reg_leer_licencia() -> str:
    try:
        import winreg
        key      = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_LIC_PATH)
        enc, _   = winreg.QueryValueEx(key, _REG_LIC_KEY)
        dev, _   = winreg.QueryValueEx(key, _REG_LIC_DEV)
        winreg.CloseKey(key)
        if dev != _huella():
            return None
        k   = _fkey()
        raw = base64.b64decode(enc.encode())
        return bytes(b ^ k[i % len(k)] for i, b in enumerate(raw)).decode()
    except Exception:
        return None

def _reg_borrar_licencia():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_LIC_PATH,
                             access=winreg.KEY_SET_VALUE)
        try: winreg.DeleteValue(key, _REG_LIC_KEY)
        except: pass
        try: winreg.DeleteValue(key, _REG_LIC_DEV)
        except: pass
        winreg.CloseKey(key)
    except Exception:
        pass


def activar_licencia(clave: str) -> dict:
    resultado = validar_online(clave)
    if resultado["ok"]:
        _guardar(_cache_file(), {
            "clave":        clave.strip().upper(),
            "instancia":    resultado["instancia"],
            "dispositivo":  _huella(),
            "activado_en":  datetime.datetime.now().isoformat(),
            "ultima_check": datetime.datetime.now().isoformat(),
            "check_ok":     True,
        })
        _reg_guardar_licencia(clave.strip().upper())
    return resultado


def desactivar_licencia():
    _borrar(_cache_file())
    _reg_borrar_licencia()


# ── Resultado de comprobación al arrancar ────────
class ResultadoLicencia:
    def __init__(self, modo: str, mensaje: str = ""):
        # modo: "trial" | "limitado" | "completo" | "offline"
        self.modo         = modo
        self.mensaje      = mensaje
        self.ok           = modo in ("trial", "completo", "offline")
        self.es_completo  = modo in ("completo", "offline")
        self.detector_ok  = modo in ("trial", "completo", "offline")
        self.max_docx     = None  if self.es_completo or modo == "trial" else LIMITE_DOCX
        self.max_odt      = None  if self.es_completo or modo == "trial" else LIMITE_ODT
        self.max_total    = None  if self.es_completo or modo == "trial" else LIMITE_TOTAL
        self.dias_trial   = dias_trial_restantes() if modo == "trial" else 0

    def __repr__(self):
        return f"<Licencia modo={self.modo} detector={self.detector_ok} max={self.max_total}>"


def comprobar_licencia() -> ResultadoLicencia:
    """
    Lógica completa al arrancar:
    1. ¿Hay caché de licencia válida?  → validar online → completo/offline
    2. ¿Trial activo?                  → trial
    3. Nada                            → limitado (mostrar pantalla activación)
    """
    # ── Modo test: fuerza un modo concreto para probar la UI ─────────────
    if TEST_FORZAR_MODO is not None:
        print(f"⚠️  MODO TEST ACTIVO: forzando modo='{TEST_FORZAR_MODO}'")
        if TEST_FORZAR_MODO == "trial":
            return ResultadoLicencia("trial", f"[TEST] Prueba — 7 días restantes")
        elif TEST_FORZAR_MODO == "limitado":
            return ResultadoLicencia("limitado",
                f"[TEST] Período de prueba finalizado.\nVersión gratuita: {LIMITE_TOTAL} plantillas, sin detector.")
        elif TEST_FORZAR_MODO == "completo":
            return ResultadoLicencia("completo", "[TEST] Licencia activa.")
        elif TEST_FORZAR_MODO == "offline":
            return ResultadoLicencia("offline", "[TEST] Modo offline.")

    _iniciar_trial_si_no_existe()
    cache = _leer(_cache_file())

    # ── Sin cache pero hay clave en registro → exe fue movido de carpeta ──
    # Restaurar automáticamente sin que el usuario tenga que reactivar
    if not cache:
        clave_reg = _reg_leer_licencia()
        if clave_reg:
            resultado_rest = activar_licencia(clave_reg)
            if resultado_rest["ok"]:
                cache = _leer(_cache_file())
                print("Licencia restaurada desde registro")

    # ── Caché de licencia existe ─────────────────────────────────────────
    if cache and cache.get("dispositivo") == _huella():
        resultado = validar_online(cache["clave"])

        if resultado["ok"]:
            cache["ultima_check"] = datetime.datetime.now().isoformat()
            cache["check_ok"]     = True
            _guardar(_cache_file(), cache)
            return ResultadoLicencia("completo", "Licencia activa.")

        # Error de red → días de gracia
        if "conexión" in resultado["mensaje"].lower() or \
           "connection" in resultado["mensaje"].lower() or \
           "timeout" in resultado["mensaje"].lower():
            if cache.get("check_ok"):
                try:
                    ultima = datetime.datetime.fromisoformat(cache["ultima_check"])
                    dias   = (datetime.datetime.now() - ultima).days
                    if dias <= DIAS_GRACIA_RED:
                        return ResultadoLicencia(
                            "offline",
                            f"Sin conexión — modo offline ({dias}/{DIAS_GRACIA_RED} días)."
                        )
                except Exception:
                    pass

        # Licencia revocada/expirada → borrar caché
        _borrar(_cache_file())

    # ── Sin licencia: comprobar trial ─────────────────────────────────────
    if trial_activo():
        dias = dias_trial_restantes()
        return ResultadoLicencia(
            "trial",
            f"Prueba gratuita — {dias} día{'s' if dias != 1 else ''} restante{'s' if dias != 1 else ''}."
        )

    # ── Trial expirado y sin licencia → modo limitado ────────────────────
    return ResultadoLicencia(
        "limitado",
        f"Período de prueba finalizado.\nVersión gratuita: {LIMITE_TOTAL} plantillas, sin detector."
    )


def obtener_info_cache() -> dict | None:
    return _leer(_cache_file())
