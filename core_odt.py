# core_odt.py
import os
import html
import base64
import win32clipboard
from odf.opendocument import load
from odf import text
import json
import traceback
import datetime
import re
import unicodedata
import threading
import time
import keyboard
import sys
import ctypes
from ctypes import wintypes
import logging
import zipfile

# ==== CONFIGURACIÓN BASE ====
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLAS_DIR = os.path.join(BASE_DIR, "Plantillas")

# Ajustes en AppData si está disponible (permisos garantizados en empresas)
_appdata = os.environ.get("APPDATA", "")
if _appdata and os.path.isdir(_appdata):
    AJUSTES_DIR = os.path.join(_appdata, "PlantillasMacro", "Ajustes")
else:
    AJUSTES_DIR = os.path.join(BASE_DIR, "Ajustes")

if not os.path.exists(AJUSTES_DIR):
    os.makedirs(AJUSTES_DIR)

# MISMO ARCHIVO DE CONFIGURACIÓN QUE CORE.PY (aunque ya no se usa)
CONFIG_FILE = os.path.join(AJUSTES_DIR, "config.json")
# MISMO ARCHIVO DE PREFERENCIAS
PREFERENCIAS_FILE = os.path.join(AJUSTES_DIR, "preferencias.json")

class GestorPlantillasODT:
    """Clase para manejar plantillas ODT - Genera HTML como LibreOffice"""
    
    def __init__(self):
        self.preferencias = self.cargar_preferencias()
        self.logger = None
        self.buffer_teclas = []
        self.popup_activo = None
        self.ultimo_tiempo = time.time()
        self._detector_thread = None
        
        if self.preferencias.get("debug", False):
            self.setup_logging()
    
    # ==== SISTEMA DE LOGGING ====
    def setup_logging(self):
        log_dir = os.path.join(AJUSTES_DIR, "logs_odt")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        fecha_actual = datetime.datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(log_dir, f"debug_{fecha_actual}.log")
        
        self.logger = logging.getLogger('PlantillasODT')
        self.logger.setLevel(logging.DEBUG)
        
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='a')
        file_handler.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [ODT] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        self.log("Sistema de logging ODT iniciado", "info")
    
    def log(self, mensaje, nivel="debug", datos_sensibles=False):
        if not self.preferencias.get("debug", False):
            return
        
        if self.logger is None:
            print(f"📌 ODT: {mensaje}")
            return
        
        try:
            if nivel == "debug":
                self.logger.debug(mensaje)
            elif nivel == "info":
                self.logger.info(mensaje)
            elif nivel == "warning":
                self.logger.warning(mensaje)
            elif nivel == "error":
                self.logger.error(mensaje)
        except:
            print(f"📌 ODT: {mensaje}")
    
    def log_excepcion(self, e, contexto=""):
        error_msg = f"EXCEPCIÓN ODT: {contexto} - {str(e)}"
        print(f"❌ {error_msg}")
        if self.preferencias.get("debug", False) and self.logger:
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
    
    # ==== FUNCIONES BÁSICAS ====
    @staticmethod
    def eliminar_acentos(texto):
        texto = unicodedata.normalize('NFD', texto)
        texto = "".join(c for c in texto if unicodedata.category(c) != 'Mn')
        return texto.lower()
    
    # ==== PREFERENCIAS ====
    def cargar_preferencias(self):
        """Carga las preferencias del usuario (debe existir el archivo)"""
        try:
            if os.path.exists(PREFERENCIAS_FILE):
                with open(PREFERENCIAS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    
                    # SANITIZAR VALORES
                    data["detonante"] = data.get("detonante") or "/"
                    data["ruta_plantillas"] = data.get("ruta_plantillas") or PLANTILLAS_DIR
                    data["debug"] = data.get("debug") in [True, "True", "true", 1, "1"]
                    data["privacidad_logs"] = data.get("privacidad_logs") in [True, "True", "true", 1, "1"]
                    data["detector_activado"] = data.get("detector_activado") in [True, "True", "true", 1, "1"]
                    data["pegar_automaticamente"] = data.get("pegar_automaticamente") in [True, "True", "true", 1, "1"]
                    data["mostrar_iconos"] = data.get("mostrar_iconos") in [True, "True", "true", 1, "1"]
                    data["atajos_activados"] = data.get("atajos_activados") in [True, "True", "true", 1, "1"]
                    
                    # Números
                    try:
                        data["tiempo_reinicio_buffer"] = int(data.get("tiempo_reinicio_buffer") or 2)
                    except:
                        data["tiempo_reinicio_buffer"] = 2
                    
                    try:
                        data["max_resultados_popup"] = int(data.get("max_resultados_popup") or 100)
                    except:
                        data["max_resultados_popup"] = 100
                    
                    try:
                        data["ancho_minimo_popup"] = int(data.get("ancho_minimo_popup") or 15)
                    except:
                        data["ancho_minimo_popup"] = 15
                    
                    # fuente_html
                    if "fuente_html" not in data or not isinstance(data["fuente_html"], dict):
                        data["fuente_html"] = {}
                    
                    data["fuente_html"]["modo"] = data["fuente_html"].get("modo") or "origen"
                    data["fuente_html"]["familia"] = data["fuente_html"].get("familia") or "Arial"
                    
                    try:
                        data["fuente_html"]["tamaño"] = int(data["fuente_html"].get("tamaño") or 10)
                    except:
                        data["fuente_html"]["tamaño"] = 10
                    
                    data["fuente_html"]["color"] = data["fuente_html"].get("color") or "#000000"
                    
                    return data
            else:
                print(f"⚠️ Archivo de preferencias no encontrado: {PREFERENCIAS_FILE}")
                return {}
        except Exception as e:
            print(f"Error al cargar preferencias: {e}")
            return {}

    def guardar_preferencias(self, nuevas_preferencias):
        try:
            os.makedirs(AJUSTES_DIR, exist_ok=True)
            with open(PREFERENCIAS_FILE, "w", encoding="utf-8") as f:
                json.dump(nuevas_preferencias, f, ensure_ascii=False, indent=4)
            self.preferencias = nuevas_preferencias
            return True
        except Exception as e:
            print("Error guardando preferencias ODT:", e)
            return False
    
    # ==== CONFIGURACIÓN (ELIMINADA - AHORA EN INTERFAZ) ====
    # NOTA: Los métodos cargar_configuracion y guardar_configuracion 
    # se han eliminado porque ahora son responsabilidad de la interfaz
    
    # ==== PORTAPAPELES ====
    HTML_HEADER = """Version:1.0
StartHTML:{st_html:0>10}
EndHTML:{end_html:0>10}
StartFragment:{st_frag:0>10}
EndFragment:{end_frag:0>10}
"""
    
    def make_html_clipboard(self, html_body: str) -> bytes:
        fragment_start = "<!--StartFragment-->"
        fragment_end = "<!--EndFragment-->"
        if fragment_start not in html_body:
            html_full = "<html><body>" + fragment_start + html_body + fragment_end + "</body></html>"
        else:
            html_full = html_body
        
        header = self.HTML_HEADER.format(st_html=0, end_html=0, st_frag=0, end_frag=0)
        full = header + html_full
        b = full.encode("utf-8")
        st_html = len(header.encode("utf-8"))
        end_html = len(b)
        st_frag = full.index(fragment_start) + len(header)
        end_frag = full.index(fragment_end) + len(header)
        
        header = self.HTML_HEADER.format(st_html=st_html, end_html=end_html, st_frag=st_frag, end_frag=end_frag)
        final = (header + html_full).encode("utf-8")
        return final
    
    def set_clipboard_html(self, html_string: str):
        try:
            data = self.make_html_clipboard(html_string)
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            plain_text = re.sub(r'<[^>]+>', '', html.unescape(html_string))
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(cf_html, data)
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)
            finally:
                win32clipboard.CloseClipboard()
            
            self.log("HTML ODT copiado al portapapeles", "info")
        except Exception as e:
            self.log_excepcion(e, "set_clipboard_html")
    
    # ==== CONVERSIÓN ODT A HTML ====
    def odt_to_html(self, odt_path):
        """Convierte ODT a HTML leyendo el XML real del content.xml.
        Soporta: negritas, cursiva, subrayado, tachado, color, tamaño,
        superíndice/subíndice, headings, listas, tablas, hipervínculos e imágenes."""
        import zipfile
        from xml.etree import ElementTree as ET

        # ── Namespaces ODT ────────────────────────────────────────────────
        NS = {
            "text":  "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
            "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
            "fo":    "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
            "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
            "draw":  "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
            "xlink": "http://www.w3.org/1999/xlink",
            "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
            "svg":   "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
        }
        def T(ns, tag):
            return f"{{{NS[ns]}}}{tag}"

        def _attr(el, ns, name, default=""):
            return el.get(f"{{{NS[ns]}}}{name}", default)

        # ── Leer archivos del ZIP ─────────────────────────────────────────
        try:
            with zipfile.ZipFile(odt_path, "r") as z:
                content_xml = z.read("content.xml")
                try:
                    styles_xml = z.read("styles.xml")
                except KeyError:
                    styles_xml = b"<root/>"
                images = {name: z.read(name)
                          for name in z.namelist()
                          if name.startswith("Pictures/")}
        except Exception as e:
            self.log_excepcion(e, "odt_to_html lectura zip")
            return f"<div>Error leyendo ODT: {e}</div>"

        root_c  = ET.fromstring(content_xml)
        root_s  = ET.fromstring(styles_xml)

        # ── Recopilar estilos automáticos + globales ──────────────────────
        # Mapeamos nombre_estilo → dict de propiedades CSS
        estilos: dict = {}

        def _leer_props(style_el):
            props = {}
            # Propiedades de texto (fo:font-weight, fo:font-style, etc.)
            for pp in style_el.iter():
                fw = pp.get(f"{{{NS['fo']}}}font-weight", "")
                if fw == "bold":
                    props["font-weight"] = "bold"
                fi = pp.get(f"{{{NS['fo']}}}font-style", "")
                if fi == "italic":
                    props["font-style"] = "italic"
                td_parts = []
                if pp.get(f"{{{NS['style']}}}text-underline-style","") not in ("","none"):
                    td_parts.append("underline")
                if pp.get(f"{{{NS['style']}}}text-line-through-style","") not in ("","none"):
                    td_parts.append("line-through")
                if td_parts:
                    props["text-decoration"] = " ".join(td_parts)
                col = pp.get(f"{{{NS['fo']}}}color", "")
                if col and col != "#000000":
                    props["color"] = col
                fsz = pp.get(f"{{{NS['fo']}}}font-size", "")
                if fsz:
                    props["font-size"] = fsz
                bg = pp.get(f"{{{NS['fo']}}}background-color", "")
                if bg and bg not in ("transparent",""):
                    props["background-color"] = bg
                va = pp.get(f"{{{NS['style']}}}text-position", "")
                if "super" in va.lower():
                    props["vertical-align"] = "super"
                elif "sub" in va.lower():
                    props["vertical-align"] = "sub"
                # Alineación de párrafo
                ta = pp.get(f"{{{NS['fo']}}}text-align", "")
                if ta:
                    props["text-align"] = ta
            return props

        for tree in (root_c, root_s):
            for sty in tree.iter(T("style","style")):
                name = _attr(sty, "style", "name")
                if name:
                    estilos[name] = _leer_props(sty)
            for sty in tree.iter(T("style","default-style")):
                family = _attr(sty, "style", "family")
                if family:
                    estilos[f"__default_{family}"] = _leer_props(sty)

        def _css_from_style(name):
            props = estilos.get(name, {})
            return ";".join(f"{k}:{v}" for k, v in props.items() if v)

        # ── Configuración de fuente global (preferencias) ─────────────────
        fuente_cfg  = self.preferencias.get("fuente_html", {})
        modo_fuente = fuente_cfg.get("modo", "origen")
        if modo_fuente == "personalizado":
            fam = fuente_cfg.get("familia", "Arial")
            fsz = fuente_cfg.get("tamaño", 10)
            fcl = fuente_cfg.get("color", "#000000")
            html_parts = [f'<div style="font-family:{fam},sans-serif;font-size:{fsz}pt;color:{fcl};">']
        else:
            html_parts = ["<div>"]

        # ── Detectar si un estilo de párrafo es heading ───────────────────
        HEADING_TAGS = {
            "Heading_20_1":"h1","Heading_20_2":"h2","Heading_20_3":"h3",
            "Heading_20_4":"h4","Heading_20_5":"h5","Heading_20_6":"h6",
            "Heading 1":"h1","Heading 2":"h2","Heading 3":"h3",
        }
        HEADING_SIZES = {"h1":"2em","h2":"1.5em","h3":"1.17em",
                         "h4":"1em","h5":"0.83em","h6":"0.67em"}

        # ── Convertir un <text:span> o texto a HTML inline ────────────────
        URL_RE = re.compile(r'((?:https?://|www\.)[^\s<>"]+)')

        def _span_html(el):
            """Procesa texto de un nodo (texto directo + spans anidados)."""
            out = []
            # Texto directo antes del primer hijo
            if el.text:
                t = html.escape(el.text)
                t = URL_RE.sub(r'<a href="\1">\1</a>', t)
                out.append(t)
            for child in el:
                ctag = child.tag
                if ctag == T("text","span"):
                    sname = _attr(child, "text", "style-name")
                    css   = _css_from_style(sname)
                    inner = _span_html(child)
                    if css:
                        out.append(f'<span style="{css}">{inner}</span>')
                    else:
                        out.append(inner)
                elif ctag == T("text","a"):
                    href  = _attr(child, "xlink", "href")
                    inner = _span_html(child)
                    out.append(f'<a href="{html.escape(href)}">{inner}</a>')
                elif ctag == T("text","line-break"):
                    out.append("<br/>")
                elif ctag == T("text","tab"):
                    out.append("&nbsp;&nbsp;&nbsp;&nbsp;")
                elif ctag == T("text","s"):   # spaces
                    n = int(child.get(f"{{{NS['text']}}}c", 1))
                    out.append("&nbsp;" * n)
                elif ctag == T("draw","frame"):
                    # Imagen incrustada
                    img_el = child.find(T("draw","image"))
                    if img_el is not None:
                        href = _attr(img_el, "xlink", "href")
                        if href in images:
                            ext = href.rsplit(".", 1)[-1].lower()
                            b64 = base64.b64encode(images[href]).decode()
                            out.append(
                                f'<img src="data:image/{ext};base64,{b64}" '
                                f'style="max-width:100%;display:block;margin:4px 0;"/>')
                else:
                    # Texto del hijo desconocido
                    if child.text:
                        t = html.escape(child.text)
                        out.append(URL_RE.sub(r'<a href="\1">\1</a>', t))
                # Tail (texto tras la etiqueta de cierre)
                if child.tail:
                    t = html.escape(child.tail)
                    out.append(URL_RE.sub(r'<a href="\1">\1</a>', t))
            return "".join(out)

        # ── Convertir <text:p> a HTML ─────────────────────────────────────
        def _parrafo_html(p_el, list_item=False):
            sname = _attr(p_el, "text", "style-name")
            props = estilos.get(sname, {})
            inner = _span_html(p_el)
            if not inner.strip():
                return None

            # Alineación
            align = props.get("text-align", "")
            style_parts = []
            if align and align != "start":
                ta = {"end":"right","center":"center","justify":"justify"}.get(align, align)
                style_parts.append(f"text-align:{ta}")

            style_attr = f' style="{";".join(style_parts)}"' if style_parts else ""

            if list_item:
                return inner

            # Heading?
            htag = HEADING_TAGS.get(sname)
            if not htag:
                # Intentar por nombre de parent style
                parent = estilos.get(sname, {})
                for k in HEADING_TAGS:
                    if sname.startswith(k[:8]):
                        htag = HEADING_TAGS[k]
                        break
            if htag:
                hsz = HEADING_SIZES.get(htag, "1em")
                return f'<{htag} style="font-size:{hsz};font-weight:bold;margin:8px 0 4px{";"+";".join(style_parts) if style_parts else ""}">{inner}</{htag}>'

            return f"<p{style_attr}>{inner}</p>"

        # ── Procesar tabla ────────────────────────────────────────────────
        def _tabla_html(tbl_el):
            rows = []
            for row_el in tbl_el.iter(T("table","table-row")):
                cells = []
                for cell_el in row_el.iter(T("table","table-cell")):
                    cell_content = []
                    for child in cell_el:
                        if child.tag == T("text","p"):
                            h = _parrafo_html(child)
                            if h:
                                cell_content.append(h)
                        elif child.tag == T("text","list"):
                            cell_content.append(_lista_html(child))
                    inner = "".join(cell_content) or "&nbsp;"
                    cells.append(f'<td style="border:1px solid #ccc;padding:4px 8px;">{inner}</td>')
                if cells:
                    rows.append("<tr>" + "".join(cells) + "</tr>")
            return ('<table style="border-collapse:collapse;width:100%;margin:6px 0;">' +
                    "".join(rows) + "</table>")

        # ── Procesar lista ────────────────────────────────────────────────
        def _lista_html(list_el, nivel=0):
            # Detectar tipo desde el primer item
            list_style_name = _attr(list_el, "text", "style-name")
            # Intentar detectar ol vs ul desde el estilo
            is_ol = False
            if list_style_name:
                for tree in (root_c, root_s):
                    ls = tree.find(
                        f".//{{{NS['text']}}}list-style[@{{{NS['style']}}}name='{list_style_name}']")
                    if ls is not None:
                        if ls.find(T("text","list-level-style-number")) is not None:
                            is_ol = True
                        break
            tag    = "ol" if is_ol else "ul"
            indent = 20 * (nivel + 1)
            ls_type = "decimal" if is_ol else "disc"
            parts  = [f'<{tag} style="margin-left:{indent}px;list-style-type:{ls_type};">']
            for item_el in list_el:
                if item_el.tag != T("text","list-item"):
                    continue
                item_content = []
                for child in item_el:
                    if child.tag == T("text","p"):
                        h = _parrafo_html(child, list_item=True)
                        if h:
                            item_content.append(h)
                    elif child.tag == T("text","list"):
                        item_content.append(_lista_html(child, nivel + 1))
                parts.append("<li>" + "".join(item_content) + "</li>")
            parts.append(f"</{tag}>")
            return "".join(parts)

        # ── Recorrer el body del documento ────────────────────────────────
        body = root_c.find(f".//{{{NS['office']}}}text")
        if body is None:
            html_parts.append("<p>No se pudo leer el contenido del ODT</p>")
        else:
            last_empty = False
            for el in body:
                etag = el.tag
                if etag == T("text","p"):
                    h = _parrafo_html(el)
                    if h is None:
                        if not last_empty:
                            html_parts.append("<p><br></p>")
                            last_empty = True
                    else:
                        last_empty = False
                        html_parts.append(h)
                elif etag == T("text","h"):
                    # Headings marcados con <text:h>
                    level = int(_attr(el, "text", "outline-level") or "1")
                    level = max(1, min(6, level))
                    inner = _span_html(el)
                    if inner.strip():
                        sz = HEADING_SIZES.get(f"h{level}", "1em")
                        html_parts.append(
                            f'<h{level} style="font-size:{sz};font-weight:bold;'
                            f'margin:8px 0 4px">{inner}</h{level}>')
                        last_empty = False
                elif etag == T("text","list"):
                    html_parts.append(_lista_html(el))
                    last_empty = False
                elif etag == T("table","table"):
                    html_parts.append(_tabla_html(el))
                    last_empty = False

        html_parts.append("</div>")
        resultado = "".join(html_parts)
        self.log(f"HTML ODT generado: {len(resultado)} caracteres", "debug")
        return resultado

    # ==== PROCESAR PLANTILLA ====
    def procesar_plantilla(self, path):
        """Procesa plantilla ODT"""
        try:
            self.log(f"Procesando ODT: {path}", "info")
            
            if not os.path.exists(path):
                print("❌ Archivo ODT no existe")
                return False
            
            if os.path.getsize(path) == 0:
                print("❌ Archivo ODT vacío")
                return False
            
            html = self.odt_to_html(path)
            self.set_clipboard_html(html)
            print(f"✅ ODT copiado: {os.path.basename(path)}")
            return True
            
        except Exception as e:
            print(f"❌ Error ODT: {e}")
            self.log_excepcion(e, f"procesar_plantilla {path}")
            return False
    
    # ==== BUSCAR PLANTILLAS ====
    def buscar_plantillas(self, texto_filtro):
        """Busca plantillas ODT"""
        ruta = self.preferencias.get("ruta_plantillas", PLANTILLAS_DIR)
        if not os.path.exists(ruta):
            ruta = PLANTILLAS_DIR
        
        resultados = []
        filtro = self.eliminar_acentos(texto_filtro)
        max_res = self.preferencias.get("max_resultados_popup", 10)
        
        for root, dirs, files in os.walk(ruta):
            for f in files:
                if f.lower().endswith('.odt'):
                    nombre = os.path.splitext(f)[0]
                    if filtro in self.eliminar_acentos(nombre) or texto_filtro in nombre.lower():
                        resultados.append((nombre, os.path.join(root, f)))
                        if len(resultados) >= max_res:
                            break
            if len(resultados) >= max_res:
                break
        
        return sorted(resultados)
    
    # ==== DETECTOR (COMPATIBILIDAD) ====
    def iniciar_detector(self, *args, **kwargs):
        pass
    
    def limpiar_buffer(self):
        self.buffer_teclas = []
    
    def set_popup_activo(self, popup):
        self.popup_activo = popup
    
    def get_popup_activo(self):
        return self.popup_activo