# core.py
import os
import html
import base64
import win32clipboard
from docx import Document
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
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
from logging.handlers import RotatingFileHandler
import zipfile
import shutil

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

# SOLO ESTOS DOS ARCHIVOS (config.json ya no se usa en core, pero lo dejamos por compatibilidad)
CONFIG_FILE = os.path.join(AJUSTES_DIR, "config.json")
PREFERENCIAS_FILE = os.path.join(AJUSTES_DIR, "preferencias.json")
WINDOW_TITLE = "Plantillas"

class GestorPlantillas:
    """Clase principal que maneja toda la lógica de la aplicación"""
    
    def __init__(self):
        self.preferencias = self.cargar_preferencias()
        self.logger = None
        self.buffer_teclas = []
        self.popup_activo = None
        self.ultimo_tiempo = time.time()
        self._detector_thread = None
        
        # Inicializar logging si está activado
        if self.preferencias.get("debug", False):
            self.setup_logging()
    
    # ==== SISTEMA DE LOGGING ====
    def setup_logging(self):
        """Configura el sistema de logging"""
        log_dir = os.path.join(AJUSTES_DIR, "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        fecha_actual = datetime.datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(log_dir, f"debug_{fecha_actual}.log")
        
        self.logger = logging.getLogger('PlantillasApp')
        self.logger.setLevel(logging.DEBUG)
        
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='a')
        file_handler.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        self.log("Sistema de logging iniciado", "info")
    
    def log(self, mensaje, nivel="debug", datos_sensibles=False):
        """Registra mensajes de log"""
        if not self.preferencias.get("debug", False):
            return
        
        if datos_sensibles and nivel == "debug":
            return
        
        if self.logger is None:
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
            elif nivel == "critical":
                self.logger.critical(mensaje)
        except:
            pass
    
    def log_excepcion(self, e, contexto=""):
        """Registra una excepción"""
        if not self.preferencias.get("debug", False):
            return
        error_msg = f"EXCEPCIÓN: {contexto} - {str(e)}\n{traceback.format_exc()}"
        self.log(error_msg, "error")
    
    # ==== FUNCIONES DE UTILIDAD ====
    @staticmethod
    def eliminar_acentos(texto):
        """Elimina acentos de un texto"""
        texto = unicodedata.normalize('NFD', texto)
        texto = "".join(c for c in texto if unicodedata.category(c) != 'Mn')
        return texto.lower()
    
    # ==== GESTIÓN DE PREFERENCIAS ====
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
        """Guarda las preferencias del usuario"""
        try:
            os.makedirs(AJUSTES_DIR, exist_ok=True)
            with open(PREFERENCIAS_FILE, "w", encoding="utf-8") as f:
                json.dump(nuevas_preferencias, f, ensure_ascii=False, indent=4)
            self.preferencias = nuevas_preferencias
            return True
        except Exception as e:
            print("Error guardando preferencias:", e)
            return False
    
    # ==== GESTIÓN DE CONFIGURACIÓN (ELIMINADA - AHORA EN INTERFAZ) ====
    # NOTA: Los métodos cargar_configuracion y guardar_configuracion 
    # se han eliminado porque ahora son responsabilidad de la interfaz
    
    # ==== MANEJO DEL PORTAPAPELES ====
    HTML_HEADER = """Version:1.0
StartHTML:{st_html:0>10}
EndHTML:{end_html:0>10}
StartFragment:{st_frag:0>10}
EndFragment:{end_frag:0>10}
"""
    
    def make_html_clipboard(self, html_body: str) -> bytes:
        """Prepara HTML para el portapapeles"""
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
    
    def ET_to_plaintext(self, html_string: str) -> str:
        """Convierte HTML a texto plano"""
        text = html.unescape(html_string)
        text = re.sub(r'<[^>]+>', '', text)
        return text
    
    def set_clipboard_html(self, html_string: str):
        """Copia HTML al portapapeles"""
        try:
            data = self.make_html_clipboard(html_string)
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            plain_text = self.ET_to_plaintext(html_string)
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(cf_html, data)
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)
            finally:
                win32clipboard.CloseClipboard()
            
            self.log("HTML copiado al portapapeles", "info")
        except Exception as e:
            print("Error al poner en portapapeles:", e)
            self.log_excepcion(e, "set_clipboard_html")
    
    # ==== CONVERSIÓN DOCX A HTML ====
    def docx_to_html_base64(self, docx_path):
        """Convierte DOCX a HTML conservando: negritas, cursiva, subrayado,
        tachado, color, tamaño, superíndice/subíndice, headings, tablas,
        listas multinivel, hipervínculos e imágenes."""
        from lxml import etree

        doc  = Document(docx_path)
        self.log(f"Convirtiendo DOCX a HTML: {os.path.basename(docx_path)}", "debug")

        # ── Imágenes ──────────────────────────────────────────────────────
        rels_img = {r.rId: r.target_part
                    for r in doc.part.rels.values() if "image" in r.reltype}

        # ── Resolver formato de lista desde numbering.xml ─────────────────
        def _resolver_lista(para):
            """Devuelve (tipo, nivel) o (None, 0)."""
            numPr = para._p.find(".//w:numPr", para._p.nsmap)
            if numPr is None:
                return None, 0
            ilvl_el = numPr.find("w:ilvl", para._p.nsmap)
            numId_el = numPr.find("w:numId", para._p.nsmap)
            nivel = int(ilvl_el.get(qn("w:val"), 0)) if ilvl_el is not None else 0
            numId = numId_el.get(qn("w:val"), "0") if numId_el is not None else "0"

            # Intentar resolver desde numbering.xml
            try:
                numbering_part = doc.part.numbering_part
                nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                abstractNumId_el = numbering_part._element.find(
                    f".//w:num[@w:numId='{numId}']/w:abstractNumId", nsmap)
                if abstractNumId_el is not None:
                    abId = abstractNumId_el.get(qn("w:val"), "0")
                    lvl_el = numbering_part._element.find(
                        f".//w:abstractNum[@w:abstractNumId='{abId}']"
                        f"/w:lvl[@w:ilvl='{nivel}']/w:numFmt", nsmap)
                    if lvl_el is not None:
                        fmt = lvl_el.get(qn("w:val"), "").lower()
                        return ("ul" if fmt == "bullet" else "ol"), nivel
            except Exception:
                pass
            return "ul", nivel   # fallback

        # ── Obtener estilo de párrafo ──────────────────────────────────────
        def _estilo_parrafo(para):
            """Devuelve el nombre del estilo ('Heading 1', 'Normal', etc.)."""
            pStyle = para._p.find(".//w:pStyle", para._p.nsmap)
            if pStyle is not None:
                return pStyle.get(qn("w:val"), "")
            return ""

        def _alineacion(para):
            jc = para._p.find(".//w:jc", para._p.nsmap)
            if jc is not None:
                v = jc.get(qn("w:val"), "").lower()
                m = {"center": "center", "right": "right",
                     "both": "justify", "distribute": "justify"}
                return m.get(v, "")
            return ""

        # ── Procesar un <w:r> → HTML inline ───────────────────────────────
        URL_RE = re.compile(r'((?:https?://|www\.)[^\s<>"]+)')

        # Namespaces para imágenes
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
        NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        def _extraer_imagen(r_el):
            """Busca imagen en <w:r><w:drawing> y devuelve <img> o vacío."""
            drawing = r_el.find(f"{{{r_el.nsmap.get('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')}}}drawing")
            if drawing is None:
                # Buscar con xpath completo
                drawing = r_el.find(".//w:drawing", r_el.nsmap)
            if drawing is None:
                return ""
            # Buscar blip con namespace correcto
            blip = drawing.find(f".//{{{NS_A}}}blip")
            if blip is None:
                return ""
            rId = blip.get(f"{{{NS_R}}}embed")
            if not rId:
                rId = blip.get(qn("r:embed"))
            if rId and rId in rels_img:
                img = rels_img[rId]
                ext = img.content_type.split("/")[-1]
                b64 = base64.b64encode(img.blob).decode()
                return (f'<img src="data:image/{ext};base64,{b64}" '
                        f'style="max-width:100%;display:block;margin:4px 0;"/>')
            return ""

        def _run_to_html(r_el):
            # Comprobar primero si es un run con imagen
            img_html = _extraer_imagen(r_el)
            if img_html:
                return img_html

            txt = "".join(t.text or ""
                          for t in r_el.findall(".//w:t", r_el.nsmap))
            if not txt:
                return ""
            safe = html.escape(txt)
            safe = URL_RE.sub(r'<a href="\1">\1</a>', safe)

            styles = []
            rPr = r_el.find("w:rPr", r_el.nsmap)
            if rPr is not None:
                # Negrita
                b = rPr.find("w:b", rPr.nsmap)
                if b is not None and b.get(qn("w:val"), "true") not in ("0","false"):
                    styles.append("font-weight:bold")
                # Cursiva
                i = rPr.find("w:i", rPr.nsmap)
                if i is not None and i.get(qn("w:val"), "true") not in ("0","false"):
                    styles.append("font-style:italic")
                # Subrayado
                u = rPr.find("w:u", rPr.nsmap)
                if u is not None and u.get(qn("w:val"), "none") not in ("none",""):
                    styles.append("text-decoration:underline")
                # Tachado
                strike = rPr.find("w:strike", rPr.nsmap)
                if strike is not None and strike.get(qn("w:val"), "true") not in ("0","false"):
                    styles.append("text-decoration:line-through")
                # Color
                color_el = rPr.find("w:color", rPr.nsmap)
                if color_el is not None:
                    c = color_el.get(qn("w:val"), "auto")
                    if c and c.lower() not in ("auto", "000000"):
                        styles.append(f"color:#{c}")
                # Tamaño
                sz = rPr.find("w:sz", rPr.nsmap)
                if sz is not None:
                    try:
                        pt = int(sz.get(qn("w:val"), 0)) / 2
                        if pt > 0:
                            styles.append(f"font-size:{pt}pt")
                    except Exception:
                        pass
                # Superíndice / subíndice
                vertAlign = rPr.find("w:vertAlign", rPr.nsmap)
                if vertAlign is not None:
                    v = vertAlign.get(qn("w:val"), "")
                    if v == "superscript":
                        safe = f"<sup>{safe}</sup>"
                    elif v == "subscript":
                        safe = f"<sub>{safe}</sub>"
                # Resaltado
                highlight = rPr.find("w:highlight", rPr.nsmap)
                if highlight is not None:
                    col_map = {"yellow":"#ffff00","green":"#00ff00","cyan":"#00ffff",
                               "magenta":"#ff00ff","blue":"#0000ff","red":"#ff0000",
                               "darkBlue":"#00008b","darkCyan":"#008b8b",
                               "darkGreen":"#006400","darkMagenta":"#8b008b",
                               "darkRed":"#8b0000","darkYellow":"#808000",
                               "darkGray":"#808080","lightGray":"#d3d3d3"}
                    hc = col_map.get(highlight.get(qn("w:val"), ""), "")
                    if hc:
                        styles.append(f"background-color:{hc}")

            if styles:
                return f'<span style="{";".join(styles)}">{safe}</span>'
            return safe

        # ── Procesar párrafo → HTML ────────────────────────────────────────
        def _parrafo_a_html(para, tag="p", extra_style=""):
            parts = []
            align = _alineacion(para)
            style_parts = []
            if align:
                style_parts.append(f"text-align:{align}")
            if extra_style:
                style_parts.append(extra_style)
            style_attr = f' style="{";".join(style_parts)}"' if style_parts else ""

            for child in para._p:
                ctag = child.tag.split("}")[-1]
                if ctag == "hyperlink":
                    rId = child.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    href = para.part.rels[rId].target_ref if rId and rId in para.part.rels else None
                    inner = "".join(_run_to_html(r)
                                    for r in child.findall(".//w:r", child.nsmap))
                    if href:
                        parts.append(f'<a href="{html.escape(href)}">{inner}</a>')
                    else:
                        parts.append(inner)
                elif ctag == "r":
                    parts.append(_run_to_html(child))
                elif ctag == "drawing":
                    # Drawing directo en el párrafo (no dentro de <w:r>)
                    blip = child.find(f".//{{{NS_A}}}blip")
                    if blip is not None:
                        rId = blip.get(f"{{{NS_R}}}embed") or blip.get(qn("r:embed"))
                        if rId and rId in rels_img:
                            img = rels_img[rId]
                            ext = img.content_type.split("/")[-1]
                            b64 = base64.b64encode(img.blob).decode()
                            parts.append(
                                f'<img src="data:image/{ext};base64,{b64}" '
                                f'style="max-width:100%;display:block;margin:4px 0;"/>')

            joined = "".join(parts).strip()
            if not joined:
                return None
            return f"<{tag}{style_attr}>{joined}</{tag}>"

        # ── Procesar tabla → HTML ──────────────────────────────────────────
        def _tabla_a_html(tabla):
            rows_html = []
            for row in tabla.rows:
                cells_html = []
                for cell in row.cells:
                    # Contenido de la celda: sus párrafos
                    cell_content = []
                    for p in cell.paragraphs:
                        h = _parrafo_a_html(p)
                        if h:
                            cell_content.append(h)
                    inner = "".join(cell_content) or "&nbsp;"
                    cells_html.append(f'<td style="border:1px solid #ccc;padding:4px 8px;">{inner}</td>')
                rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
            return ('<table style="border-collapse:collapse;width:100%;margin:6px 0;">' +
                    "".join(rows_html) + "</table>")

        # ── Configuración de fuente global ────────────────────────────────
        fuente_cfg  = self.preferencias.get("fuente_html", {})
        modo_fuente = fuente_cfg.get("modo", "origen")
        if modo_fuente == "personalizado":
            fam = fuente_cfg.get("familia", "Arial")
            fsz = fuente_cfg.get("tamaño", 10)
            fcl = fuente_cfg.get("color", "#000000")
            html_parts = [f'<div style="font-family:{fam},sans-serif;font-size:{fsz}pt;color:{fcl};">']
        else:
            html_parts = ["<div>"]

        list_stack    = []   # [(tipo, nivel)]
        last_was_empty = False

        # ── Iterar elementos del body en orden ────────────────────────────
        # doc.element.body contiene párrafos y tablas en orden real
        HEADING_RE = re.compile(r"Heading\s*(\d+)", re.IGNORECASE)
        HEADING_SIZES = {"1":"2em","2":"1.5em","3":"1.17em",
                         "4":"1em","5":"0.83em","6":"0.67em"}

        for block in doc.element.body:
            btag = block.tag.split("}")[-1]

            # ── Tabla ─────────────────────────────────────────────────────
            if btag == "tbl":
                while list_stack:
                    html_parts.append(f'</{list_stack.pop()[0]}>')
                from docx.table import Table as DocxTable
                tabla = DocxTable(block, doc)
                html_parts.append(_tabla_a_html(tabla))
                last_was_empty = False
                continue

            # ── Párrafo ───────────────────────────────────────────────────
            if btag != "p":
                continue

            from docx.text.paragraph import Paragraph as DocxParagraph
            para = DocxParagraph(block, doc)

            list_type, nivel = _resolver_lista(para)
            estilo = _estilo_parrafo(para)

            # Gestionar stack de listas
            if list_type:
                while list_stack and list_stack[-1][1] > nivel:
                    html_parts.append(f'</{list_stack.pop()[0]}>')
                if not list_stack or list_stack[-1][1] < nivel:
                    ltype = list_type
                    ls_style = "disc" if ltype == "ul" else "decimal"
                    html_parts.append(
                        f'<{ltype} style="margin-left:{20*(nivel+1)}px;'
                        f'list-style-type:{ls_style};">')
                    list_stack.append((ltype, nivel))
            else:
                while list_stack:
                    html_parts.append(f'</{list_stack.pop()[0]}>')

            # Headings
            hm = HEADING_RE.match(estilo)
            if hm and not list_type:
                n = hm.group(1)
                sz = HEADING_SIZES.get(n, "1em")
                h = _parrafo_a_html(
                    para, tag=f"h{n}",
                    extra_style=f"font-size:{sz};font-weight:bold;margin:8px 0 4px")
                if h:
                    html_parts.append(h)
                last_was_empty = False
                continue

            # Párrafo normal o ítem de lista
            h = _parrafo_a_html(para)
            if h is None:
                if not last_was_empty:
                    html_parts.append("<p><br></p>")
                    last_was_empty = True
            else:
                last_was_empty = False
                if list_type:
                    html_parts.append(f"<li>{h[3:-4]}</li>")  # quitar <p>...</p>
                else:
                    html_parts.append(h)

        while list_stack:
            html_parts.append(f'</{list_stack.pop()[0]}>')

        html_parts.append("</div>")
        self.log("Conversión DOCX → HTML finalizada", "debug")
        return "".join(html_parts)

    # ==== PROCESAR PLANTILLA ====
    def procesar_plantilla(self, path):
        """Procesa una plantilla y la copia al portapapeles"""
        try:
            self.log(f"Procesando plantilla: {path}", "info")
            if not os.path.exists(path) or os.path.getsize(path) == 0:
                print("Error: El archivo está vacío o no existe")
                self.log(f"Archivo vacío o no existe: {path}", "error")
                return False
            
            ext = os.path.splitext(path)[1].lower()
            if ext == ".docx":
                try:
                    html_result = self.docx_to_html_base64(path)
                except Exception as e:
                    print(f"Error DOCX: {e}")
                    self.log_excepcion(e, f"Error procesando DOCX: {path}")
                    return False
            else:
                print("Error: Solo se admiten archivos .docx")
                self.log(f"Formato no admitido: {ext}", "error")
                return False
            
            self.set_clipboard_html(html_result)
            print(f"Éxito: '{os.path.basename(path)}' copiada")
            self.log(f"Plantilla copiada: {os.path.basename(path)}", "info")
            return True
        except Exception as e:
            print(f"Error inesperado: {e}")
            self.log_excepcion(e, f"Error inesperado procesando {path}")
            return False
    
    # ==== BUSCAR PLANTILLAS ====
    def buscar_plantillas(self, texto_filtro):
        """Busca plantillas que coincidan con el filtro"""
        ruta = self.preferencias.get("ruta_plantillas", PLANTILLAS_DIR)
        if not os.path.exists(ruta):
            ruta = PLANTILLAS_DIR
            
        resultados = []
        filtro_limpio = self.eliminar_acentos(texto_filtro)
        max_res = self.preferencias.get("max_resultados_popup", 10)
        
        for raiz, carpetas, archivos in os.walk(ruta):
            for f in archivos:
                if f.lower().endswith('.docx'):
                    nombre = os.path.splitext(f)[0]
                    nombre_limpio = self.eliminar_acentos(nombre)
                    
                    if filtro_limpio in nombre_limpio or texto_filtro in nombre.lower():
                        ruta_completa = os.path.join(raiz, f)
                        resultados.append((nombre, ruta_completa))
                        
                        if len(resultados) >= max_res:
                            break
            if len(resultados) >= max_res:
                break
        
        return sorted(resultados)
    
    # ==== DETECTOR DE COMANDOS ====
    def iniciar_detector(self, callback_mostrar_popup=None, callback_actualizar_popup=None):
        """Inicia el detector de comandos en un hilo separado"""
        if not self.preferencias.get("detector_activado", True):
            print("Detector de comandos desactivado en configuración.")
            return
        
        self.callback_mostrar_popup = callback_mostrar_popup
        self.callback_actualizar_popup = callback_actualizar_popup
        
        def detectar():
            """Función interna del detector"""
            def on_key(event):
                # Comprobar si el detector está activado
                if not self.preferencias.get("detector_activado", True):
                    return
                
                modo_privacidad = self.preferencias.get("privacidad_logs", True)
                detonante = self.preferencias.get("detonante", "/")
                tiempo_reinicio = self.preferencias.get("tiempo_reinicio_buffer", 2)
                
                # Si es una tecla especial, ignorar (excepto backspace y enter)
                if len(event.name) > 1 and event.name not in ['backspace', 'enter', 'space']:
                    return
                
                # === NUEVO: Detectar si estamos empezando un comando ===
                # Si la tecla actual es el detonante (o parte de él)
                if event.name == detonante:
                    # Limpiar buffer y empezar nuevo comando
                    self.buffer_teclas = [detonante]
                    self.ultimo_tiempo = time.time()
                    self.log(f"🎯 Nuevo comando iniciado con '{detonante}'", "info")
                    return
                
                # Si hay un popup activo, manejar la actualización del filtro
                if self.popup_activo and hasattr(self.popup_activo, 'winfo_exists') and self.popup_activo.winfo_exists():
                    if event.name == 'backspace':
                        if len(self.buffer_teclas) > len(detonante):
                            self.buffer_teclas.pop()
                            buffer_actual = ''.join(self.buffer_teclas)
                            if buffer_actual.startswith(detonante):
                                filtro = buffer_actual[len(detonante):]
                                if self.callback_actualizar_popup:
                                    self.callback_actualizar_popup(filtro)
                    elif len(event.name) == 1 or event.name == 'space':
                        self.buffer_teclas.append(event.name if event.name != 'space' else ' ')
                        buffer_actual = ''.join(self.buffer_teclas)
                        if buffer_actual.startswith(detonante):
                            filtro = buffer_actual[len(detonante):]
                            if self.callback_actualizar_popup:
                                self.callback_actualizar_popup(filtro)
                    return
                
                # === NUEVO: Si no hay popup, solo guardar en buffer si empezamos con detonante ===
                # Reiniciar buffer si pasa tiempo sin escribir
                if time.time() - self.ultimo_tiempo > tiempo_reinicio:
                    self.buffer_teclas = []
                
                self.ultimo_tiempo = time.time()
                
                # Manejar backspace
                if event.name == 'backspace' and self.buffer_teclas:
                    self.buffer_teclas.pop()
                    buffer_actual = ''.join(self.buffer_teclas)
                    if buffer_actual == detonante and self.popup_activo:
                        if hasattr(self.popup_activo, 'destroy'):
                            self.popup_activo.destroy()
                        self.popup_activo = None
                    return
                
                # Añadir tecla al buffer SOLO si ya empezamos con el detonante
                if len(event.name) == 1 or event.name == 'space':
                    # Si el buffer está vacío, solo guardar si es el detonante
                    if not self.buffer_teclas:
                        if event.name == detonante:
                            self.buffer_teclas.append(event.name)
                            self.log(f"Detonante '{detonante}' detectado", "info")
                        return
                    
                    # Si ya tenemos algo en el buffer, añadir la tecla
                    self.buffer_teclas.append(event.name if event.name != 'space' else ' ')
                    buffer_actual = ''.join(self.buffer_teclas)
                    
                    # Si el buffer empieza con el detonante y tiene más caracteres
                    if buffer_actual.startswith(detonante) and len(buffer_actual) > len(detonante):
                        filtro = buffer_actual[len(detonante):]
                        try:
                            punto = wintypes.POINT()
                            ctypes.windll.user32.GetCursorPos(ctypes.byref(punto))
                            x, y = punto.x, punto.y
                            if self.callback_mostrar_popup:
                                self.callback_mostrar_popup(filtro, x, y)
                        except Exception as e:
                            self.log_excepcion(e, "mostrar_popup desde detector")
                            if self.callback_mostrar_popup:
                                self.callback_mostrar_popup(filtro)
                    # Si no empieza con detonante, limpiar buffer
                    elif not buffer_actual.startswith(detonante):
                        self.buffer_teclas = []
            
            keyboard.on_press(on_key)
            keyboard.wait()
        
        self._detector_thread = threading.Thread(target=detectar, daemon=True)
        self._detector_thread.start()
        detonante = self.preferencias.get("detonante", "/")
        print(f"Detector de comandos '{detonante}' activado.")
    
    def limpiar_buffer(self):
        """Limpia el buffer de teclas"""
        self.buffer_teclas = []
    
    def set_popup_activo(self, popup):
        """Establece el popup activo"""
        self.popup_activo = popup
    
    def get_popup_activo(self):
        """Obtiene el popup activo"""
        return self.popup_activo