import os
import zipfile
import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
import html
import base64
import win32clipboard
from docx import Document
from win10toast import ToastNotifier
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import json
import traceback
import datetime
import re
import unicodedata

# ==== CONFIGURACIÓN BASE ====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLAS_DIR = os.path.join(BASE_DIR, "Plantillas")
AJUSTES_DIR = os.path.join(BASE_DIR, "Ajustes")
if not os.path.exists(AJUSTES_DIR):
    os.makedirs(AJUSTES_DIR)
CONFIG_FILE = os.path.join(AJUSTES_DIR, "config.json")
WINDOW_TITLE = "Plantillas"

# --- Guardar y cargar configuración ---
def cargar_configuracion():
    default = {
        "modo_oscuro": False,
        "x": None, "y": None, "width": None, "height": None,
        "debug": False
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k in default:
                    if k in data:
                        default[k] = data[k]
    except Exception as e:
        print("No se pudo cargar la configuración:", e)
    return default

# cargamos settings global
settings = cargar_configuracion()

def guardar_configuracion(modo_oscuro=None, geom=None, debug=None):
    global settings
    try:
        if not isinstance(settings, dict):
            settings = cargar_configuracion()
        if modo_oscuro is not None:
            settings["modo_oscuro"] = bool(modo_oscuro)
        if geom:
            settings["x"], settings["y"], settings["width"], settings["height"] = geom
        if debug is not None:
            settings["debug"] = bool(debug)
        os.makedirs(AJUSTES_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False)
    except Exception as e:
        print("No se pudo guardar la configuración:", e)

# ==== CONSTANTES / NAMESPACES ====
NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'text':  'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'draw':  'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'xlink': 'http://www.w3.org/1999/xlink'
}

HTML_HEADER = """Version:1.0
StartHTML:{st_html:0>10}
EndHTML:{end_html:0>10}
StartFragment:{st_frag:0>10}
EndFragment:{end_frag:0>10}
"""

# ---- Inicializar notificador ----
toaster = ToastNotifier()

def mostrar_notificacion(titulo, mensaje, duracion=2):
    try:
        toaster.show_toast(titulo, mensaje, duration=duracion, threaded=True)
    except Exception:
        pass

def make_html_clipboard(html_body: str) -> bytes:
    fragment_start = "<!--StartFragment-->"
    fragment_end = "<!--EndFragment-->"
    if fragment_start not in html_body:
        html_full = "<html><body>" + fragment_start + html_body + fragment_end + "</body></html>"
    else:
        html_full = html_body
    header = HTML_HEADER.format(st_html=0, end_html=0, st_frag=0, end_frag=0)
    full = header + html_full
    b = full.encode("utf-8")
    st_html = len(header.encode("utf-8"))
    end_html = len(b)
    st_frag = full.index(fragment_start) + len(header)
    end_frag = full.index(fragment_end) + len(header)
    header = HTML_HEADER.format(st_html=st_html, end_html=end_html, st_frag=st_frag, end_frag=end_frag)
    final = (header + html_full).encode("utf-8")
    return final

def set_clipboard_html(html_string: str):
    try:
        data = make_html_clipboard(html_string)
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        plain_text = ET_to_plaintext(html_string)
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(cf_html, data)
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print("Error al poner en portapapeles:", e)
        traceback.print_exc()

def ET_to_plaintext(html_string: str) -> str:
    import re
    text = html.unescape(html_string)
    text = re.sub(r'<[^>]+>', '', text)
    return text

# ---- DOCX a HTML con viñetas y enlaces (con logging a archivo si debug=True) ----

def docx_to_html_base64(docx_path):
    from lxml import etree
    import base64, datetime

    doc = Document(docx_path)

    # Relaciones de imágenes
    rels = {r.rId: r.target_part for r in doc.part.rels.values() if "image" in r.reltype}

    def convertir_urls_a_links(texto):
        patron = r'((?:https?://|www\.)[^\s<>"]+)'
        return re.sub(patron, r'<a href="\1">\1</a>', texto)

    # ===== FUNCIONES AUXILIARES (v2.1) =====
    def obtener_nivel_lista(para):
        ilvl = para._p.find('.//w:ilvl', para._p.nsmap)
        if ilvl is not None:
            try:
                return int(ilvl.attrib.get(qn("w:val"), 0))
            except Exception:
                return 0
        return 0

    def detectar_tipo_lista(para):
        numPr = para._p.find('.//w:numPr', para._p.nsmap)
        if numPr is None:
            return None

        numFmt = para._p.find('.//w:numFmt', para._p.nsmap)
        if numFmt is not None:
            fmt = numFmt.attrib.get(qn("w:val"), "").lower()
            if fmt == "bullet":
                return "ul"
            return "ol"
        return "ul"

    # ===== LOG (sin cambios) =====
    log_file = None
    if settings.get("debug", False):
        os.makedirs(os.path.join(AJUSTES_DIR, "log"), exist_ok=True)
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        log_file = os.path.join(AJUSTES_DIR, "log", f"{base_name}_{timestamp}_log.txt")

        def log(msg):
            try:
                with open(log_file, "a", encoding="utf-8") as f:
                    f.write(msg + "\n")
            except Exception:
                pass
    else:
        def log(msg):
            pass

    html_parts = ['<div style="font-family: Arial, sans-serif; font-size:10pt; color:#000000;">']
    list_stack = []
    last_was_empty = False

    # ===== PROCESAR PÁRRAFOS =====
    for para_idx, para in enumerate(doc.paragraphs, 1):
        list_type = detectar_tipo_lista(para)
        nivel = obtener_nivel_lista(para)

        # --- Gestión de listas (apertura / cierre) ---
        if list_type:
            while len(list_stack) > nivel:
                html_parts.append(f'</{list_stack.pop()}>')

            if len(list_stack) < nivel + 1:
                html_parts.append(
                    f'<{list_type} style="margin-left:20px; list-style-type:{"disc" if list_type=="ul" else "decimal"};">'
                )
                list_stack.append(list_type)
        else:
            while list_stack:
                html_parts.append(f'</{list_stack.pop()}>')

        para_html = []
        has_visible_content = False

        # --- Texto e hipervínculos ---
        for child in para._p:
            tag = child.tag.split('}')[-1]

            if tag == "hyperlink":
                rId = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                href = None
                if rId and rId in para.part.rels:
                    href = para.part.rels[rId].target_ref

                inner_texts = []
                for r in child.findall(".//w:r", para._p.nsmap):
                    t = "".join(x.text or "" for x in r.findall(".//w:t", para._p.nsmap))
                    if t:
                        inner_texts.append(t)

                texto = "".join(inner_texts).strip()
                if href:
                    para_html.append(f'<a href="{html.escape(href)}">{html.escape(texto)}</a>')
                else:
                    para_html.append(html.escape(texto))
                has_visible_content = True

            elif tag == "r":
                txt = "".join(x.text or "" for x in child.findall(".//w:t", para._p.nsmap))
                if txt:
                    safe = html.escape(txt)
                    safe = convertir_urls_a_links(safe)
                    styles = []
                    rPr = child.find(".//w:rPr", para._p.nsmap)
                    if rPr is not None:
                        if rPr.find(".//w:b", para._p.nsmap) is not None:
                            styles.append("font-weight:bold")
                        if rPr.find(".//w:i", para._p.nsmap) is not None:
                            styles.append("font-style:italic")
                        if rPr.find(".//w:u", para._p.nsmap) is not None:
                            styles.append("text-decoration:underline")
                    style_attr = f' style="{";".join(styles)}"' if styles else ""
                    para_html.append(f'<span{style_attr}>{safe}</span>')
                    has_visible_content = True

        # --- Imágenes ---
        for drawing in para._p.findall(".//w:drawing", para._p.nsmap):
            blip = drawing.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
            if blip is not None:
                rId = blip.attrib.get(qn("r:embed"))
                if rId and rId in rels:
                    img_part = rels[rId]
                    ext = img_part.content_type.split("/")[-1]
                    img_b64 = base64.b64encode(img_part.blob).decode()
                    para_html.append(
                        f'<img src="data:image/{ext};base64,{img_b64}" '
                        f'style="max-width:100%; display:block; margin:6px 0;"/>'
                    )
                    has_visible_content = True

        joined = "".join(para_html).strip()

        if not has_visible_content:
            if not last_was_empty:
                html_parts.append("<p><br></p>")
                last_was_empty = True
            continue
        else:
            last_was_empty = False

        # --- Salida final ---
        if list_type:
            html_parts.append(f"<li>{joined}</li>")
        else:
            html_parts.append(f"<p>{joined}</p>")

    # --- Cierre de listas abiertas ---
    while list_stack:
        html_parts.append(f'</{list_stack.pop()}>')

    html_parts.append("</div>")
    log("✅ Conversión DOCX → HTML finalizada correctamente (v2.1)")
    return "".join(html_parts)

# ---- ODT a HTML (igual que antes, respetando listas y enlaces) ----
def odt_to_html_base64(odt_path):
    try:
        with zipfile.ZipFile(odt_path, 'r') as z:
            pics = {os.path.basename(n): z.read(n) for n in z.namelist() if n.startswith("Pictures/")}
            content_xml = z.read('content.xml').decode('utf-8')
    except Exception as e:
        raise
    root = ET.fromstring(content_xml)
    body = root.find('office:body', {'office':NS['office']})
    if body is None: body=root
    text_elem = body.find('office:text', {'office':NS['office']})
    if text_elem is None: text_elem=body
    html_parts = ['<div style="font-family: Arial, sans-serif; font-size:10pt; color:#000000;">']

    def process_paragraph(node):
        pieces = []
        if node.text: pieces.append(html.escape(node.text))
        for child in node:
            ttag = child.tag.split('}')[-1]
            if ttag=='span':
                txt = "".join(child.itertext())
                txt = html.escape(txt)
                style_name = child.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name') or child.attrib.get('text:style-name')
                open_tags=[]
                if style_name:
                    sn = style_name.lower()
                    if 'bold' in sn or 'negrita' in sn: open_tags.append('b')
                    if 'italic' in sn or 'cursiva' in sn: open_tags.append('i')
                    if 'underline' in sn or 'subray' in sn: open_tags.append('u')
                inner = txt
                for t in reversed(open_tags):
                    inner = f'<{t}>' + inner + f'</{t}>'
                pieces.append(inner)
                if child.tail: pieces.append(html.escape(child.tail))
            elif ttag=='a':
                href = child.attrib.get('{http://www.w3.org/1999/xlink}href') or child.attrib.get('xlink:href')
                link_text = "".join(child.itertext()) or href
                pieces.append(f'<a href="{html.escape(href)}">{html.escape(link_text)}</a>')
                if child.tail: pieces.append(html.escape(child.tail))
            elif ttag=='line-break':
                pieces.append("<br>")
            else:
                if child.text: pieces.append(html.escape(child.text))
                if child.tail: pieces.append(html.escape(child.tail))
        return "".join(pieces)

    for node in list(text_elem):
        tag_local = node.tag.split('}')[-1]
        if tag_local=='p':
            html_parts.append(f"<p>{process_paragraph(node)}</p>")
        elif tag_local=='list':
            list_tag="ul"
            items_html=[]
            for item in node.findall('text:list-item', NS):
                item_texts=[]
                for p in item.findall('text:p', NS):
                    item_texts.append(process_paragraph(p))
                items_html.append("<li>"+"<br>".join(item_texts)+"</li>")
            html_parts.append(f"<{list_tag}>"+"".join(items_html)+f"</{list_tag}>")
        elif tag_local=='frame':
            img_elem=node.find('.//draw:image', NS)
            if img_elem is not None:
                href=img_elem.attrib.get('{http://www.w3.org/1999/xlink}href')
                if href and href.startswith("Pictures/"):
                    img_name=href.split("/")[-1]
                    if img_name in pics:
                        ext=img_name.split('.')[-1].lower()
                        b64=base64.b64encode(pics[img_name]).decode()
                        html_parts.append(f'<img src="data:image/{ext};base64,{b64}" style="max-width:100%;"/><br>')
        else:
            for p in node.findall('.//text:p', NS):
                html_parts.append(f"<p>{process_paragraph(p)}</p>")
    html_parts.append("</div>")
    return "".join(html_parts)

# ---- Procesar plantilla (docx/odt) ----
def procesar_plantilla(path):
    try:
        if not os.path.exists(path) or os.path.getsize(path)==0:
            mostrar_notificacion("❌ Error", "El archivo está vacío o no existe", duracion=2)
            return
        ext = os.path.splitext(path)[1].lower()
        html_result = ""
        if ext==".odt":
            try:
                html_result = odt_to_html_base64(path)
            except Exception as e:
                mostrar_notificacion("⚠️ Error ODT", f"{e}", duracion=3)
                return
        elif ext==".docx":
            try:
                html_result = docx_to_html_base64(path)
            except Exception as e:
                mostrar_notificacion("⚠️ Error DOCX", f"{e}", duracion=3)
                return
        else:
            mostrar_notificacion("❌ Error", "Formato no soportado", duracion=2)
            return
        set_clipboard_html(html_result)
        mostrar_notificacion("✅ Éxito", f"'{os.path.basename(path)}' copiada al portapapeles", duracion=2)
    except Exception as e:
        mostrar_notificacion("❌ Error inesperado", f"{e}", duracion=3)

# ---- GUI ----
def crear_interfaz():
    if not os.path.exists(PLANTILLAS_DIR): os.makedirs(PLANTILLAS_DIR)
    root = tk.Tk()
    root.title(WINDOW_TITLE)

    # usamos settings ya cargado
    modo_oscuro = tk.BooleanVar(value=settings.get("modo_oscuro", False))

    try:
        if settings.get("x") is not None:
            geom_w = int(settings.get("width") or 220)
            geom_h = int(settings.get("height") or 250)
            geom_x = int(settings.get("x") or 0)
            geom_y = int(settings.get("y") or 0)
            root.geometry(f"{geom_w}x{geom_h}+{geom_x}+{geom_y}")
        else:
            archivos=[f for f in os.listdir(PLANTILLAS_DIR) if f.lower().endswith(('.odt', '.docx'))]
            height=min(len(archivos),10)*30+150
            height=max(height,250)
            width=220
            root.geometry(f"{width}x{height}")
    except Exception:
        pass

    def on_closing():
        try:
            root.update_idletasks()
            x=root.winfo_x()
            y=root.winfo_y()
            w=root.winfo_width()
            h=root.winfo_height()
            guardar_configuracion(modo_oscuro=modo_oscuro.get(), geom=(x,y,w,h))
        except Exception:
            pass
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_closing)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Vertical.TScrollbar", gripcount=0, background="#d9d9d9", troughcolor="#f0f0f0", arrowcolor="black")

    ctrl=tk.Frame(root,bg=root["bg"]); ctrl.pack(fill="x", pady=4)
    def crear_boton_control(texto, comando):
        return tk.Button(ctrl, text=texto, command=comando, bg=root["bg"], fg="black", relief="raised", bd=2, activebackground="#ddd", activeforeground="black")
    # 1. Creas el botón con la nueva lógica lambda
    btn_refrescar = crear_boton_control("Refrescar", lambda: listar_botones())
    
    # 2. Le dices que se coloque a la izquierda
    btn_refrescar.pack(side="left", padx=2)
    btn_abrir = crear_boton_control("Carpeta plantilla", lambda:os.startfile(PLANTILLAS_DIR)); btn_abrir.pack(side="left", padx=2)

    ajustes_menu = tk.Menubutton(ctrl, text="⚙️", relief="flat", bg=root["bg"], fg="black",
                                activebackground="#ddd", activeforeground="black")
    menu = tk.Menu(ajustes_menu, tearoff=0)

    # Toggle tema oscuro / claro
    def toggle_tema():
        modo_oscuro.set(not modo_oscuro.get())
        guardar_configuracion(modo_oscuro=modo_oscuro.get())
        listar_botones()  # recrea los botones con el nuevo tema
        aplicar_tema()    # aplica los colores a todos los elementos
    menu.add_command(label="Modo oscuro / claro", command=toggle_tema)

    # Activar/desactivar debug
    debug_var = tk.BooleanVar(value=settings.get("debug", False))
    def toggle_debug():
        val = debug_var.get()
        guardar_configuracion(debug=val)
    menu.add_separator()
    menu.add_checkbutton(label="Activar modo depuración (guardar log)",
                        command=toggle_debug,
                        variable=debug_var)

    # Opción para abrir el editor de plantillas
    def abrir_editor():
        try:
            vbs_path = os.path.join(BASE_DIR, "Editor", "lanzar_editor.vbs")
            if os.path.exists(vbs_path):
                os.startfile(vbs_path)
            else:
                messagebox.showerror("Error", f"No se encontró el archivo: {vbs_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el editor:\n{e}")

    menu.add_separator()
    menu.add_command(label="Iniciar generador de plantillas", command=abrir_editor)

    # Cerrar aplicación
    menu.add_separator()
    menu.add_command(label="Cerrar aplicación", command=root.destroy)

    ajustes_menu.configure(menu=menu)
    ajustes_menu.pack(side="left", padx=2)

    # menu.add_separ(); mratoenu.add_command(label="Cerrar aplicación",command=root.destroy)
    ajustes_menu.configure(menu=menu); ajustes_menu.pack(side="left", padx=2)

    separador=tk.Frame(root,height=2,bg="#888888"); separador.pack(fill="x", pady=(0,4))
    titulo_label=tk.Label(root,text="Plantillas",font=("Segoe UI",11),bg=root["bg"],fg="black"); titulo_label.pack(pady=4)

    # --- BARRA DE BÚSQUEDA ---
    search_frame = tk.Frame(root, bg=root["bg"])
    search_frame.pack(fill="x", padx=10, pady=5)
    
    tk.Label(search_frame, text="Buscar:", bg=root["bg"], fg="black", font=("Segoe UI", 10, "bold")).pack(side="left")   

    search_var = tk.StringVar()
    def al_escribir(*args):
        # Llama a listar_botones con lo que el usuario escribió
        listar_botones(search_var.get())
        
    search_var.trace_add("write", al_escribir)
    
    entry_busqueda = tk.Entry(search_frame, textvariable=search_var, font=("Segoe UI", 10))
    entry_busqueda.pack(side="left", fill="x", expand=True, padx=5)

        # Para que el fondo cambie, pero la palabra "Buscar:" sea SIEMPRE negra
    def actualizar_color_busqueda():
        # Eliminamos la línea que elegía 'white'
        for widget in search_frame.winfo_children():
            if isinstance(widget, tk.Label):
                # Forzamos fg="black" (color de letra) 
                # y mantenemos el bg (fondo) acorde al programa
                widget.configure(fg="black", bg=root["bg"])

    frame=tk.Frame(root,bg=root["bg"]); frame.pack(fill="both", expand=True, padx=4)
    canvas=tk.Canvas(frame,borderwidth=0,highlightthickness=0,bg=root["bg"])
    scroll=ttk.Scrollbar(frame,orient="vertical",command=canvas.yview,style="Vertical.TScrollbar")
    canvas.configure(yscrollcommand=scroll.set)
    inner=tk.Frame(canvas,bg=root["bg"])
    inner_id=canvas.create_window((0,0),window=inner,anchor="nw")
    def actualizar_scroll(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        try: canvas.itemconfig(inner_id,width=canvas.winfo_width())
        except Exception: pass
    inner.bind("<Configure>", actualizar_scroll)
    scroll.pack(side="right",fill="y"); canvas.pack(side="left",fill="both",expand=True)
    frame.update_idletasks(); canvas.configure(scrollregion=canvas.bbox("all"))
    def _on_mousewheel(event):
        try: canvas.yview_scroll(int(-1*(event.delta/120)),"units")
        except Exception: pass
    canvas.bind("<Enter>",lambda e:canvas.bind_all("<MouseWheel>",_on_mousewheel))
    canvas.bind("<Leave>",lambda e:canvas.unbind_all("<MouseWheel>"))

    def eliminar_acentos(texto):
        # Normaliza el texto a formato NFD (separa la letra del acento)
        texto = unicodedata.normalize('NFD', texto)
        # Filtra y se queda solo con los caracteres que no sean acentos (marcas)
        texto = "".join(c for c in texto if unicodedata.category(c) != 'Mn')
        return texto.lower()

    def listar_botones(filtro=""):
            for w in inner.winfo_children(): 
                w.destroy()
            
            btn_bg, btn_fg, active_bg = ("#3c3c3c", "white", "#555555") if modo_oscuro.get() else ("#f0f0f0", "black", "#ddd")
            dir_bg = "#4a4a4a" if modo_oscuro.get() else "#e1e1e1"

                # --- MODO BÚSQUEDA ---
            if filtro:
                encontrados = []
                filtro_limpio = eliminar_acentos(filtro) # <--- Limpiamos el filtro una vez
                
                for raiz, carpetas, archivos in os.walk(PLANTILLAS_DIR):
                    for f in archivos:
                        # Limpiamos el nombre del archivo para comparar
                        if f.lower().endswith(('.odt', '.docx')):
                            if filtro_limpio in eliminar_acentos(f):
                                encontrados.append(os.path.join(raiz, f))
                
                if not encontrados:
                    tk.Label(inner, text="No se encontraron coincidencias", fg="gray", bg=inner["bg"]).pack(pady=10)
                
                for ruta_completa in sorted(encontrados):
                    nombre_archivo = os.path.basename(ruta_completa)
                    nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
                    
                    btn_file = tk.Button(inner, text=nombre_sin_ext, font=("Segoe UI", 10),
                                    height=2, width=22, bg=btn_bg, fg=btn_fg,
                                    command=lambda r=ruta_completa: procesar_plantilla(r),
                                    relief="raised", bd=2, activebackground=active_bg)
                    btn_file.pack(pady=2, anchor="center", fill="x")
                return # Salimos para no ejecutar la lógica de carpetas normal

            # --- MODO NAVEGACIÓN NORMAL (si no hay filtro) ---
            global ruta_navegacion_actual
            if os.path.abspath(ruta_navegacion_actual) != os.path.abspath(PLANTILLAS_DIR):
                def volver():
                    global ruta_navegacion_actual
                    ruta_navegacion_actual = os.path.dirname(ruta_navegacion_actual)
                    listar_botones()
                    
                b_volver = tk.Button(inner, text="⬅ Volver atrás", font=("Segoe UI", 10, "bold"),
                                height=1, width=20, bg="#d9534f", fg="white", 
                                command=volver, relief="flat")
                b_volver.pack(pady=(0, 10), fill="x")

            try:
                contenido = os.listdir(ruta_navegacion_actual)
                carpetas = [f for f in contenido if os.path.isdir(os.path.join(ruta_navegacion_actual, f))]
                archivos = [f for f in contenido if f.lower().endswith(('.odt', '.docx'))]

                for c in sorted(carpetas):
                    def abrir_carpeta(nombre=c):
                        global ruta_navegacion_actual
                        ruta_navegacion_actual = os.path.join(ruta_navegacion_actual, nombre)
                        listar_botones()
                    tk.Button(inner, text=f"📁 {c}", font=("Segoe UI", 10, "bold"), height=2, width=22, 
                            bg=dir_bg, fg=btn_fg, command=abrir_carpeta, relief="raised", bd=2).pack(pady=2, fill="x")

                for a in sorted(archivos):
                    nombre_sin_ext = os.path.splitext(a)[0]
                    ruta_hijo = os.path.join(ruta_navegacion_actual, a)
                    tk.Button(inner, text=f"📄 {nombre_sin_ext}", font=("Segoe UI", 10), height=2, width=22, 
                            bg=btn_bg, fg=btn_fg, command=lambda r=ruta_hijo: procesar_plantilla(r),
                            relief="raised", bd=2).pack(pady=2, fill="x")
            except Exception as e:
                tk.Label(inner, text=f"Error: {e}", fg="red").pack()

    def aplicar_tema():
        bg_color = "#2e2e2e" if modo_oscuro.get() else "SystemButtonFace"
        fg_color = "white" if modo_oscuro.get() else "black"
        canvas_bg = "#333333" if modo_oscuro.get() else "#f0f0f0"  # fondo del canvas

        try: 
            actualizar_color_busqueda() 
        except: 
            pass

        # Root y frames principales
        root.configure(bg=bg_color)
        ctrl.configure(bg=bg_color)
        frame.configure(bg=bg_color)
        inner.configure(bg=canvas_bg)
        canvas.configure(bg=canvas_bg)
        titulo_label.configure(bg=bg_color, fg=fg_color)

        # Scrollbar
        style.configure("Vertical.TScrollbar",
                        background="#4a4a4a" if modo_oscuro.get() else "#d9d9d9",
                        troughcolor="#2e2e2e" if modo_oscuro.get() else "#f0f0f0",
                        arrowcolor=fg_color)

        # Botones de control
        for b in [btn_refrescar, btn_abrir, ajustes_menu]:
            b.configure(bg=bg_color, fg=fg_color, 
                        activebackground="#555555" if modo_oscuro.get() else "#ddd",
                        activeforeground=fg_color)

        # Botones y etiquetas dentro de inner
        for w in inner.winfo_children():
            if isinstance(w, tk.Label):
                w.configure(bg=canvas_bg, fg=fg_color)
            elif isinstance(w, tk.Button):
                w.configure(bg="#f0f0f0" if not modo_oscuro.get() else "#3c3c3c",
                            fg="black" if not modo_oscuro.get() else "white",
                            activebackground="#ddd" if not modo_oscuro.get() else "#555555",
                            activeforeground="black" if not modo_oscuro.get() else "white")
                


    listar_botones()
    aplicar_tema()
    root.mainloop()

if __name__=="__main__":
    
    # Variable para rastrear la carpeta actual
    ruta_navegacion_actual = PLANTILLAS_DIR
    
    crear_interfaz()
