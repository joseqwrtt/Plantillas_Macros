# interfaz_ctk.py  ·  v3.1
# Interfaz moderna con CustomTkinter + Popup nativo Win32 sin foco
# Bloqueo real de teclas de navegación cuando el popup está visible
# Requiere: pip install customtkinter

import os
import sys
import time
import threading
import ctypes
import ctypes.wintypes as wintypes
import tkinter as tk
import tkinter.font as tkfont

try:
    import customtkinter as ctk
except ImportError:
    print("❌ Falta customtkinter. Instálalo con: pip install customtkinter")
    sys.exit(1)

from interfaz_base import InterfazUsuario

# ══════════════════════════════════════════════════════════
#  CONSTANTES WIN32
# ══════════════════════════════════════════════════════════
WS_EX_NOACTIVATE   = 0x08000000
WS_EX_TOPMOST      = 0x00000008
WS_EX_TOOLWINDOW   = 0x00000080
WS_EX_LAYERED      = 0x00080000
GWL_EXSTYLE        = -20
SWP_NOMOVE         = 0x0002
SWP_NOSIZE         = 0x0001
SWP_NOACTIVATE     = 0x0010
SWP_SHOWWINDOW     = 0x0040
HWND_TOPMOST       = -1

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32



# ── Paletas de colores por tema ────────────────────────────────────────────
# ══════════════════════════════════════════════════════════
#  POPUP NATIVO SIN FOCO (Win32 + Tkinter canvas)
# ══════════════════════════════════════════════════════════
class PopupSinFoco:
    """
    Popup de selección de plantillas que NO roba el foco.
    Usa WS_EX_NOACTIVATE + WS_EX_TOOLWINDOW via ctypes.
    """

    ALTO_ITEM       = 26
    ANCHO_MIN       = 280
    PAD_X           = 10
    ITEMS_VISIBLES  = 20          # máx de items sin scrollbar
    ANCHO_SCROLL    = 8           # ancho de la barra de scroll
    COLOR_BG    = "#1e1e2e"
    COLOR_SEL   = "#3b82f6"
    COLOR_TEXT  = "#cdd6f4"
    COLOR_SUB   = "#6c7086"
    COLOR_BORDE = "#313244"
    RADIO       = 8           # radio esquinas del borde

    def __init__(self, gestor, on_seleccion):
        self.gestor       = gestor
        self.on_seleccion = on_seleccion
        self.ventana      = None
        self.canvas       = None
        self.items        = []
        self.indice_sel   = 0
        self._hwnd           = None
        self._lock           = threading.Lock()
        self._activo         = False
        self._paleta         = {"popup_bg":"#1e1e2e","popup_sel":"#3b82f6","popup_text":"#cdd6f4","popup_sub":"#6c7086","popup_borde":"#313244"}
        self._scroll_offset  = 0
        self._scrollbar      = None

    def aplicar_tema(self, modo):
        """Actualiza los colores del popup según el tema activo."""
        self._paleta = {"popup_bg":"#1e1e2e","popup_sel":"#3b82f6","popup_text":"#cdd6f4","popup_sub":"#6c7086","popup_borde":"#313244"}
        # Redibujar si el popup está visible
        if self._ventana_existe():
            bg = self._paleta["popup_bg"]
            self.ventana.configure(bg=bg)
            if self.canvas:
                self.canvas.configure(bg=bg)
            self._dibujar()

    # ── Crear ventana ──────────────────────────────────────
    def _crear_ventana(self, x, y):
        """Crea la ventana Tk sin decoración y aplica flags Win32."""
        self._scroll_offset = 0   # índice del primer item visible

        v = tk.Toplevel()
        v.withdraw()
        v.overrideredirect(True)
        v.attributes("-topmost", True)
        v.configure(bg=self.COLOR_BG)
        v.resizable(False, False)

        # Frame contenedor (canvas + scrollbar opcional)
        self._frame_popup = tk.Frame(v, bg=self.COLOR_BG, bd=0)
        self._frame_popup.pack(fill="both", expand=True)

        # Canvas principal
        self.canvas = tk.Canvas(
            self._frame_popup, bg=self.COLOR_BG,
            highlightthickness=0, bd=0, cursor="arrow"
        )
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<ButtonRelease-1>", self._click_canvas)
        self.canvas.bind("<Motion>",          self._hover_canvas)
        self.canvas.bind("<MouseWheel>",      self._on_mousewheel)

        # Scrollbar (se muestra solo si hay más de ITEMS_VISIBLES items)
        self._scrollbar = tk.Scrollbar(
            self._frame_popup, orient="vertical",
            width=self.ANCHO_SCROLL,
            command=self._on_scroll_cmd,
            bg="#313244", troughcolor="#1e1e2e",
            activebackground="#3b82f6", relief="flat", bd=0
        )

        self.ventana = v
        v.update_idletasks()

        hwnd = self._get_hwnd()
        if hwnd:
            self._hwnd = hwnd
            ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ex_style |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)
            user32.SetWindowPos(
                hwnd, HWND_TOPMOST,
                x, y, 1, 1,
                SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW
            )

        return v

    def _on_mousewheel(self, event):
        """Desplaza la lista con la rueda del ratón."""
        delta = -1 if event.delta > 0 else 1
        self._scroll_by(delta)

    def _on_scroll_cmd(self, *args):
        """Comando recibido desde la scrollbar."""
        if args[0] == "moveto":
            total = len(self.items)
            visible = min(total, self.ITEMS_VISIBLES)
            new_offset = int(float(args[1]) * total)
            self._scroll_offset = max(0, min(new_offset, total - visible))
        elif args[0] == "scroll":
            self._scroll_by(int(args[1]))
        self._dibujar()

    def _scroll_by(self, delta):
        total   = len(self.items)
        visible = min(total, self.ITEMS_VISIBLES)
        self._scroll_offset = max(0, min(self._scroll_offset + delta, total - visible))
        # Mantener la selección visible
        if self.indice_sel < self._scroll_offset:
            self.indice_sel = self._scroll_offset
        elif self.indice_sel >= self._scroll_offset + visible:
            self.indice_sel = self._scroll_offset + visible - 1
        self._dibujar()

    def _actualizar_scrollbar(self, total, visible):
        """Muestra u oculta la scrollbar y actualiza su posición."""
        if total > self.ITEMS_VISIBLES:
            self._scrollbar.pack(side="right", fill="y")
            lo = self._scroll_offset / total
            hi = (self._scroll_offset + visible) / total
            self._scrollbar.set(lo, hi)
        else:
            self._scrollbar.pack_forget()

    def _get_hwnd(self):
        """Obtiene el HWND de la ventana Tk."""
        try:
            return self.ventana.winfo_id()
        except Exception:
            return None

    # ── API pública ────────────────────────────────────────
    def mostrar(self, filtro, x, y):
        """Muestra el popup en (x, y) con los resultados para 'filtro'."""
        items = self.gestor.buscar_plantillas(filtro)
        if not items:
            self.cerrar()
            return

        with self._lock:
            max_items = self.gestor.preferencias.get("max_resultados_popup", 100)
            self.items      = items[:max_items]
            self.indice_sel = 0

        if self.ventana is None or not self._ventana_existe():
            self.ventana = None
            self._crear_y_dibujar(x, y)
        else:
            self._dibujar()
            self._mover(x, y)

        self._activo = True

        # Activar hotkeys suprimidas solo cuando el popup es visible
        if hasattr(self.gestor, "activar_hotkeys_popup"):
            self.gestor.activar_hotkeys_popup()

    def _crear_y_dibujar(self, x, y):
        self._crear_ventana(x, y)
        self._dibujar()
        ancho, alto_visible, alto_total, necesita_scroll = self._calcular_dimensiones()
        # Ancho real incluye scrollbar si es necesario
        ancho_real = ancho + (self.ANCHO_SCROLL if necesita_scroll else 0)
        sw = user32.GetSystemMetrics(0)
        sh = user32.GetSystemMetrics(1)
        if x + ancho_real > sw:
            x = sw - ancho_real - 4
        if y + alto_visible > sh:
            y = y - alto_visible - 4

        self.ventana.geometry(f"{ancho_real}x{alto_visible}+{x}+{y}")
        self.ventana.deiconify()
        if self._hwnd:
            user32.SetWindowPos(
                self._hwnd, HWND_TOPMOST,
                x, y, ancho_real, alto_visible,
                SWP_NOACTIVATE | SWP_SHOWWINDOW
            )

    def actualizar_filtro(self, filtro):
        """Actualiza resultados con nuevo filtro (llamado desde hilo detector)."""
        items = self.gestor.buscar_plantillas(filtro)
        if not items:
            self.cerrar()
            return
        with self._lock:
            max_items = self.gestor.preferencias.get("max_resultados_popup", 100)
            self.items      = items[:max_items]
            self.indice_sel = 0
        if self.ventana and self._ventana_existe():
            try:
                self.ventana.after(0, self._redibujar_seguro)
            except Exception:
                pass

    def _redibujar_seguro(self):
        if self._ventana_existe():
            self._scroll_offset = 0   # reset scroll al actualizar filtro
            self._dibujar()
            ancho, alto_visible, alto_total, necesita_scroll = self._calcular_dimensiones()
            ancho_real = ancho + (self.ANCHO_SCROLL if necesita_scroll else 0)
            x = self.ventana.winfo_x()
            y = self.ventana.winfo_y()
            self.ventana.geometry(f"{ancho_real}x{alto_visible}+{x}+{y}")
            if self._hwnd:
                user32.SetWindowPos(
                    self._hwnd, HWND_TOPMOST,
                    x, y, ancho_real, alto_visible,
                    SWP_NOACTIVATE | SWP_SHOWWINDOW
                )

    def mover_sel(self, delta):
        """Mueve la selección arriba (-1) o abajo (+1) y arrastra el scroll si hace falta."""
        with self._lock:
            n = len(self.items)
            if n == 0:
                return
            self.indice_sel = (self.indice_sel + delta) % n
        # Ajustar scroll para mantener el item seleccionado visible
        visible = min(n, self.ITEMS_VISIBLES)
        offset  = getattr(self, "_scroll_offset", 0)
        if self.indice_sel < offset:
            self._scroll_offset = self.indice_sel
        elif self.indice_sel >= offset + visible:
            self._scroll_offset = self.indice_sel - visible + 1
        if self.ventana and self._ventana_existe():
            try:
                self.ventana.after(0, self._dibujar)
            except Exception:
                pass

    def confirmar(self):
        """Ejecuta la plantilla seleccionada."""
        with self._lock:
            if not self.items or self.indice_sel >= len(self.items):
                return None
            nombre, ruta = self.items[self.indice_sel]
        self.cerrar()
        self.on_seleccion(ruta)
        return ruta

    def cerrar(self):
        """Cierra el popup."""
        self._activo = False
        # Desactivar hotkeys de navegación al cerrar
        if hasattr(self.gestor, "desactivar_hotkeys_popup"):
            self.gestor.desactivar_hotkeys_popup()
        if self.ventana:
            try:
                self.ventana.after(0, self._destruir)
            except Exception:
                pass

    def _destruir(self):
        try:
            if self._ventana_existe():
                self.ventana.destroy()
        except Exception:
            pass
        self.ventana = None
        self._hwnd   = None
        self.canvas  = None

    def esta_activo(self):
        return self._activo and self._ventana_existe()

    def _ventana_existe(self):
        try:
            return self.ventana is not None and self.ventana.winfo_exists()
        except Exception:
            return False

    # ── Dibujo ─────────────────────────────────────────────
    def _calcular_dimensiones(self):
        font_nombre  = tkfont.Font(family="Segoe UI", size=10)
        font_carpeta = tkfont.Font(family="Segoe UI", size=8)
        FIJO = self.PAD_X * 2 + 22 + 8
        max_ancho = 0
        for nombre, ruta in self.items:
            carpeta = os.path.basename(os.path.dirname(ruta))
            carpeta_visible = carpeta and carpeta.lower() not in ("plantillas", "")
            w = font_nombre.measure(nombre) + FIJO
            if carpeta_visible:
                w += font_carpeta.measure(carpeta) + 14
            if w > max_ancho:
                max_ancho = w
        necesita_scroll = len(self.items) > self.ITEMS_VISIBLES
        if necesita_scroll:
            max_ancho += self.ANCHO_SCROLL + 2   # espacio para la scrollbar
        ancho       = max(max_ancho, self.ANCHO_MIN)
        n_visibles  = min(len(self.items), self.ITEMS_VISIBLES)
        alto_visible = n_visibles * self.ALTO_ITEM + 8
        alto_total   = len(self.items) * self.ALTO_ITEM + 8
        return ancho, alto_visible, alto_total, necesita_scroll

    def _dibujar(self):
        if not self._ventana_existe():
            return
        c = self.canvas
        c.delete("all")
        with self._lock:
            items  = list(self.items)
            sel    = self.indice_sel

        total   = len(items)
        visible = min(total, self.ITEMS_VISIBLES)
        offset  = getattr(self, "_scroll_offset", 0)
        offset  = max(0, min(offset, total - visible))
        self._scroll_offset = offset

        # Items visibles en esta "ventana" de scroll
        items_vis = items[offset: offset + visible]

        ancho, alto_visible, alto_total, necesita_scroll = self._calcular_dimensiones()
        c.configure(width=ancho, height=alto_visible)

        p = self._paleta
        c.configure(bg=p["popup_bg"])

        # Fondo
        self._rect_redondeado(c, 0, 0, ancho-1, alto_visible-1, self.RADIO,
                               fill=p["popup_bg"], outline=p["popup_borde"], width=1)

        FONT_NOMBRE  = tkfont.Font(family="Segoe UI", size=10)
        FONT_CARPETA = tkfont.Font(family="Segoe UI", size=8)

        for i, (nombre, ruta) in enumerate(items_vis):
            idx_real = offset + i   # índice real en self.items
            y0 = 4 + i * self.ALTO_ITEM
            y1 = y0 + self.ALTO_ITEM - 2

            if idx_real == sel:
                self._rect_redondeado(c, 4, y0, ancho-5, y1, 5,
                                       fill=p["popup_sel"], outline="")
                color_txt = "#ffffff"
            else:
                color_txt = p["popup_text"]

            ext   = os.path.splitext(ruta)[1].lower()
            icono = "📄" if ext == ".docx" else "📘"
            c.create_text(
                self.PAD_X + 2, y0 + self.ALTO_ITEM // 2,
                text=icono, anchor="w",
                font=("Segoe UI Emoji", 9), fill=color_txt
            )

            carpeta         = os.path.basename(os.path.dirname(ruta))
            carpeta_visible = carpeta and carpeta.lower() not in ("plantillas", "")
            x_nombre_ini    = self.PAD_X + 22
            x_nombre_max    = ancho - self.PAD_X

            if carpeta_visible:
                ancho_carpeta = FONT_CARPETA.measure(carpeta) + 8
                x_nombre_max  = ancho - self.PAD_X - ancho_carpeta - 6
                c.create_text(
                    ancho - self.PAD_X, y0 + self.ALTO_ITEM // 2,
                    text=carpeta, anchor="e",
                    font=("Segoe UI", 8), fill=p["popup_sub"]
                )

            espacio      = x_nombre_max - x_nombre_ini
            texto_nombre = nombre
            if FONT_NOMBRE.measure(texto_nombre) > espacio:
                while texto_nombre and FONT_NOMBRE.measure(texto_nombre + "…") > espacio:
                    texto_nombre = texto_nombre[:-1]
                texto_nombre += "…"

            c.create_text(
                x_nombre_ini, y0 + self.ALTO_ITEM // 2,
                text=texto_nombre, anchor="w",
                font=("Segoe UI", 10), fill=color_txt
            )

        # Scrollbar
        self._actualizar_scrollbar(total, visible)

    def _rect_redondeado(self, canvas, x1, y1, x2, y2, r, **kwargs):
        points = [
            x1+r, y1,  x2-r, y1,
            x2,   y1,  x2,   y1+r,
            x2,   y2-r,x2,   y2,
            x2-r, y2,  x1+r, y2,
            x1,   y2,  x1,   y2-r,
            x1,   y1+r,x1,   y1,
            x1+r, y1,
        ]
        return canvas.create_polygon(points, smooth=True, **kwargs)

    def _mover(self, x, y):
        if self._ventana_existe():
            self.ventana.geometry(f"+{x}+{y}")

    # ── Eventos del canvas ─────────────────────────────────
    def _y_to_indice(self, y_canvas):
        offset   = getattr(self, "_scroll_offset", 0)
        visible  = min(len(self.items), self.ITEMS_VISIBLES)
        i_local  = max(0, min((y_canvas - 4) // self.ALTO_ITEM, visible - 1))
        return offset + i_local

    def _click_canvas(self, event):
        with self._lock:
            self.indice_sel = self._y_to_indice(event.y)
        self.confirmar()

    def _hover_canvas(self, event):
        idx = self._y_to_indice(event.y)
        with self._lock:
            self.indice_sel = idx
        self._dibujar()


# ══════════════════════════════════════════════════════════
#  PANEL PRINCIPAL (CustomTkinter)
# ══════════════════════════════════════════════════════════
class InterfazCTK(InterfazUsuario):
    """
    Interfaz principal moderna con CustomTkinter.
    Panel lateral con lista de plantillas, búsqueda y ajustes.
    """

    TITULO  = "Plantillas Macro"
    VERSION = "3.0"

    def __init__(self):
        ctk.set_appearance_mode("dark")   # solo modo oscuro
        ctk.set_default_color_theme("blue")
        self.gestor           = None
        self.app              = None
        self.popup            = None
        self._ultimo_filtro   = ""
        self._ultimo_popup_xy = (0, 0)
        self._lista_items     = []
        self._ttk_style       = None
        self._modo_simple     = False   # False = completo, True = simple
        self._ruta_actual     = None    # carpeta navegada actualmente

    def ejecutar(self, gestor):
        self.gestor = gestor
        gestor.interfaz = self

        self.app = ctk.CTk()
        self.app.title(self.TITULO)
        self.app.geometry("820x580")
        self.app.minsize(680, 480)
        self.app.protocol("WM_DELETE_WINDOW", self._salir)

        # Icono (si existe)
        try:
            ico = os.path.join(os.path.dirname(__file__), "icon.ico")
            if os.path.exists(ico):
                self.app.iconbitmap(ico)
        except Exception:
            pass

        self._construir_ui()
        # Ruta raíz de navegación
        self._ruta_actual = self.gestor.preferencias.get("ruta_plantillas", "")
        self._cargar_lista_plantillas()

        # Popup sin foco
        self.popup = PopupSinFoco(gestor, self._on_plantilla_seleccionada)
        gestor.set_popup_activo(self.popup)

        # Actualizar banner de licencia en sidebar
        self._actualizar_banner_licencia()

        self.app.mainloop()

    def mostrar_popup_comandos(self, filtro, x=None, y=None):
        if x is None or y is None:
            pt = wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(pt))
            x, y = pt.x, pt.y + 20
        self._ultimo_filtro   = filtro
        self._ultimo_popup_xy = (x, y)
        self.app.after(0, lambda: self.popup.mostrar(filtro, x, y))

    def actualizar_popup_comandos(self, filtro):
        self._ultimo_filtro = filtro
        self.app.after(0, lambda: self.popup.actualizar_filtro(filtro))

    def mostrar_popup_desde_detector(self, filtro, x=None, y=None):
        self.mostrar_popup_comandos(filtro, x, y)

    def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
        from tkinter import messagebox
        if tipo == "error":
            messagebox.showerror(titulo, mensaje)
        elif tipo == "warning":
            messagebox.showwarning(titulo, mensaje)
        else:
            messagebox.showinfo(titulo, mensaje)

    def preguntar(self, titulo, mensaje):
        from tkinter import messagebox
        return messagebox.askyesno(titulo, mensaje)

    def seleccionar_carpeta(self, titulo, directorio_inicial):
        from tkinter import filedialog
        return filedialog.askdirectory(title=titulo, initialdir=directorio_inicial)

    def recargar_preferencias(self):
        # Resetear navegación a la nueva raíz
        nueva_raiz = self.gestor.preferencias.get("ruta_plantillas", "")
        self._ruta_actual = nueva_raiz
        self.app.after(0, self._actualizar_campo_ruta)

    def recargar_contenido(self):
        self.app.after(0, self._cargar_lista_plantillas)

    # ── Construcción de la UI ──────────────────────────────
    def _construir_ui(self):
        self.app.grid_columnconfigure(0, weight=0)   # sidebar fijo
        self.app.grid_columnconfigure(1, weight=1)   # contenido
        self.app.grid_rowconfigure(0, weight=1)

        self._construir_sidebar()
        self._construir_area_principal()

    # ── Modo simple ───────────────────────────────────────
    def _toggle_modo_simple(self):
        if self._modo_simple:
            self._salir_modo_simple()
        else:
            self._entrar_modo_simple()

    def _entrar_modo_simple(self):
        self._modo_simple = True
        self._geom_completa = self.app.geometry()

        # Ocultar sidebar
        self._sidebar.grid_remove()
        self.app.grid_columnconfigure(0, weight=0, minsize=0)
        self.app.grid_columnconfigure(1, weight=1)

        # Sin bordes ni padding en el frame
        self._frame_plantillas.configure(corner_radius=0, fg_color="transparent")
        self._lista_canvas.configure(bg="#1c1c2e")
        self._lista_frame.configure(bg="#1c1c2e")

        # Quitar padding de la barra de búsqueda y reducir
        self._search_bar.grid_configure(padx=4, pady=(4, 2))
        self._entry_busqueda.configure(height=28)

        # Ocultar barra de estado del modo completo
        self._lbl_estado.grid_remove()
        # Mostrar barra de estado del modo simple
        self._lbl_estado_simple.grid()

        # Botón cambia a "volver"
        self._btn_modo_simple.configure(text="⊞")

        # Resetear navegación al entrar en modo simple
        self._ruta_actual = self.gestor.preferencias.get("ruta_plantillas", "")
        self._vista_plantillas()
        self._cargar_lista_plantillas()

        self.app.minsize(180, 120)
        self.app.geometry("240x500")
        self.app.attributes("-topmost", True)
        self.app.resizable(True, True)

    def _salir_modo_simple(self):
        self._modo_simple = False

        # Restaurar sidebar
        self._sidebar.grid(row=0, column=0, sticky="nsew")
        self.app.grid_columnconfigure(0, weight=0)

        # Restaurar bordes y padding
        self._frame_plantillas.configure(corner_radius=12,
                                          fg_color=("gray86", "gray17"))
        _bg2 = "#1c1c2e"
        self._lista_canvas.configure(bg=_bg2)
        self._lista_frame.configure(bg=_bg2)
        self._search_bar.grid_configure(padx=8, pady=(8, 4))
        self._entry_busqueda.configure(height=32)

        # Restaurar barra de estado modo completo
        self._lbl_estado.grid()
        # Ocultar barra de estado modo simple
        self._lbl_estado_simple.configure(text="")
        self._lbl_estado_simple.grid_remove()

        # Restaurar botón
        self._btn_modo_simple.configure(text="⊡")

        # Recargar en modo completo
        self._cargar_lista_plantillas()

        self.app.attributes("-topmost", False)
        self.app.minsize(680, 480)
        self.app.resizable(True, True)
        if hasattr(self, "_geom_completa") and self._geom_completa:
            self.app.geometry(self._geom_completa)
        else:
            self.app.geometry("820x580")

    # ── Sidebar ────────────────────────────────────────────
    def _construir_sidebar(self):
        sb = ctk.CTkFrame(self.app, width=200, corner_radius=0,
                          fg_color=("#e2e5f5", "#0f0f1a"))
        sb.grid(row=0, column=0, sticky="nsew")
        self._sidebar = sb   # referencia para mostrar/ocultar en modo simple
        sb.grid_propagate(False)
        # row 5 = espacio flexible (empuja todo lo demás hacia abajo)
        sb.grid_rowconfigure(5, weight=1)

        # Logo / título
        ctk.CTkLabel(
            sb, text="📋", font=ctk.CTkFont(size=32)
        ).grid(row=0, column=0, pady=(24, 0), padx=20, sticky="w")

        ctk.CTkLabel(
            sb, text=self.TITULO,
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=1, column=0, pady=(4, 20), padx=20, sticky="w")

        # Botones de navegación  (filas 2, 3, 4)
        self._btn_plantillas = self._nav_btn(sb, "🗂  Plantillas",  2, self._vista_plantillas)
        self._btn_ajustes    = self._nav_btn(sb, "⚙️  Ajustes",      3, self._vista_ajustes)
        self._btn_acerca     = self._nav_btn(sb, "ℹ️  Acerca de",    4, self._vista_acerca)

        # ── Zona inferior (fija): abrir carpeta, licencia, detector, versión ──

        # Botón abrir carpeta  (row 6)
        def _abrir_carpeta():
            ruta = self.gestor.preferencias.get("ruta_plantillas", "")
            if ruta and os.path.exists(ruta):
                os.startfile(ruta)
            else:
                self.mostrar_mensaje("Carpeta no encontrada", "La carpeta de plantillas no existe.\nConfigúrala en Ajustes.", "warning")

        ctk.CTkButton(
            sb, text="📁  Abrir plantillas", anchor="w",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            hover_color=("#2d2d4e", "#1e1e3a"),
            command=_abrir_carpeta, height=34, corner_radius=8
        ).grid(row=6, column=0, padx=10, pady=(0, 2), sticky="ew")

        # Banner de licencia  (row 7)
        self._lbl_licencia = ctk.CTkLabel(
            sb, text="",
            font=ctk.CTkFont(size=10),
            text_color="gray", cursor="hand2"
        )
        self._lbl_licencia.grid(row=7, column=0, pady=(0, 2), padx=16, sticky="w")
        self._lbl_licencia.bind("<Button-1>", lambda e: self._vista_acerca())

        # Estado del detector  (row 8)
        self._lbl_detector = ctk.CTkLabel(
            sb, text="● Detector activo",
            font=ctk.CTkFont(size=11),
            text_color="#4ade80"
        )
        self._lbl_detector.grid(row=8, column=0, pady=(0, 4), padx=16, sticky="w")

        ctk.CTkLabel(
            sb, text=f"v{self.VERSION}",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        ).grid(row=9, column=0, pady=(0, 16), padx=16, sticky="w")

    def _nav_btn(self, parent, texto, row, cmd):
        btn = ctk.CTkButton(
            parent, text=texto, anchor="w",
            font=ctk.CTkFont(size=13),
            fg_color="transparent",
            hover_color=("#2d2d4e", "#1e1e3a"),
            command=cmd, height=38, corner_radius=8
        )
        btn.grid(row=row, column=0, padx=10, pady=2, sticky="ew")
        return btn

    # ── Área principal ─────────────────────────────────────
    def _construir_area_principal(self):
        self._frame_main = ctk.CTkFrame(self.app, corner_radius=0,
                                         fg_color="transparent")
        self._frame_main.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self._frame_main.grid_rowconfigure(0, weight=1)
        self._frame_main.grid_columnconfigure(0, weight=1)

        self._frame_plantillas = self._crear_frame_plantillas()
        self._frame_ajustes    = self._crear_frame_ajustes()
        self._frame_acerca     = self._crear_frame_acerca()

        self._vista_plantillas()

    def _mostrar_frame(self, frame):
        for f in (self._frame_plantillas, self._frame_ajustes, self._frame_acerca):
            f.grid_remove()
        frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)

    def _vista_plantillas(self):
        self._seleccionar_nav(self._btn_plantillas)
        self._mostrar_frame(self._frame_plantillas)

    def _vista_ajustes(self):
        self._seleccionar_nav(self._btn_ajustes)
        self._mostrar_frame(self._frame_ajustes)
        self._actualizar_campo_ruta()

    def _vista_acerca(self):
        self._seleccionar_nav(self._btn_acerca)
        self._mostrar_frame(self._frame_acerca)

    def _seleccionar_nav(self, btn_activo):
        for b in (self._btn_plantillas, self._btn_ajustes, self._btn_acerca):
            b.configure(
                fg_color=("#3b82f6", "#2563eb") if b is btn_activo else "transparent"
            )

    # ── Frame: Plantillas ──────────────────────────────────
    def _crear_frame_plantillas(self):
        f = ctk.CTkFrame(self._frame_main, corner_radius=12)

        f.grid_rowconfigure(1, weight=1)   # lista ocupa todo
        f.grid_columnconfigure(0, weight=1)

        # ── Fila única: buscador + recarga + modo simple ──
        self._search_bar = ctk.CTkFrame(f, fg_color="transparent")
        self._search_bar.grid(row=0, column=0, padx=8, pady=(8, 4), sticky="ew")
        self._search_bar.grid_columnconfigure(0, weight=1)

        self._var_busqueda = tk.StringVar()
        self._var_busqueda.trace_add("write", self._on_busqueda_cambio)
        self._entry_busqueda = ctk.CTkEntry(
            self._search_bar,
            placeholder_text="🔍  Buscar...",
            textvariable=self._var_busqueda,
            height=32, corner_radius=8
        )
        self._entry_busqueda.grid(row=0, column=0, sticky="ew")

        btn_reload = ctk.CTkButton(
            self._search_bar, text="↺", width=32, height=32,
            corner_radius=8, font=ctk.CTkFont(size=15),
            fg_color="transparent", border_width=1,
            command=self._cargar_lista_plantillas
        )
        btn_reload.grid(row=0, column=1, padx=(4, 0))

        self._btn_modo_simple = ctk.CTkButton(
            self._search_bar, text="⊡", width=32, height=32,
            corner_radius=8, font=ctk.CTkFont(size=14),
            fg_color="transparent", border_width=1,
            command=self._toggle_modo_simple
        )
        self._btn_modo_simple.grid(row=0, column=2, padx=(4, 0))

        # ── Lista: canvas + frame tk puro (sin parpadeo) ──
        _bg = "#1c1c2e"
        lista_outer = tk.Frame(f, bg=_bg)
        lista_outer.grid(row=1, column=0, padx=8, pady=(0, 4), sticky="nsew")
        lista_outer.grid_rowconfigure(0, weight=1)
        lista_outer.grid_columnconfigure(0, weight=1)

        self._lista_canvas = tk.Canvas(
            lista_outer, bg=_bg,
            highlightthickness=0, bd=0
        )
        self._lista_scroll = tk.Scrollbar(
            lista_outer, orient="vertical",
            command=self._lista_canvas.yview,
            bg="#313244", troughcolor="#1c1c2e",
            activebackground="#60a5fa", relief="flat", bd=0, width=8
        )
        self._lista_canvas.configure(yscrollcommand=self._lista_scroll.set)
        self._lista_canvas.grid(row=0, column=0, sticky="nsew")
        self._lista_scroll.grid(row=0, column=1, sticky="ns")

        self._lista_frame = tk.Frame(self._lista_canvas, bg=_bg)
        self._lista_frame_id = self._lista_canvas.create_window(
            (0, 0), window=self._lista_frame, anchor="nw"
        )

        def _on_frame_configure(e):
            self._lista_canvas.configure(
                scrollregion=self._lista_canvas.bbox("all")
            )
        def _on_canvas_configure(e):
            self._lista_canvas.itemconfig(
                self._lista_frame_id, width=e.width
            )
        def _on_mousewheel(e):
            self._lista_canvas.yview_scroll(int(-1*(e.delta/120)), "units")

        self._lista_frame.bind("<Configure>", _on_frame_configure)
        self._lista_canvas.bind("<Configure>", _on_canvas_configure)
        self._lista_canvas.bind("<Enter>",
            lambda e: self._lista_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self._lista_canvas.bind("<Leave>",
            lambda e: self._lista_canvas.unbind_all("<MouseWheel>"))

        # Barra de estado (solo modo completo)
        self._lbl_estado = ctk.CTkLabel(
            f, text="", font=ctk.CTkFont(size=11), text_color="gray"
        )
        self._lbl_estado.grid(row=2, column=0, padx=12, pady=(0, 6), sticky="w")

        # Barra de estado modo simple (oculta inicialmente)
        self._lbl_estado_simple = tk.Label(
            f, text="", font=("Segoe UI", 9),
            fg="#4ade80", bg="#1c1c2e", anchor="w"
        )
        self._lbl_estado_simple.grid(row=3, column=0, padx=6, pady=(0, 3), sticky="ew")
        self._lbl_estado_simple.grid_remove()   # oculta hasta modo simple

        return f

    def _es_modo_limitado(self) -> bool:
        """True si la licencia es limitada (trial expirado sin licencia)."""
        lic = getattr(self.gestor, "licencia", None)
        return lic is not None and lic.modo == "limitado"

    def _ruta_raiz_forzada(self) -> str:
        """En modo limitado devuelve siempre la ruta junto al exe/script."""
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, "Plantillas")

    def _navegar(self, ruta):
        """Entra en una subcarpeta — bloqueado en modo limitado."""
        if self._es_modo_limitado():
            return   # sin navegación en versión limitada
        self._ruta_actual = ruta
        self._var_busqueda.set("")
        self._cargar_lista_plantillas()

    def _limites_licencia(self):
        """Devuelve (max_docx, max_odt) según la licencia activa. None = sin límite."""
        lic = getattr(self.gestor, "licencia", None)
        if lic is None:
            return None, None
        return getattr(lic, "max_docx", None), getattr(lic, "max_odt", None)

    def _aplicar_limite_plantillas(self, archivos: list) -> list:
        """
        Filtra la lista respetando límites DOCX/ODT de la licencia.
        El límite es GLOBAL — no se puede esquivar navegando subcarpetas.
        En modo limitado siempre se cuenta desde cero (sin carpetas).
        """
        max_docx, max_odt = self._limites_licencia()
        if max_docx is None and max_odt is None:
            return archivos   # sin límite

        result = []
        n_docx = 0
        n_odt  = 0
        for nombre in archivos:
            ext = os.path.splitext(nombre)[1].lower()
            if ext == ".docx":
                if max_docx is None or n_docx < max_docx:
                    result.append(nombre)
                    n_docx += 1
            elif ext == ".odt":
                if max_odt is None or n_odt < max_odt:
                    result.append(nombre)
                    n_odt += 1
            else:
                result.append(nombre)
        return result

    def _aplicar_limite_busqueda(self, items: list) -> list:
        """
        Límite global en búsqueda recursiva.
        Cuenta docx y odt por separado sobre TODOS los resultados encontrados.
        """
        max_docx, max_odt = self._limites_licencia()
        if max_docx is None and max_odt is None:
            return items
        result = []
        nd, no = 0, 0
        for nombre, ruta in items:
            ext = os.path.splitext(ruta)[1].lower()
            if ext == ".docx" and (max_docx is None or nd < max_docx):
                result.append((nombre, ruta)); nd += 1
            elif ext == ".odt" and (max_odt is None or no < max_odt):
                result.append((nombre, ruta)); no += 1
        return result

    def _cargar_lista_plantillas(self):
        """Carga el contenido del directorio actual (modo explorador)."""
        for widget in self._lista_frame.winfo_children():
            widget.destroy()
        self._lista_frame.grid_columnconfigure(0, weight=1)

        # En modo limitado: forzar siempre la ruta del proyecto, ignorar preferencias
        if self._es_modo_limitado():
            ruta_forzada = self._ruta_raiz_forzada()
            self._ruta_actual = ruta_forzada
            ruta_raiz  = ruta_forzada
            ruta_actual = ruta_forzada
        else:
            if not self._ruta_actual:
                self._ruta_actual = self.gestor.preferencias.get("ruta_plantillas", "")
            ruta_raiz  = self.gestor.preferencias.get("ruta_plantillas", "")
            ruta_actual = self._ruta_actual
        filtro = self._var_busqueda.get() if hasattr(self, "_var_busqueda") else ""
        simple = self._modo_simple

        # Colores dinámicos según tema activo
        BG  = "#1c1c2e"
        FG  = "#cdd6f4"
        FG2 = "#e2e8f0"
        HOV = "#2a2a3e"
        SEP = "#313244"
        ACC = "#60a5fa"
        SUB = "#6c7086"

        # Actualizar fondo del canvas al tema actual
        try:
            self._lista_canvas.configure(bg=BG)
            self._lista_frame.configure(bg=BG)
        except Exception:
            pass

        row = 0

        # ── Con filtro: búsqueda recursiva en toda la raíz ──────────────
        if filtro:
            # En modo limitado la búsqueda también se restringe a la raíz
            if self._es_modo_limitado():
                ruta_busq = self._ruta_raiz_forzada()
                items_raw = self.gestor.buscar_plantillas_en(filtro, ruta_busq)
            else:
                items_raw = self.gestor.buscar_plantillas(filtro)
            items = self._aplicar_limite_busqueda(items_raw)
            self._lista_items = items
            if not items:
                tk.Label(self._lista_frame, text="Sin resultados",
                         font=("Segoe UI", 11), fg="#6c7086", bg=BG
                         ).grid(row=0, column=0, pady=20)
                if not simple:
                    self._lbl_estado.configure(text="0 resultados")
                self._lista_canvas.yview_moveto(0)
                return
            if simple:
                for i, (nombre, ruta) in enumerate(items):
                    self._crear_item_simple(i, nombre, ruta)
            else:
                for nombre, ruta in items:
                    self._crear_item_plantilla(row, nombre, ruta)
                    row += 1
                if not simple:
                    n = len(items)
                    self._lbl_estado.configure(
                        text=f"{n} resultado{'s' if n!=1 else ''}")
            self._lista_canvas.yview_moveto(0)
            return

        # ── Sin filtro: explorador de directorios ────────────────────────
        try:
            contenido   = os.listdir(ruta_actual)
        except Exception:
            contenido   = []

        # En modo limitado no se muestran subcarpetas
        if self._es_modo_limitado():
            carpetas = []
        else:
            carpetas = sorted([c for c in contenido
                               if os.path.isdir(os.path.join(ruta_actual, c))
                               and not c.startswith(".")])
        archivos_raw = sorted([f for f in contenido
                               if os.path.splitext(f)[1].lower() in (".docx", ".odt")])
        archivos = self._aplicar_limite_plantillas(archivos_raw)

        # Botón "atrás" si no estamos en la raíz
        es_raiz = os.path.normpath(ruta_actual) == os.path.normpath(ruta_raiz)
        if not es_raiz:
            btn_atras = tk.Frame(self._lista_frame, bg=BG, cursor="hand2")
            btn_atras.grid(row=row, column=0, sticky="ew", padx=4, pady=(2, 4))
            btn_atras.grid_columnconfigure(1, weight=1)

            tk.Label(btn_atras, text="←", bg=BG, fg=ACC,
                     font=("Segoe UI", 12, "bold"), width=2
                     ).grid(row=0, column=0, padx=(6, 2), pady=3)
            lbl_atras = tk.Label(btn_atras, text="Volver atrás", bg=BG, fg=ACC,
                                  font=("Segoe UI", 10), anchor="w", cursor="hand2")
            lbl_atras.grid(row=0, column=1, sticky="ew", padx=2, pady=3)

            def _atras():
                padre = os.path.dirname(ruta_actual)
                # No subir más allá de la raíz
                if os.path.normpath(padre) >= os.path.normpath(ruta_raiz):
                    self._navegar(padre)
            for w in (btn_atras, lbl_atras, btn_atras.winfo_children()[0]):
                w.bind("<Button-1>", lambda e: _atras())
                w.bind("<Enter>",    lambda e, f=btn_atras, l=lbl_atras:
                           [f.configure(bg=HOV), l.configure(bg=HOV)])
                w.bind("<Leave>",    lambda e, f=btn_atras, l=lbl_atras:
                           [f.configure(bg=BG), l.configure(bg=BG)])
            row += 1

        # Separador visual si hay botón atrás
        if not es_raiz:
            tk.Frame(self._lista_frame, bg=SEP, height=1
                     ).grid(row=row, column=0, sticky="ew", padx=8, pady=(0, 4))
            row += 1

        # Carpetas
        for nombre_carp in carpetas:
            ruta_carp = os.path.join(ruta_actual, nombre_carp)
            f_item = tk.Frame(self._lista_frame, bg=BG, cursor="hand2")
            f_item.grid(row=row, column=0, sticky="ew", padx=4, pady=1)
            f_item.grid_columnconfigure(1, weight=1)

            tk.Label(f_item, text="📁", bg=BG, fg="#fbbf24",
                     font=("Segoe UI Emoji", 11), width=2
                     ).grid(row=0, column=0, padx=(6, 2), pady=4)
            lbl_c = tk.Label(f_item, text=nombre_carp, bg=BG, fg=FG2,
                              font=("Segoe UI", 10), anchor="w", cursor="hand2")
            lbl_c.grid(row=0, column=1, sticky="ew", padx=2, pady=4)
            tk.Label(f_item, text="›", bg=BG, fg=SUB,
                     font=("Segoe UI", 12)).grid(row=0, column=2, padx=(0, 8))

            def _abrir(rc=ruta_carp):
                self._navegar(rc)
            def _hov_in(e, w=f_item, l=lbl_c):
                w.configure(bg=HOV); l.configure(bg=HOV)
                for c in w.winfo_children(): c.configure(bg=HOV)
            def _hov_out(e, w=f_item, l=lbl_c):
                w.configure(bg=BG); l.configure(bg=BG)
                for c in w.winfo_children(): c.configure(bg=BG)

            for w in (f_item, lbl_c):
                w.bind("<Button-1>", lambda e, fn=_abrir: fn())
                w.bind("<Enter>",    _hov_in)
                w.bind("<Leave>",    _hov_out)
            for child in f_item.winfo_children():
                child.bind("<Button-1>", lambda e, fn=_abrir: fn())
                child.bind("<Enter>",    _hov_in)
                child.bind("<Leave>",    _hov_out)
            row += 1

        # Separador entre carpetas y archivos
        if carpetas and archivos:
            tk.Frame(self._lista_frame, bg=SEP, height=1
                     ).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
            row += 1

        # Archivos
        for nombre_arch in archivos:
            ruta_arch = os.path.join(ruta_actual, nombre_arch)
            nombre    = os.path.splitext(nombre_arch)[0]
            if simple:
                self._crear_item_simple(row, nombre, ruta_arch)
            else:
                self._crear_item_plantilla(row, nombre, ruta_arch)
            row += 1

        # Barra de estado
        if not simple:
            total = len(archivos)
            self._lbl_estado.configure(
                text=f"{total} plantilla{'s' if total!=1 else ''}"
                     + (f"  ·  {len(carpetas)} carpeta{'s' if len(carpetas)!=1 else ''}"
                        if carpetas else ""))

        self._lista_canvas.yview_moveto(0)

    def _crear_item_simple(self, row, nombre, ruta):
        """Item minimalista para modo simple: toda la fila es clickable."""
        ext   = os.path.splitext(ruta)[1].lower()
        icono = "📄" if ext == ".docx" else "📘"

        # Un solo Label que ocupa toda la fila
        lbl = tk.Label(
            self._lista_frame,
            text=f"  {icono}  {nombre}",
            anchor="w",
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=6, pady=5,
        )
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        self._lista_frame.grid_columnconfigure(0, weight=1)

        # Colores hover según tema
        def _enter(e):
            lbl.configure(bg="#3b82f6", fg="#ffffff")
        def _leave(e):
            _reset_color()
        def _reset_color():
            lbl.configure(bg="#1c1c2e", fg="#cdd6f4")

        _reset_color()
        lbl.bind("<Enter>",           _enter)
        lbl.bind("<Leave>",           _leave)
        lbl.bind("<Button-1>",        lambda e, r=ruta, n=nombre: self._copiar_plantilla(r, n))

    def _crear_item_plantilla(self, row, nombre, ruta):
        """Item tk puro — sin parpadeo, fluido."""
        BG      = "#1c1c2e"
        BG_HOV  = "#1e3a5f"
        FG      = "#cdd6f4"
        FG_BTN  = "#60a5fa"
        ext     = os.path.splitext(ruta)[1].lower()
        icono   = "📄" if ext == ".docx" else "📘"

        item = tk.Frame(self._lista_frame, bg=BG, cursor="hand2")
        item.grid(row=row, column=0, sticky="ew", padx=4, pady=1)
        item.grid_columnconfigure(1, weight=1)

        tk.Label(item, text=icono, bg=BG, fg=FG,
                 font=("Segoe UI Emoji", 11), width=2
                 ).grid(row=0, column=0, padx=(6, 2), pady=4)

        lbl = tk.Label(item, text=nombre, bg=BG, fg=FG,
                       font=("Segoe UI", 10), anchor="w", cursor="hand2")
        lbl.grid(row=0, column=1, sticky="ew", padx=(2, 4), pady=4)

        btn = tk.Label(item, text="Copiar", bg=BG, fg=FG_BTN,
                       font=("Segoe UI", 9), cursor="hand2", padx=6)
        btn.grid(row=0, column=2, padx=(0, 6), pady=4)

        def _enter(e):
            for w in (item, lbl, btn):
                w.configure(bg=BG_HOV)
            item.nametowidget(item.winfo_parent()).configure(bg=BG_HOV) if False else None
        def _leave(e):
            for w in (item, lbl, btn):
                w.configure(bg=BG)
        def _copiar(e=None):
            self._copiar_plantilla(ruta, nombre)

        for w in (item, lbl, btn):
            w.bind("<Enter>",    _enter)
            w.bind("<Leave>",    _leave)
            w.bind("<Button-1>", _copiar)
        # También el label del icono
        item.winfo_children()[0].bind("<Enter>",    _enter)
        item.winfo_children()[0].bind("<Leave>",    _leave)
        item.winfo_children()[0].bind("<Button-1>", _copiar)

    def _on_busqueda_cambio(self, *_):
        self._cargar_lista_plantillas()

    # ── Frame: Ajustes (con pestañas completas) ────────────
    def _crear_frame_ajustes(self):
        import tkinter.ttk as ttk
        import datetime, zipfile, shutil

        outer = ctk.CTkFrame(self._frame_main, corner_radius=12)
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            outer, text="⚙️  Preferencias",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=(16, 8), padx=16, sticky="w")

        # ── Notebook ttk embebido ──────────────────────────
        style = ttk.Style()
        self._ttk_style = style
        style.theme_use("default")
        style.configure("TNotebook",        background="#1c1c2e", borderwidth=0)
        style.configure("TNotebook.Tab",    background="#2a2a3e", foreground="#cdd6f4",
                         padding=[10, 5],   font=("Segoe UI", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", "#3b82f6")],
                  foreground=[("selected", "white")])
        style.configure("TFrame", background="#1c1c2e")
        style.configure("TLabel", background="#1c1c2e", foreground="#cdd6f4",
                         font=("Segoe UI", 10))
        style.configure("TCheckbutton", background="#1c1c2e", foreground="#cdd6f4",
                         font=("Segoe UI", 10))
        style.configure("TLabelframe",  background="#1c1c2e", foreground="#60a5fa")
        style.configure("TLabelframe.Label", background="#1c1c2e", foreground="#60a5fa",
                         font=("Segoe UI", 10, "bold"))
        style.configure("TSeparator",   background="#313244")
        style.configure("TSpinbox",     fieldbackground="#2a2a3e", foreground="#cdd6f4",
                         background="#2a2a3e", font=("Segoe UI", 10))
        style.configure("TCombobox",    fieldbackground="#2a2a3e", foreground="#cdd6f4",
                         background="#2a2a3e", font=("Segoe UI", 10))
        style.configure("TEntry",       fieldbackground="#2a2a3e", foreground="#cdd6f4",
                         background="#2a2a3e", font=("Segoe UI", 10))

        nb = ttk.Notebook(outer)
        nb.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 4))

        BG  = "#1c1c2e"
        FGC = "#cdd6f4"
        WBG = "#2a2a3e"
        SEP = "#313244"

        def lbl(parent, text, bold=False, color=None, **kw):
            f = ("Segoe UI", 10, "bold") if bold else ("Segoe UI", 10)
            return tk.Label(parent, text=text, bg=BG, fg=color or FGC, font=f, **kw)

        def entry(parent, var, width=40):
            return tk.Entry(parent, textvariable=var, width=width,
                            bg=WBG, fg=FGC, insertbackground=FGC,
                            relief="flat", font=("Segoe UI", 10))

        def check(parent, text, var):
            return tk.Checkbutton(parent, text=text, variable=var,
                                  bg=BG, fg=FGC, selectcolor=WBG,
                                  activebackground=BG, activeforeground=FGC,
                                  font=("Segoe UI", 10))

        def spinbox(parent, var, from_, to, width=8):
            return tk.Spinbox(parent, from_=from_, to=to, textvariable=var,
                              width=width, bg=WBG, fg=FGC,
                              buttonbackground=SEP, insertbackground=FGC,
                              relief="flat", font=("Segoe UI", 10))

        def sep(parent):
            tk.Frame(parent, bg=SEP, height=1).pack(fill="x", pady=8)

        pref = self.gestor.preferencias

        # ════════ PESTAÑA 1: DETECTOR ════════
        t1 = tk.Frame(nb, bg=BG)
        nb.add(t1, text="🔍 Detector")

        lbl(t1, "Configuración del detector de comandos", bold=True,
            color="#60a5fa").pack(anchor="w", pady=(12, 8), padx=10)


        sep(t1)

        self._v_detonante       = tk.StringVar(value=pref.get("detonante", "/"))
        self._v_detector        = tk.BooleanVar(value=pref.get("detector_activado", True))
        self._v_t_reinicio      = tk.IntVar(value=pref.get("tiempo_reinicio_buffer", 2))
        self._v_max_res         = tk.IntVar(value=pref.get("max_resultados_popup", 100))
        self._v_ancho_min       = tk.IntVar(value=pref.get("ancho_minimo_popup", 15))
        self._v_pegar           = tk.BooleanVar(value=pref.get("pegar_automaticamente", True))
        self._v_iconos          = tk.BooleanVar(value=pref.get("mostrar_iconos", True))

        r = tk.Frame(t1, bg=BG); r.pack(fill="x", padx=10, pady=3)
        lbl(r, "Carácter detonante:").pack(side="left")
        self._entry_detonante_w = entry(r, self._v_detonante, width=5)
        self._entry_detonante_w.pack(side="left", padx=(6, 0))
        lbl(r, "  (ej: /, #, !)", color="#6c7086").pack(side="left")

        sep(t1)

        self._chk_detector = check(t1, "Activar detector de comandos", self._v_detector)
        self._chk_detector.pack(anchor="w", padx=10, pady=2)
        check(t1, "Pegar automáticamente al seleccionar",    self._v_pegar).pack(anchor="w", padx=10, pady=2)
        check(t1, "Mostrar iconos en el popup",              self._v_iconos).pack(anchor="w", padx=10, pady=2)

        sep(t1)

        r2 = tk.Frame(t1, bg=BG); r2.pack(fill="x", padx=10, pady=3)
        lbl(r2, "Tiempo reinicio buffer (seg):").pack(side="left")
        spinbox(r2, self._v_t_reinicio, 1, 10).pack(side="left", padx=6)

        r3 = tk.Frame(t1, bg=BG); r3.pack(fill="x", padx=10, pady=3)
        lbl(r3, "Máximo resultados en popup:").pack(side="left")
        spinbox(r3, self._v_max_res, 5, 200).pack(side="left", padx=6)

       # r4 = tk.Frame(t1, bg=BG); r4.pack(fill="x", padx=10, pady=3)
       # lbl(r4, "Ancho mínimo popup (caracteres):").pack(side="left")
       # spinbox(r4, self._v_ancho_min, 10, 80).pack(side="left", padx=6)

        # ════════ PESTAÑA 2: ATAJOS ════════
        t2 = tk.Frame(nb, bg=BG)
        nb.add(t2, text="⌨  Atajos")

        lbl(t2, "Atajos de teclado", bold=True, color="#60a5fa").pack(anchor="w", pady=(12, 8), padx=10)

        self._v_atajos = tk.BooleanVar(value=pref.get("atajos_activados", True))
        check(t2, "Activar atajos de teclado", self._v_atajos).pack(anchor="w", padx=10, pady=2)

        sep(t2)

        grid_f = tk.Frame(t2, bg=BG)
        grid_f.pack(fill="x", padx=16, pady=4)
        atajos = [
            ("🔍 Buscar",              "Ctrl+F  /  F3"),
            ("🔄 Refrescar lista",      "F5  /  Ctrl+R"),
            ("🧹 Limpiar búsqueda",     "Escape"),
            ("❌ Cerrar app",           "Ctrl+Q"),
            ("🌐 Detector de comandos", f"{pref.get('detonante','/')}palabra"),
            ("⬆⬇  Navegar popup",      "↑ ↓  (bloqueadas en Word)"),
            ("✅ Seleccionar popup",    "Enter  (bloqueado en Word)"),
            ("✖  Cerrar popup",        "Esc   (bloqueado en Word)"),
        ]
        for i, (desc, tecla) in enumerate(atajos):
            lbl(grid_f, desc,  bold=True ).grid(row=i, column=0, sticky="w", pady=2, padx=4)
            lbl(grid_f, tecla, color="#60a5fa").grid(row=i, column=1, sticky="w", pady=2, padx=8)

        info = tk.LabelFrame(t2, text="Nota", bg=BG, fg="#60a5fa",
                              font=("Segoe UI", 9, "bold"))
        info.pack(fill="x", padx=10, pady=10)
        lbl(info, "• ↑ ↓ Enter Esc se bloquean globalmente cuando el popup está visible.\n"
                  "• El cursor en Word/Excel NO se mueve mientras navegas el popup.",
            color="#6c7086", justify="left").pack(anchor="w", padx=6, pady=4)

        # ════════ PESTAÑA 3: RUTAS ════════
        t3 = tk.Frame(nb, bg=BG)
        nb.add(t3, text="📁 Rutas")

        lbl(t3, "Configuración de directorios", bold=True, color="#60a5fa").pack(anchor="w", pady=(12, 8), padx=10)
        lbl(t3, "Carpeta de plantillas:").pack(anchor="w", padx=10)

        self._v_ruta = tk.StringVar(value=pref.get("ruta_plantillas", ""))
        rf = tk.Frame(t3, bg=BG); rf.pack(fill="x", padx=10, pady=4)
        e_ruta = tk.Entry(rf, textvariable=self._v_ruta, state="readonly", width=50,
                          bg=WBG, fg=FGC, readonlybackground=WBG,
                          relief="flat", font=("Segoe UI", 10))
        e_ruta.pack(side="left", fill="x", expand=True, padx=(0, 6))

        def _examinar():
            from tkinter import filedialog
            ruta = filedialog.askdirectory(title="Seleccionar carpeta de plantillas",
                                           initialdir=self._v_ruta.get())
            if ruta:
                self._v_ruta.set(ruta)
        self._btn_examinar = tk.Button(rf, text="Examinar...", command=_examinar,
                  bg="#3b82f6", fg="white", relief="flat",
                  font=("Segoe UI", 10), padx=8)
        self._btn_examinar.pack(side="right")

        import sys as _sys
        RUTA_ORIG = os.path.join(
            os.path.dirname(_sys.executable if getattr(_sys, "frozen", False)
                            else os.path.abspath(__file__)),
            "Plantillas"
        )

        def _restaurar():
            if self.preguntar("Confirmar", f"¿Restaurar ruta por defecto?\n{RUTA_ORIG}"):
                self._v_ruta.set(RUTA_ORIG)
        self._btn_restaurar = tk.Button(t3, text="Restaurar ruta por defecto", command=_restaurar,
                  bg="#374151", fg="white", relief="flat",
                  font=("Segoe UI", 10), padx=8)
        self._btn_restaurar.pack(anchor="w", padx=10, pady=4)

        lbl(t3, f"Ajustes:   {os.path.dirname(pref.get('ruta_plantillas',''))}",
            color="#6c7086").pack(anchor="w", padx=10, pady=(8, 0))

        # ════════ PESTAÑA 4: DEPURACIÓN ════════
        t4 = tk.Frame(nb, bg=BG)
        nb.add(t4, text="🐛 Depuración")

        lbl(t4, "Opciones de depuración y logs", bold=True, color="#60a5fa").pack(anchor="w", pady=(12, 8), padx=10)

        self._v_debug      = tk.BooleanVar(value=pref.get("debug", False))
        self._v_privacidad = tk.BooleanVar(value=pref.get("privacidad_logs", True))

        check(t4, "Activar modo depuración (guardar logs)", self._v_debug).pack(anchor="w", padx=10, pady=2)
        check(t4, "🔒 Modo privacidad (no registrar teclas fuera de comandos)",
              self._v_privacidad).pack(anchor="w", padx=10, pady=2)

        lbl(t4, "  • Activado: solo registra cuando se usa el detector\n"
                "  • Desactivado: registra todas las teclas (solo depuración avanzada)",
            color="#6c7086", justify="left").pack(anchor="w", padx=20)

        warn = tk.LabelFrame(t4, text="⚠️ Información", bg=BG, fg="#60a5fa",
                              font=("Segoe UI", 9, "bold"))
        warn.pack(fill="x", padx=10, pady=8)
        lbl(warn, "Los logs se guardan localmente.\nRevísalos antes de enviarlos al soporte técnico.\n"
                  "Desactiva la depuración cuando no sea necesaria.",
            color="#a78bfa", justify="left").pack(anchor="w", padx=6, pady=4)

        log_dir = os.path.join(os.path.dirname(pref.get("ruta_plantillas", "")),
                               "..", "Ajustes", "logs")
        log_dir = os.path.normpath(log_dir)
        lbl(t4, f"Directorio: {log_dir}", color="#6c7086").pack(anchor="w", padx=10, pady=(4, 0))

        btn_frame = tk.Frame(t4, bg=BG); btn_frame.pack(fill="x", padx=10, pady=6)

        def _abrir_logs():
            if os.path.exists(log_dir):
                os.startfile(log_dir)
            else:
                self.mostrar_mensaje("Sin logs", "Todavía no hay logs generados.")
        def _limpiar_logs():
            if os.path.exists(log_dir):
                try:
                    shutil.rmtree(log_dir); os.makedirs(log_dir)
                    self.mostrar_mensaje("Logs", "Logs eliminados correctamente.")
                except Exception as ex:
                    self.mostrar_mensaje("Error", str(ex), "error")

        for txt, cmd in [("📂 Abrir carpeta", _abrir_logs), ("🗑️ Limpiar logs", _limpiar_logs)]:
            tk.Button(btn_frame, text=txt, command=cmd, bg="#374151", fg="white",
                      relief="flat", font=("Segoe UI", 10), padx=8).pack(side="left", padx=4)

        # ════════ PESTAÑA 5: FORMATO ════════
        t5 = tk.Frame(nb, bg=BG)
        nb.add(t5, text="📝 Formato")

        lbl(t5, "Configuración del formato al pegar", bold=True, color="#60a5fa").pack(anchor="w", pady=(12, 8), padx=10)

        fuente_actual = pref.get("fuente_html", {"modo": "personalizado", "familia": "Arial", "tamaño": 10, "color": "#000000"})
        self._v_fuente_modo    = tk.StringVar(value=fuente_actual.get("modo", "origen"))
        self._v_fuente_familia = tk.StringVar(value=fuente_actual.get("familia", "Arial"))
        self._v_fuente_tamaño  = tk.IntVar(value=fuente_actual.get("tamaño", 10))
        self._v_fuente_color   = tk.StringVar(value=fuente_actual.get("color", "#000000"))

        rm = tk.Frame(t5, bg=BG); rm.pack(fill="x", padx=10, pady=4)
        lbl(rm, "Modo fuente:").pack(side="left")
        for val, txt in [("origen", "Respetar original"), ("personalizado", "Personalizado")]:
            tk.Radiobutton(rm, text=txt, variable=self._v_fuente_modo, value=val,
                           bg=BG, fg=FGC, selectcolor=WBG,
                           activebackground=BG, activeforeground=FGC,
                           font=("Segoe UI", 10)).pack(side="left", padx=6)

        r5a = tk.Frame(t5, bg=BG); r5a.pack(fill="x", padx=10, pady=3)
        lbl(r5a, "Familia de fuente:").pack(side="left")
        fuentes_lista = ["Arial", "Verdana", "Times New Roman", "Courier New",
                         "Georgia", "Calibri", "Segoe UI", "Tahoma"]
        cb = ttk.Combobox(r5a, textvariable=self._v_fuente_familia,
                          values=fuentes_lista, width=18,
                          font=("Segoe UI", 10))
        cb.pack(side="left", padx=6)

        r5b = tk.Frame(t5, bg=BG); r5b.pack(fill="x", padx=10, pady=3)
        lbl(r5b, "Tamaño (pt):").pack(side="left")
        spinbox(r5b, self._v_fuente_tamaño, 8, 72).pack(side="left", padx=6)

        r5c = tk.Frame(t5, bg=BG); r5c.pack(fill="x", padx=10, pady=3)
        lbl(r5c, "Color del texto:").pack(side="left")
        entry(r5c, self._v_fuente_color, width=10).pack(side="left", padx=6)

        # Muestra un cuadradito con el color actual
        self._prev_color_btn = tk.Label(r5c, text="   ", bg=self._v_fuente_color.get(),
                                        relief="solid", bd=1, width=3, cursor="hand2")
        self._prev_color_btn.pack(side="left", padx=(0, 6))

        def _elegir_color():
            from tkinter import colorchooser
            color = colorchooser.askcolor(
                color=self._v_fuente_color.get(),
                title="Seleccionar color del texto"
            )
            if color and color[1]:          # color[1] es el hex "#rrggbb"
                self._v_fuente_color.set(color[1])
                self._prev_color_btn.configure(bg=color[1])

        tk.Button(r5c, text="🎨 Elegir", command=_elegir_color,
                  bg="#374151", fg="white", relief="flat",
                  font=("Segoe UI", 10), padx=8, cursor="hand2").pack(side="left")

        # Actualizar el cuadradito cuando se escribe el hex a mano
        def _sync_color_btn(*_):
            try:
                color = self._v_fuente_color.get()
                self._prev_color_btn.configure(bg=color)
            except Exception:
                pass
        self._v_fuente_color.trace_add("write", _sync_color_btn)

        sep(t5)

        prev_lf = tk.LabelFrame(t5, text="Vista previa", bg=BG, fg="#60a5fa",
                                 font=("Segoe UI", 9, "bold"))
        prev_lf.pack(fill="x", padx=10, pady=4)
        self._prev_texto = tk.Text(prev_lf, height=3, width=40,
                                   bg="#ffffff", fg="#000000",
                                   relief="solid", bd=1,
                                   font=("Arial", 10),
                                   padx=8, pady=6)   # siempre blanco — simula Word
        self._prev_texto.pack(padx=6, pady=6, fill="x")
        self._prev_texto.insert("1.0", "Texto de ejemplo con el formato seleccionado.\nAsí se verá al pegar en Word.")

        def _actualizar_preview(*_):
            try:
                self._prev_texto.configure(
                    font=(self._v_fuente_familia.get(), self._v_fuente_tamaño.get()),
                    fg=self._v_fuente_color.get(),
                    bg="#ffffff"   # siempre blanco — simula un doc de Word
                )
            except Exception:
                pass
        self._v_fuente_familia.trace_add("write", _actualizar_preview)
        self._v_fuente_tamaño.trace_add("write",  _actualizar_preview)
        self._v_fuente_color.trace_add("write",   _actualizar_preview)

        # ════════ BOTONES GUARDAR / CANCELAR ════════
        btn_row = tk.Frame(outer, bg="#1c1c2e")
        btn_row.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 12))

        tk.Button(btn_row, text="💾  Guardar", command=self._guardar_ajustes,
                  bg="#3b82f6", fg="white", font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=16, pady=6).pack(side="right", padx=4)
        tk.Button(btn_row, text="↺  Recargar",
                  command=self._actualizar_campo_ruta,
                  bg="#374151", fg="white", font=("Segoe UI", 11),
                  relief="flat", padx=12, pady=6).pack(side="right", padx=4)

        # Aplicar restricciones visuales si el modo es limitado
        self.app.after(100, self._aplicar_restricciones_ajustes)

        return outer

    def _actualizar_banner_licencia(self):
        """Actualiza el texto del banner de licencia en el sidebar."""
        if not hasattr(self, "_lbl_licencia"):
            return
        lic = getattr(self.gestor, "licencia", None)
        if lic is None:
            return
        if lic.modo == "completo":
            self._lbl_licencia.configure(text="✅ Licencia activa", text_color="#4ade80")
        elif lic.modo == "offline":
            self._lbl_licencia.configure(text="📶 Licencia offline", text_color="#fbbf24")
        elif lic.modo == "trial":
            dias = getattr(lic, "dias_trial", 0)
            self._lbl_licencia.configure(
                text=f"🕐 Prueba — {dias}d restantes",
                text_color="#60a5fa"
            )
        else:
            self._lbl_licencia.configure(text="⚠️ Versión limitada", text_color="#f87171")

    def _aplicar_restricciones_ajustes(self):
        """Deshabilita los controles que no se pueden usar en modo limitado."""
        if not self._es_modo_limitado():
            return

        DISABLED_BG  = "#111120"
        DISABLED_FG  = "#44445a"

        def _deshabilitar_widget(w):
            try:
                w.configure(state="disabled")
            except Exception:
                pass
            try:
                w.configure(bg=DISABLED_BG, fg=DISABLED_FG,
                            disabledbackground=DISABLED_BG,
                            disabledforeground=DISABLED_FG)
            except Exception:
                pass

        # ── Pestaña Detector: detonante y checkbox detector ───────────────
        if hasattr(self, "_entry_detonante_w"):
            _deshabilitar_widget(self._entry_detonante_w)
        if hasattr(self, "_chk_detector"):
            _deshabilitar_widget(self._chk_detector)

        # ── Pestaña Rutas: examinar y restaurar ───────────────────────────
        if hasattr(self, "_btn_examinar"):
            _deshabilitar_widget(self._btn_examinar)
        if hasattr(self, "_btn_restaurar"):
            _deshabilitar_widget(self._btn_restaurar)

        # Mostrar banner informativo en la pestaña de rutas
        if hasattr(self, "_btn_restaurar"):
            parent = self._btn_restaurar.master
            banner = tk.Frame(parent, bg="#3b0f0f")
            banner.pack(fill="x", padx=10, pady=(8, 0))
            tk.Label(
                banner,
                text="🔒  En versión limitada la ruta está fijada al directorio del programa."
                     "\n     Activa una licencia para configurar rutas personalizadas.",
                bg="#3b0f0f", fg="#fca5a5",
                font=("Segoe UI", 9), justify="left", pady=6, padx=8
            ).pack(anchor="w")

        # Banner en pestaña detector
        if hasattr(self, "_chk_detector"):
            parent = self._chk_detector.master
            banner2 = tk.Frame(parent, bg="#3b0f0f")
            banner2.pack(fill="x", padx=10, pady=(6, 0))
            tk.Label(
                banner2,
                text="🔒  El detector global no está disponible en la versión limitada."
                     "\n     Activa una licencia para usar el detector en cualquier app.",
                bg="#3b0f0f", fg="#fca5a5",
                font=("Segoe UI", 9), justify="left", pady=6, padx=8
            ).pack(anchor="w")

    def _actualizar_estado_detector(self):
        activo = self.gestor.preferencias.get("detector_activado", True)
        self._lbl_detector.configure(
            text="● Detector activo" if activo else "○ Detector inactivo",
            text_color="#4ade80" if activo else "#f87171"
        )

    def _actualizar_campo_ruta(self):
        """Recarga todos los campos de ajustes con los valores actuales."""
        if not hasattr(self, "_v_ruta"):
            return
        pref = self.gestor.preferencias
        self._v_ruta.set(pref.get("ruta_plantillas", ""))
        self._v_detonante.set(pref.get("detonante", "/"))
        self._v_detector.set(pref.get("detector_activado", True))
        self._v_pegar.set(pref.get("pegar_automaticamente", True))
        self._v_iconos.set(pref.get("mostrar_iconos", True))
        self._v_debug.set(pref.get("debug", False))
        self._v_privacidad.set(pref.get("privacidad_logs", True))
        self._v_t_reinicio.set(pref.get("tiempo_reinicio_buffer", 2))
        self._v_max_res.set(pref.get("max_resultados_popup", 100))
        self._v_ancho_min.set(pref.get("ancho_minimo_popup", 15))
        self._v_atajos.set(pref.get("atajos_activados", True))
        fuente = pref.get("fuente_html", {})
        self._v_fuente_modo.set(fuente.get("modo", "personalizado"))
        self._v_fuente_familia.set(fuente.get("familia", "Arial"))
        self._v_fuente_tamaño.set(fuente.get("tamaño", 10))
        self._v_fuente_color.set(fuente.get("color", "#000000"))

    def _guardar_ajustes(self):
        if not hasattr(self, "_v_ruta"):
            return

        # En modo limitado la ruta no es configurable
        if self._es_modo_limitado():
            nueva_ruta = self._ruta_raiz_forzada()
        else:
            nueva_ruta = self._v_ruta.get().strip()
            if not os.path.exists(nueva_ruta):
                self.mostrar_mensaje("Ruta no válida",
                                     f"La carpeta '{nueva_ruta}' no existe.", "warning")
                return
        nuevas = self.gestor.preferencias.copy()
        nuevas["ruta_plantillas"]        = nueva_ruta
        nuevas["detonante"]              = self._v_detonante.get().strip() or "/"
        nuevas["detector_activado"]      = self._v_detector.get()
        nuevas["pegar_automaticamente"]  = self._v_pegar.get()
        nuevas["mostrar_iconos"]         = self._v_iconos.get()
        nuevas["debug"]                  = self._v_debug.get()
        nuevas["privacidad_logs"]        = self._v_privacidad.get()
        nuevas["tiempo_reinicio_buffer"] = self._v_t_reinicio.get()
        nuevas["max_resultados_popup"]   = self._v_max_res.get()
        nuevas["ancho_minimo_popup"]     = self._v_ancho_min.get()
        nuevas["atajos_activados"]       = self._v_atajos.get()
        nuevas["fuente_html"] = {
            "modo":    self._v_fuente_modo.get(),
            "familia": self._v_fuente_familia.get(),
            "tamaño":  self._v_fuente_tamaño.get(),
            "color":   self._v_fuente_color.get(),
        }
        ok = self.gestor.guardar_preferencias(nuevas)
        if ok:
            self._actualizar_estado_detector()
            self._cargar_lista_plantillas()
            self.mostrar_mensaje("Ajustes", "Preferencias guardadas correctamente ✅")
        else:
            self.mostrar_mensaje("Error", "No se pudieron guardar las preferencias.", "error")

    # ── Frame: Acerca de ──────────────────────────────────
    def _crear_frame_acerca(self):
        f = ctk.CTkFrame(self._frame_main, corner_radius=12)
        f.grid_rowconfigure(0, weight=1)
        f.grid_columnconfigure(0, weight=1)

        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(inner, text="📋", font=ctk.CTkFont(size=52)).pack()
        ctk.CTkLabel(
            inner, text=self.TITULO,
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(pady=(4, 0))
        ctk.CTkLabel(
            inner, text=f"Versión {self.VERSION}",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        ).pack()

        ctk.CTkFrame(inner, height=1, fg_color="gray").pack(
            fill="x", pady=16
        )

        info = [
            ("Motor",       "Python 3.x + customtkinter"),
            ("Formatos",    "DOCX · ODT"),
            ("Popup",       "Win32 nativo (sin robo de foco)"),
            ("Detector",    "keyboard hooks global"),
        ]
        for etiq, val in info:
            row = ctk.CTkFrame(inner, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=etiq+":", width=90, anchor="e",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="#60a5fa").pack(side="left", padx=(0, 8))
            ctk.CTkLabel(row, text=val, anchor="w",
                         font=ctk.CTkFont(size=12)).pack(side="left")

        ctk.CTkFrame(inner, height=1, fg_color="gray").pack(fill="x", pady=10)

        # ── Sección de licencia ───────────────────────────────────────────
        lic = getattr(self.gestor, "licencia", None)
        try:
            from licencia import obtener_info_cache, URL_TIENDA, desactivar_licencia
            cache = obtener_info_cache()
            _lic_mod = True
        except ImportError:
            cache     = None
            _lic_mod  = False

        if cache and lic and lic.es_completo:
            # Licencia activa
            activado = cache.get("activado_en", "")[:10]
            ctk.CTkLabel(inner,
                         text=f"✅  Licencia activa desde {activado}  ···{cache.get('clave','')[-4:]}",
                         font=ctk.CTkFont(size=11), text_color="#4ade80").pack(pady=(0,4))

            def _desactivar():
                if self.preguntar("Desactivar licencia",
                    "¿Desactivar la licencia en este dispositivo?\n"
                    "La app se cerrará y necesitarás reactivarla."):
                    desactivar_licencia()
                    self._salir()

            ctk.CTkButton(inner, text="Desactivar en este dispositivo",
                width=280, height=28, corner_radius=8,
                font=ctk.CTkFont(size=11),
                fg_color="transparent", border_width=1,
                border_color="#f87171", text_color="#f87171",
                hover_color="#3b0f0f",
                command=_desactivar).pack(pady=(0, 4))

        elif lic and lic.modo == "trial":
            dias = getattr(lic, "dias_trial", 0)
            ctk.CTkLabel(inner,
                         text=f"🕐  Período de prueba — {dias} día{'s' if dias!=1 else ''} restante{'s' if dias!=1 else ''}",
                         font=ctk.CTkFont(size=11), text_color="#60a5fa").pack(pady=(0, 4))
            ctk.CTkButton(inner, text="Activar licencia completa",
                width=240, height=30, corner_radius=8,
                font=ctk.CTkFont(size=11),
                command=lambda: self._abrir_activacion()).pack(pady=(0, 4))

        else:
            ctk.CTkLabel(inner, text="⚠️  Versión limitada (prueba expirada)",
                         font=ctk.CTkFont(size=11), text_color="#f87171").pack(pady=(0, 4))
            ctk.CTkButton(inner, text="Activar licencia",
                width=200, height=30, corner_radius=8,
                font=ctk.CTkFont(size=11),
                command=lambda: self._abrir_activacion()).pack(pady=(0, 4))

        ctk.CTkLabel(
            inner,
            text="\nTip: Escribe '/' seguido de texto para buscar plantillas\nen cualquier aplicación de Windows.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            justify="center"
        ).pack(pady=(8, 0))

        return f

    def _abrir_activacion(self):
        """Abre la ventana de activación de licencia."""
        try:
            from ventana_licencia import pedir_licencia
            from licencia import comprobar_licencia
            pedir_licencia(self.gestor.licencia, parent=self.app)
            # Re-comprobar y actualizar
            nueva_lic = comprobar_licencia()
            self.gestor.licencia = nueva_lic
            self._actualizar_banner_licencia()
            # Recargar lista por si cambió el límite
            self._cargar_lista_plantillas()
            # Si ahora tiene licencia completa, recargar "Acerca de"
            if nueva_lic.es_completo:
                self._frame_acerca.destroy()
                self._frame_acerca = self._crear_frame_acerca()
                self._vista_acerca()
        except Exception as e:
            self.mostrar_mensaje("Error", str(e), "error")

    # ── Lógica plantillas ──────────────────────────────────
    def _copiar_plantilla(self, ruta, nombre):
        ok = self.gestor.procesar_plantilla(ruta)
        if ok:
            if self._modo_simple:
                # Modo simple: etiqueta compacta debajo de la lista
                self._lbl_estado_simple.configure(text=f"✅  {nombre}")
                self.app.after(3000, lambda: self._lbl_estado_simple.configure(text=""))
            else:
                self._lbl_estado.configure(
                    text=f"✅ '{nombre}' copiada"
                )
                self.app.after(4000, lambda: self._lbl_estado.configure(text=""))
        else:
            self.mostrar_mensaje("Error", f"No se pudo procesar '{nombre}'.", "error")

    def _on_plantilla_seleccionada(self, ruta):
        """Callback cuando el popup selecciona una plantilla."""
        buffer   = self.gestor.buffer_teclas[:]
        n_borrar = len(buffer)
        self.gestor.limpiar_buffer()

        def _ejecutar():
            import keyboard as kb
            time.sleep(0.15)
            if n_borrar > 0:
                for _ in range(n_borrar):
                    kb.send("backspace")
                time.sleep(0.05)
            ok = self.gestor.procesar_plantilla(ruta)
            if ok and self.gestor.preferencias.get("pegar_automaticamente", True):
                time.sleep(0.08)
                kb.send("ctrl+v")
            # No tocar el estado de la ventana — el popup ya tiene
            # WS_EX_NOACTIVATE, así que el foco nunca salió de Word/Excel.

        threading.Thread(target=_ejecutar, daemon=True).start()

    # ── Cierre ─────────────────────────────────────────────
    def _salir(self):
        # Guardar si estaba en modo simple para recordarlo la próxima vez
        if self.popup:
            self.popup.cerrar()
        self.app.destroy()