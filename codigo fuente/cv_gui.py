import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from cv_generator import TEMPLATES, LABELS_ES, LABELS_EN, generate_pdf, generate_word, parse_text

OUTPUT_DIR = os.path.join(SCRIPT_DIR, "CVs generados")
DATA_FILE = os.path.join(SCRIPT_DIR, "datos.txt")

TEMPLATE_INFO = {
    "clasico":      {"desc": "Barras de color, nombre a la izquierda",         "color": "#1B3A5C"},
    "ejecutivo":    {"desc": "Nombre centrado, fecha en italica",              "color": "#4A1942"},
    "moderno":      {"desc": "Borde izquierdo de color, fondo gris",           "color": "#0B3D6B"},
    "elegante":     {"desc": "Secciones en MAYUSCULAS, línea fina",            "color": "#2D2D2D"},
    "profesional":  {"desc": "Borde inferior grueso, compacto",                "color": "#1B4332"},
    "creativo":     {"desc": "Puntos de color, minimalista",                   "color": "#C0392B"},
    "corporativo":  {"desc": "Header oscuro, nombre en blanco",                "color": "#1C2833"},
    "fresco":       {"desc": "Etiquetas de color, juvenil",                    "color": "#117A65"},
    "sobrio":       {"desc": "Ultra minimalista, sin adornos",                 "color": "#1A1A2E"},
    "audaz":        {"desc": "Bloque cuadrado, fuerte y directo",              "color": "#B7410E"},
}

BG       = "#F0F2F5"
CARD     = "#FFFFFF"
PRIMARY  = "#2563EB"
TEXT     = "#1E293B"
HINT     = "#94A3B8"
BORDER   = "#E2E8F0"
GREEN    = "#16A34A"
RED      = "#DC2626"
INPUT_BG = "#F8FAFC"


# ── Placeholder mixin for Entry and Text ──

class PlaceholderEntry(tk.Entry):
    def __init__(self, master, placeholder="", **kw):
        super().__init__(master, **kw)
        self.placeholder = placeholder
        self._ph_color = HINT
        self._fg = kw.get("fg", TEXT)
        self._showing_ph = False
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self._show_placeholder()

    def _show_placeholder(self):
        if not self.get():
            self._showing_ph = True
            self.insert(0, self.placeholder)
            self.config(fg=self._ph_color)

    def _on_focus_in(self, _):
        if self._showing_ph:
            self.delete(0, "end")
            self.config(fg=self._fg)
            self._showing_ph = False

    def _on_focus_out(self, _):
        if not self.get():
            self._show_placeholder()

    def get_value(self):
        return "" if self._showing_ph else self.get().strip()

    def set_value(self, val):
        self._showing_ph = False
        self.config(fg=self._fg)
        self.delete(0, "end")
        if val:
            self.insert(0, val)
        else:
            self._show_placeholder()


class PlaceholderText(tk.Text):
    def __init__(self, master, placeholder="", **kw):
        super().__init__(master, **kw)
        self.placeholder = placeholder
        self._ph_color = HINT
        self._fg = kw.get("fg", TEXT)
        self._showing_ph = False
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self._show_placeholder()

    def _show_placeholder(self):
        if not self.get("1.0", "end").strip():
            self._showing_ph = True
            self.insert("1.0", self.placeholder)
            self.config(fg=self._ph_color)

    def _on_focus_in(self, _):
        if self._showing_ph:
            self.delete("1.0", "end")
            self.config(fg=self._fg)
            self._showing_ph = False

    def _on_focus_out(self, _):
        if not self.get("1.0", "end").strip():
            self._show_placeholder()

    def get_value(self):
        return "" if self._showing_ph else self.get("1.0", "end").strip()

    def set_value(self, val):
        self._showing_ph = False
        self.config(fg=self._fg)
        self.delete("1.0", "end")
        if val:
            self.insert("1.0", val)
        else:
            self._show_placeholder()


class CVGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CV Generator")
        self.root.configure(bg=BG)
        self.root.minsize(780, 620)

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w, win_h = 830, 720
        x = (screen_w - win_w) // 2
        y = max(0, (screen_h - win_h) // 2 - 30)
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")

        self.experiences = []
        self.education_list = []
        self.foto_path = None

        self._setup_styles()
        self._build_ui()
        self._load_if_exists()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TCombobox", font=("Segoe UI", 10))

    # ── Scrollable container ──

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=BG)

        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_win = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def resize_canvas(event):
            canvas.itemconfig(canvas_win, width=event.width)
        canvas.bind("<Configure>", resize_canvas)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.canvas = canvas
        self._build_content(self.scroll_frame)

    # ── Main content ──

    def _build_content(self, parent):
        # ── Header con titulo y botones guardar/cargar ──
        hdr = tk.Frame(parent, bg=PRIMARY)
        hdr.pack(fill="x")
        hdr_inner = tk.Frame(hdr, bg=PRIMARY)
        hdr_inner.pack(fill="x", padx=25, pady=(18, 14))

        tk.Label(hdr_inner, text="Generador de CV",
                 font=("Segoe UI", 22, "bold"), fg="white", bg=PRIMARY).pack(side="left")
        tk.Label(hdr_inner, text="Llena tus datos, elige un diseño y genera tu CV en un click",
                 font=("Segoe UI", 10), fg="#BDD4FF", bg=PRIMARY).pack(side="left", padx=(15, 0), pady=(6, 0))

        self._flat_btn(hdr_inner, "Pegar texto del cliente", "#3B82F6",
                       self._paste_text_dialog, 10).pack(side="right")

        content = tk.Frame(parent, bg=BG)
        content.pack(fill="x", padx=22, pady=(12, 0))

        # ── 1. PLANTILLA ──
        self._section_number(content, "1", "Elige el diseño de tu CV")
        card = self._card(content)
        self._build_template_selector(card)

        # ── FOTO ──
        foto_frame = tk.Frame(content, bg=BG)
        foto_frame.pack(fill="x", pady=(10, 0))

        self.foto_var = tk.BooleanVar(value=False)
        foto_check = tk.Checkbutton(foto_frame,
            text="  Incluir foto en el CV  (esquina superior izquierda)",
            variable=self.foto_var,
            font=("Segoe UI", 10), fg=TEXT, bg=BG, activebackground=BG,
            selectcolor=CARD, cursor="hand2",
            command=self._toggle_foto)
        foto_check.pack(anchor="w")

        self.foto_controls = tk.Frame(content, bg=BG)
        self._foto_frame_ref = foto_frame

        foto_inner = tk.Frame(self.foto_controls, bg=CARD, relief="solid", bd=1)
        foto_inner.pack(fill="x", padx=15, pady=(2, 6))

        self._flat_btn(foto_inner, "Seleccionar foto", PRIMARY,
                       self._select_foto, 9).pack(side="left", padx=(10, 8), pady=8)
        self.foto_label = tk.Label(foto_inner, text="Ningún archivo seleccionado",
                                    font=("Segoe UI", 9), fg=HINT, bg=CARD)
        self.foto_label.pack(side="left", padx=(0, 10), pady=8)

        self.foto_clear_btn = tk.Button(foto_inner, text="Quitar",
            font=("Segoe UI", 8), fg="white", bg=RED, activebackground="#B91C1C",
            relief="flat", cursor="hand2", padx=6, pady=2,
            command=self._clear_foto)
        self.foto_clear_btn.pack(side="right", padx=(0, 10), pady=8)
        self.foto_clear_btn.pack_forget()

        # ── 2. DATOS PERSONALES ──
        self._section_number(content, "2", "Tus datos personales")
        card = self._card(content)
        self.nombre_entry = self._entry(card, "Tu nombre completo",
                                         "Ej: Laura Aguilar Martinez")
        self.contacto_entry = self._entry(card, "Teléfono, email y/o LinkedIn (separados con |)",
                                           "Ej: 81 1234 5678 | mi.correo@gmail.com | linkedin.com/in/mi-perfil")
        self.titulo_entry = self._entry(card, "Puesto o area profesional",
                                         "Ej: Ejecutiva Comercial | Ventas B2B | Desarrollo de Negocio")
        tk.Frame(card, height=6, bg=CARD).pack()

        # ── 3. RESUMEN ──
        self._section_number(content, "3", "Resumen profesional  (opcional)")
        card = self._card(content)
        self.resumen_text = self._textarea(card,
            "Escribe un parrafo corto describiendo tu experiencia y que buscas.\n"
            "Ej: Profesional con 5 años de experiencia en ventas B2B,\n"
            "especializada en apertura de cuentas y crecimiento de cartera.", 4)

        # ── 4. LOGROS ──
        self._section_number(content, "4", "Logros clave  (opcional)")
        card = self._card(content)
        self.logros_text = self._textarea(card,
            "Escribe un logro por línea. El guión se agrega solo.\n"
            "Ej:\n"
            "1er lugar en ventas por 3 años consecutivos\n"
            "Incremento del 100% en ventas anuales\n"
            "Apertura de cuenta por +$2 MDP mensuales", 5)

        # ── 5. EXPERIENCIA ──
        self._section_number(content, "5", "Experiencia laboral")
        card = self._card(content)
        self.exp_container = tk.Frame(card, bg=CARD)
        self.exp_container.pack(fill="x", padx=15, pady=(0, 5))
        self._empty_msg(self.exp_container, "exp",
                        "Aún no has agregado experiencia. Dale click al botón de abajo.")
        self._flat_btn(card, "+  Agregar trabajo", PRIMARY, self._add_experience_dialog, 10).pack(
            anchor="w", padx=15, pady=(2, 14))

        # ── 6. EDUCACION ──
        self._section_number(content, "6", "Educación")
        card = self._card(content)
        self.edu_container = tk.Frame(card, bg=CARD)
        self.edu_container.pack(fill="x", padx=15, pady=(0, 5))
        self._empty_msg(self.edu_container, "edu",
                        "Aún no has agregado educación. Dale click al botón de abajo.")
        self._flat_btn(card, "+  Agregar carrera / título", PRIMARY, self._add_education_dialog, 10).pack(
            anchor="w", padx=15, pady=(2, 14))

        # ── 7. CURSOS ──
        self._section_number(content, "7", "Cursos y certificaciones  (opcional)")
        card = self._card(content)
        self.cursos_text = self._textarea(card,
            "Escribe un curso o certificación por línea.\n"
            "Ej:\n"
            "Certificación Google Ads\n"
            "Diplomado en Marketing Digital", 4)

        # ── 8. HABILIDADES ──
        self._section_number(content, "8", "Habilidades  (opcional)")
        card = self._card(content)
        self.habilidades_text = self._textarea(card,
            "Escribe una categoría y sus habilidades separadas por : (dos puntos)\n"
            "Ej:\n"
            "Comercial: Ventas B2B, Prospección, Negociación\n"
            "Herramientas: Excel, Salesforce, SAP", 3)

        # ── 9. IDIOMAS ──
        self._section_number(content, "9", "Idiomas  (opcional)")
        card = self._card(content)
        self.idiomas_text = self._textarea(card,
            "Escribe un idioma por línea.\n"
            "Ej:\n"
            "Espanol - Nativo\n"
            "Ingles - Avanzado", 3)

        # ── OPCION INGLES ──
        opt_frame = tk.Frame(parent, bg=BG)
        opt_frame.pack(fill="x", padx=22, pady=(16, 0))

        self.english_var = tk.BooleanVar(value=False)
        eng_check = tk.Checkbutton(opt_frame,
            text="  También exportar en inglés  (genera versión ES + EN)",
            variable=self.english_var,
            font=("Segoe UI", 10), fg=TEXT, bg=BG, activebackground=BG,
            selectcolor=CARD, cursor="hand2")
        eng_check.pack(anchor="w")

        # ── BOTON GENERAR ──
        gen_frame = tk.Frame(parent, bg=BG)
        gen_frame.pack(fill="x", padx=22, pady=(10, 25))

        gen_btn = tk.Button(gen_frame, text="Generar mi CV",
                            font=("Segoe UI", 16, "bold"), fg="white", bg=GREEN,
                            activebackground="#15803D", relief="flat", cursor="hand2",
                            padx=40, pady=12, command=self._generate)
        gen_btn.pack(fill="x", ipady=4)

        hint = tk.Label(gen_frame, text="Se generan PDF + Word en 1 página. Con inglés activado se generan 4 archivos.",
                        font=("Segoe UI", 9), fg=HINT, bg=BG)
        hint.pack(pady=(6, 0))

    # ── UI helpers ──

    def _section_number(self, parent, num, text):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(14, 4))
        circle = tk.Label(f, text=num, font=("Segoe UI", 10, "bold"), fg="white",
                          bg=PRIMARY, width=3, height=1)
        circle.pack(side="left")
        tk.Label(f, text=text, font=("Segoe UI", 12, "bold"), fg=TEXT, bg=BG).pack(
            side="left", padx=(8, 0))

    def _card(self, parent):
        wrapper = tk.Frame(parent, bg=BORDER)
        wrapper.pack(fill="x", pady=(0, 2))
        card = tk.Frame(wrapper, bg=CARD, padx=1, pady=1)
        card.pack(fill="x", padx=1, pady=1)
        return card

    def _entry(self, parent, label, placeholder):
        f = tk.Frame(parent, bg=CARD)
        f.pack(fill="x", padx=15, pady=(8, 2))
        tk.Label(f, text=label, font=("Segoe UI", 10, "bold"), fg=TEXT, bg=CARD).pack(anchor="w")
        entry = PlaceholderEntry(f, placeholder=placeholder,
                                  font=("Segoe UI", 11), bg=INPUT_BG, fg=TEXT,
                                  relief="solid", bd=1, insertbackground=TEXT)
        entry.pack(fill="x", ipady=5, pady=(2, 0))
        return entry

    def _textarea(self, parent, placeholder, height):
        f = tk.Frame(parent, bg=CARD)
        f.pack(fill="x", padx=15, pady=(8, 12))
        text = PlaceholderText(f, placeholder=placeholder,
                                font=("Segoe UI", 10), bg=INPUT_BG, fg=TEXT,
                                relief="solid", bd=1, height=height,
                                wrap="word", insertbackground=TEXT)
        text.pack(fill="x")
        return text

    def _flat_btn(self, parent, text, color, command, size=10):
        btn = tk.Button(parent, text=text, font=("Segoe UI", size, "bold"),
                        fg="white", bg=color, activebackground=color,
                        relief="flat", cursor="hand2", padx=14, pady=5,
                        command=command)
        return btn

    def _empty_msg(self, parent, tag, text):
        lbl = tk.Label(parent, text=text, font=("Segoe UI", 9), fg=HINT, bg=CARD)
        lbl.pack(anchor="w", pady=(4, 4))
        lbl._tag = tag

    # ── Photo helpers ──

    def _toggle_foto(self):
        if self.foto_var.get():
            self.foto_controls.pack(fill="x", pady=(2, 0), after=self._foto_frame_ref)
        else:
            self.foto_controls.pack_forget()

    def _select_foto(self):
        path = filedialog.askopenfilename(
            title="Selecciona una foto",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif"),
                       ("Todos los archivos", "*.*")])
        if path:
            self.foto_path = path
            self.foto_label.config(text=os.path.basename(path), fg=GREEN)
            self.foto_clear_btn.pack(side="right", padx=(0, 10), pady=8)

    def _clear_foto(self):
        self.foto_path = None
        self.foto_label.config(text="Ningún archivo seleccionado", fg=HINT)
        self.foto_clear_btn.pack_forget()

    # ── Template selector with color preview ──

    def _build_template_selector(self, card):
        f = tk.Frame(card, bg=CARD)
        f.pack(fill="x", padx=15, pady=(8, 14))

        self.template_var = tk.StringVar(value="clasico")
        self._tpl_buttons = {}

        for i, (key, info) in enumerate(TEMPLATE_INFO.items()):
            row = tk.Frame(f, bg=CARD, cursor="hand2")
            row.pack(fill="x", pady=(0, 3))

            swatch = tk.Frame(row, bg=info["color"], width=18, height=18)
            swatch.pack(side="left", padx=(0, 8))
            swatch.pack_propagate(False)

            rb = tk.Radiobutton(row, text=f'{key.capitalize()}  —  {info["desc"]}',
                                variable=self.template_var, value=key,
                                font=("Segoe UI", 10), fg=TEXT, bg=CARD,
                                activebackground=CARD, selectcolor=CARD,
                                anchor="w", indicatoron=True)
            rb.pack(side="left", fill="x")
            self._tpl_buttons[key] = rb

    # ── Experience ──

    def _add_experience_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Agregar Trabajo")
        dlg.configure(bg=CARD)
        dlg.resizable(False, False)
        dlg.grab_set()

        w, h = 520, 440
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(dlg, text="Agregar trabajo", font=("Segoe UI", 14, "bold"),
                 fg=PRIMARY, bg=CARD).pack(anchor="w", padx=20, pady=(18, 12))

        fields = {}
        for label, key, ph in [
            ("Puesto", "puesto", "Ej: Ejecutiva Comercial B2B"),
            ("Empresa", "empresa", "Ej: Mercado Libre"),
            ("Fechas", "fechas", "Ej: Ene 2023 - Actual"),
        ]:
            f = tk.Frame(dlg, bg=CARD)
            f.pack(fill="x", padx=20, pady=(0, 6))
            tk.Label(f, text=label, font=("Segoe UI", 10, "bold"), fg=TEXT, bg=CARD).pack(anchor="w")
            e = PlaceholderEntry(f, placeholder=ph, font=("Segoe UI", 10),
                                  bg=INPUT_BG, fg=TEXT, relief="solid", bd=1)
            e.pack(fill="x", ipady=4, pady=(2, 0))
            fields[key] = e

        f = tk.Frame(dlg, bg=CARD)
        f.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(f, text="Que hiciste en este trabajo? (una cosa por línea)",
                 font=("Segoe UI", 10, "bold"), fg=TEXT, bg=CARD).pack(anchor="w")
        bullets_text = PlaceholderText(f,
            placeholder="Ej:\nProspección y cierre de cuentas nuevas\nGestión de cartera de 200 clientes\nCumplimiento de metas mensuales",
            font=("Segoe UI", 10), bg=INPUT_BG, fg=TEXT,
            relief="solid", bd=1, height=5, wrap="word")
        bullets_text.pack(fill="x", pady=(2, 0))

        def save():
            puesto = fields["puesto"].get_value()
            if not puesto:
                messagebox.showwarning("Falta el puesto",
                    "Escribe al menos el nombre del puesto.", parent=dlg)
                return
            exp = {
                "puesto": puesto,
                "empresa": fields["empresa"].get_value(),
                "fechas": fields["fechas"].get_value(),
                "bullets": [l.strip() for l in bullets_text.get_value().split("\n") if l.strip()],
            }
            self.experiences.append(exp)
            self._refresh_exp_list()
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg=CARD)
        btn_f.pack(fill="x", padx=20, pady=(12, 18))
        self._flat_btn(btn_f, "Agregar", GREEN, save).pack(side="right")
        self._flat_btn(btn_f, "Cancelar", HINT, dlg.destroy).pack(side="right", padx=(0, 8))

    def _refresh_exp_list(self):
        for w in self.exp_container.winfo_children():
            w.destroy()
        if not self.experiences:
            self._empty_msg(self.exp_container, "exp",
                            "Aún no has agregado experiencia. Dale click al botón de abajo.")
            return
        for i, exp in enumerate(self.experiences):
            row = tk.Frame(self.exp_container, bg="#F1F5F9", relief="solid", bd=1)
            row.pack(fill="x", pady=(0, 4))
            left = tk.Frame(row, bg="#F1F5F9")
            left.pack(side="left", fill="x", expand=True, padx=10, pady=6)
            tk.Label(left, text=exp["puesto"], font=("Segoe UI", 10, "bold"),
                     fg=TEXT, bg="#F1F5F9", anchor="w").pack(anchor="w")
            sub = exp["empresa"]
            if exp["fechas"]:
                sub += f'  |  {exp["fechas"]}'
            if sub:
                tk.Label(left, text=sub, font=("Segoe UI", 9),
                         fg=HINT, bg="#F1F5F9", anchor="w").pack(anchor="w")
            if exp["bullets"]:
                tk.Label(left, text=f'{len(exp["bullets"])} responsabilidades',
                         font=("Segoe UI", 8), fg="#64748B", bg="#F1F5F9").pack(anchor="w")
            idx = i
            tk.Button(row, text="Quitar", font=("Segoe UI", 8),
                      fg="white", bg=RED, activebackground="#B91C1C",
                      relief="flat", cursor="hand2", padx=8, pady=2,
                      command=lambda j=idx: self._remove_exp(j)).pack(side="right", padx=8, pady=8)

    def _remove_exp(self, idx):
        self.experiences.pop(idx)
        self._refresh_exp_list()

    # ── Education ──

    def _add_education_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Agregar Educacion")
        dlg.configure(bg=CARD)
        dlg.resizable(False, False)
        dlg.grab_set()

        w, h = 480, 260
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(dlg, text="Agregar educación", font=("Segoe UI", 14, "bold"),
                 fg=PRIMARY, bg=CARD).pack(anchor="w", padx=20, pady=(18, 12))

        fields = {}
        for label, key, ph in [
            ("Carrera o título", "carrera", "Ej: Lic. en Administración de Empresas"),
            ("Escuela y años", "escuela", "Ej: UANL | 2010 - 2014"),
        ]:
            f = tk.Frame(dlg, bg=CARD)
            f.pack(fill="x", padx=20, pady=(0, 6))
            tk.Label(f, text=label, font=("Segoe UI", 10, "bold"), fg=TEXT, bg=CARD).pack(anchor="w")
            e = PlaceholderEntry(f, placeholder=ph, font=("Segoe UI", 10),
                                  bg=INPUT_BG, fg=TEXT, relief="solid", bd=1)
            e.pack(fill="x", ipady=4, pady=(2, 0))
            fields[key] = e

        def save():
            carrera = fields["carrera"].get_value()
            if not carrera:
                messagebox.showwarning("Falta la carrera",
                    "Escribe al menos el nombre de la carrera.", parent=dlg)
                return
            edu = {"titulo": carrera, "institucion": fields["escuela"].get_value()}
            self.education_list.append(edu)
            self._refresh_edu_list()
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg=CARD)
        btn_f.pack(fill="x", padx=20, pady=(12, 18))
        self._flat_btn(btn_f, "Agregar", GREEN, save).pack(side="right")
        self._flat_btn(btn_f, "Cancelar", HINT, dlg.destroy).pack(side="right", padx=(0, 8))

    def _refresh_edu_list(self):
        for w in self.edu_container.winfo_children():
            w.destroy()
        if not self.education_list:
            self._empty_msg(self.edu_container, "edu",
                            "Aún no has agregado educación. Dale click al botón de abajo.")
            return
        for i, edu in enumerate(self.education_list):
            row = tk.Frame(self.edu_container, bg="#F1F5F9", relief="solid", bd=1)
            row.pack(fill="x", pady=(0, 4))
            left = tk.Frame(row, bg="#F1F5F9")
            left.pack(side="left", fill="x", expand=True, padx=10, pady=6)
            tk.Label(left, text=edu["titulo"], font=("Segoe UI", 10, "bold"),
                     fg=TEXT, bg="#F1F5F9", anchor="w").pack(anchor="w")
            if edu["institucion"]:
                tk.Label(left, text=edu["institucion"], font=("Segoe UI", 9),
                         fg=HINT, bg="#F1F5F9", anchor="w").pack(anchor="w")
            idx = i
            tk.Button(row, text="Quitar", font=("Segoe UI", 8),
                      fg="white", bg=RED, activebackground="#B91C1C",
                      relief="flat", cursor="hand2", padx=8, pady=2,
                      command=lambda j=idx: self._remove_edu(j)).pack(side="right", padx=8, pady=8)

    def _remove_edu(self, idx):
        self.education_list.pop(idx)
        self._refresh_edu_list()

    # ── Collect form data ──

    def _collect_data(self):
        nombre = self.nombre_entry.get_value()
        if not nombre:
            messagebox.showwarning("Falta tu nombre",
                "Escribe tu nombre completo en el paso 2 antes de generar.")
            return None

        logros_raw = self.logros_text.get_value()
        logros = [l.lstrip("- *").strip() for l in logros_raw.split("\n") if l.strip()] if logros_raw else []

        cursos_raw = self.cursos_text.get_value()
        cursos = [l.lstrip("- *").strip() for l in cursos_raw.split("\n") if l.strip()] if cursos_raw else []

        hab_raw = self.habilidades_text.get_value()
        habilidades = []
        if hab_raw:
            for line in hab_raw.split("\n"):
                line = line.strip()
                if ":" in line:
                    parts = line.split(":", 1)
                    habilidades.append((parts[0].strip(), parts[1].strip()))

        idiomas_raw = self.idiomas_text.get_value()
        idiomas = [l.lstrip("- *").strip() for l in idiomas_raw.split("\n") if l.strip()] if idiomas_raw else []

        resumen_raw = self.resumen_text.get_value()
        resumen = " ".join(l.strip() for l in resumen_raw.split("\n") if l.strip()) if resumen_raw else ""

        foto = None
        if self.foto_var.get() and self.foto_path and os.path.exists(self.foto_path):
            foto = self.foto_path

        return {
            "theme": self.template_var.get(),
            "nombre": nombre,
            "contacto": self.contacto_entry.get_value(),
            "titulo": self.titulo_entry.get_value(),
            "resumen": resumen,
            "logros": logros,
            "experiencia": list(self.experiences),
            "educacion": list(self.education_list),
            "cursos": cursos,
            "habilidades": habilidades,
            "idiomas": idiomas,
            "foto": foto,
        }

    # ── Generate ──

    def _data_to_text(self, data):
        lines = []
        lines.append("DATOS PERSONALES")
        lines.append(f"NOMBRE: {data['nombre']}")
        lines.append(f"CONTACTO: {data['contacto']}")
        lines.append(f"TITULO: {data['titulo']}")
        lines.append("")
        lines.append("RESUMEN PROFESIONAL")
        lines.append(data["resumen"])
        lines.append("")
        if data["logros"]:
            lines.append("LOGROS CLAVE")
            for l in data["logros"]:
                lines.append(f"- {l}")
            lines.append("")
        lines.append("EXPERIENCIA PROFESIONAL")
        for exp in data["experiencia"]:
            lines.append(f"PUESTO: {exp['puesto']}")
            lines.append(f"EMPRESA: {exp['empresa']}")
            lines.append(f"FECHAS: {exp['fechas']}")
            for b in exp["bullets"]:
                lines.append(f"- {b}")
            lines.append("")
        if data["educacion"]:
            lines.append("EDUCACION")
            for edu in data["educacion"]:
                lines.append(f"CARRERA: {edu['titulo']}")
                lines.append(f"ESCUELA: {edu['institucion']}")
            lines.append("")
        if data["cursos"]:
            lines.append("CURSOS Y CERTIFICACIONES")
            for c in data["cursos"]:
                lines.append(f"- {c}")
            lines.append("")
        if data["habilidades"]:
            lines.append("HABILIDADES")
            for cat, det in data["habilidades"]:
                lines.append(f"{cat}: {det}")
            lines.append("")
        if data["idiomas"]:
            lines.append("IDIOMAS")
            for idiom in data["idiomas"]:
                lines.append(f"- {idiom}")
        return "\n".join(lines)

    def _generate(self):
        data = self._collect_data()
        if data is None:
            return

        if self.english_var.get():
            self._english_dialog(data)
            return

        self._do_generate(data)

    def _do_generate(self, data, data_en=None):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        safe_name = data["nombre"].replace(" ", "_")
        safe_name = re.sub(r'[^\w_]', '', safe_name)

        also_english = data_en is not None
        files_created = []

        pdf_es = os.path.join(OUTPUT_DIR, safe_name + ("_CV_ES.pdf" if also_english else "_CV.pdf"))
        docx_es = os.path.join(OUTPUT_DIR, safe_name + ("_CV_ES.docx" if also_english else "_CV.docx"))

        try:
            final_scale = generate_pdf(data, pdf_es, LABELS_ES)
            files_created.append(os.path.basename(pdf_es))
        except Exception as e:
            messagebox.showerror("Error al crear el PDF",
                f"No se pudo generar el PDF:\n{e}\n\nSe creará solo el archivo de Word.")
            final_scale = 1.0

        try:
            generate_word(data, docx_es, final_scale, LABELS_ES)
            files_created.append(os.path.basename(docx_es))
        except Exception as e:
            messagebox.showerror("Error al crear el Word", f"No se pudo generar el Word:\n{e}")
            return

        if also_english:
            safe_en = data_en["nombre"].replace(" ", "_")
            safe_en = re.sub(r'[^\w_]', '', safe_en)
            try:
                scale_en = generate_pdf(data_en, os.path.join(OUTPUT_DIR, safe_en + "_CV_EN.pdf"), LABELS_EN)
                files_created.append(safe_en + "_CV_EN.pdf")
            except Exception:
                scale_en = final_scale
            try:
                generate_word(data_en, os.path.join(OUTPUT_DIR, safe_en + "_CV_EN.docx"), scale_en, LABELS_EN)
                files_created.append(safe_en + "_CV_EN.docx")
            except Exception:
                pass

        file_list = "\n".join(f"  {f}" for f in files_created)
        messagebox.showinfo("Listo! Tu CV fue creado",
            f"Tus archivos están en la carpeta:\n{OUTPUT_DIR}\n\n{file_list}\n\n"
            f"Diseño: {data['theme'].capitalize()}")

        try:
            os.startfile(OUTPUT_DIR)
        except Exception:
            pass

    def _english_dialog(self, data):
        dlg = tk.Toplevel(self.root)
        dlg.title("Versión en inglés")
        dlg.configure(bg=CARD)
        dlg.grab_set()

        w, h = 850, 750
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = max(0, self.root.winfo_y() - 20)
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        # ── Botones ARRIBA para que siempre se vean ──
        btn_f = tk.Frame(dlg, bg="#E8F5E9")
        btn_f.pack(fill="x", side="bottom", padx=0, pady=0)
        btn_inner = tk.Frame(btn_f, bg="#E8F5E9")
        btn_inner.pack(fill="x", padx=20, pady=12)

        def generate():
            en_raw = en_text.get("1.0", "end").strip()
            if not en_raw:
                messagebox.showwarning("Falta el texto en inglés",
                    "Pega el texto traducido antes de generar.", parent=dlg)
                return
            data_en = self._parse_text(en_raw)
            if data_en is None:
                messagebox.showerror("Error al leer",
                    "No se pudo leer el texto en inglés.\nVerifica que tenga el mismo formato.", parent=dlg)
                return
            data_en["theme"] = data["theme"]
            dlg.destroy()
            self._do_generate(data, data_en)

        self._flat_btn(btn_inner, "Generar CV en español e inglés", GREEN, generate, 12).pack(side="right")
        self._flat_btn(btn_inner, "Cancelar", HINT, dlg.destroy, 10).pack(side="right", padx=(0, 10))

        # ── Texto español (solo lectura, para copiar) ──
        tk.Label(dlg, text="Texto en español  —  cópialo y tradúcelo",
                 font=("Segoe UI", 11, "bold"), fg=PRIMARY, bg=CARD).pack(
                     anchor="w", padx=20, pady=(12, 2))
        tk.Label(dlg, text="Selecciona todo (Ctrl+A), copia (Ctrl+C) y pégalo en tu traductor.",
                 font=("Segoe UI", 9), fg=HINT, bg=CARD).pack(anchor="w", padx=20, pady=(0, 4))

        es_frame = tk.Frame(dlg, bg=CARD)
        es_frame.pack(fill="x", padx=20, pady=(0, 6))
        es_scroll = ttk.Scrollbar(es_frame)
        es_scroll.pack(side="right", fill="y")
        es_text = tk.Text(es_frame, font=("Consolas", 9), bg="#F1F5F9", fg=TEXT,
                          relief="solid", bd=1, wrap="word", height=12,
                          yscrollcommand=es_scroll.set)
        es_text.pack(fill="x")
        es_scroll.config(command=es_text.yview)

        spanish_txt = self._data_to_text(data)
        es_text.insert("1.0", spanish_txt)

        # ── Separador + botón traducir ──
        sep_f = tk.Frame(dlg, bg=CARD)
        sep_f.pack(fill="x", padx=20, pady=(4, 4))
        tk.Frame(sep_f, bg=BORDER, height=2).pack(fill="x", side="left", expand=True, pady=(8, 0))

        translate_btn = tk.Button(sep_f, text="Traducir gratis (puede tener errores)",
                                   font=("Segoe UI", 9, "bold"), fg="white", bg="#F59E0B",
                                   activebackground="#D97706", relief="flat", cursor="hand2",
                                   padx=12, pady=4)
        translate_btn.pack(side="right", padx=(10, 0))

        # ── Texto inglés (para pegar traducción) ──
        en_header_f = tk.Frame(dlg, bg=CARD)
        en_header_f.pack(fill="x", padx=20, pady=(4, 2))
        tk.Label(en_header_f, text="Pega aquí el texto ya traducido al inglés",
                 font=("Segoe UI", 11, "bold"), fg="#16A34A", bg=CARD).pack(
                     anchor="w", side="left")
        self._translate_status = tk.Label(en_header_f, text="", font=("Segoe UI", 9),
                                           fg="#F59E0B", bg=CARD)
        self._translate_status.pack(side="right")

        tk.Label(dlg, text="Mismo formato que arriba, solo cambia el contenido al inglés. O usa el botón de traducir.",
                 font=("Segoe UI", 9), fg=HINT, bg=CARD).pack(anchor="w", padx=20, pady=(0, 4))

        en_frame = tk.Frame(dlg, bg=CARD)
        en_frame.pack(fill="both", expand=True, padx=20, pady=(0, 6))
        en_scroll = ttk.Scrollbar(en_frame)
        en_scroll.pack(side="right", fill="y")
        en_text = tk.Text(en_frame, font=("Consolas", 9), bg=INPUT_BG, fg=TEXT,
                          relief="solid", bd=1, wrap="word", insertbackground=TEXT,
                          yscrollcommand=en_scroll.set)
        en_text.pack(fill="both", expand=True)
        en_scroll.config(command=en_text.yview)

        def do_translate():
            translate_btn.config(state="disabled", text="Traduciendo...")
            self._translate_status.config(text="Traduciendo, espera...")
            dlg.update()
            import threading
            def run():
                translated = self._auto_translate(spanish_txt)
                dlg.after(0, lambda: _on_translated(translated))
            def _on_translated(result):
                translate_btn.config(state="normal", text="Traducir gratis (puede tener errores)")
                if result:
                    en_text.delete("1.0", "end")
                    en_text.insert("1.0", result)
                    self._translate_status.config(text="Traducción lista. Revisa antes de generar.", fg="#16A34A")
                else:
                    err = getattr(self, '_translate_error', None) or "desconocido"
                    self._translate_status.config(text=f"Error: {err}", fg=RED)
                    messagebox.showwarning("No se pudo traducir",
                        f"La traducción automática falló:\n{err}\n\n"
                        "Copia el texto en español, tradúcelo manualmente\n"
                        "(Google Translate, DeepL, ChatGPT, etc.) y pégalo\n"
                        "en el campo de inglés.", parent=dlg)
            threading.Thread(target=run, daemon=True).start()

        translate_btn.config(command=do_translate)

    def _parse_text(self, raw_text):
        return parse_text(raw_text)

    # ── Auto-translate (Google Translate + MyMemory fallback) ──

    def _auto_translate(self, spanish_txt):
        import urllib.request
        import urllib.parse
        import urllib.error
        import json
        import time
        import unicodedata

        KEYWORDS = {
            "DATOS PERSONALES", "NOMBRE:", "CONTACTO:", "TITULO:",
            "RESUMEN PROFESIONAL", "LOGROS CLAVE",
            "EXPERIENCIA PROFESIONAL", "PUESTO:", "EMPRESA:", "FECHAS:",
            "EDUCACION", "CARRERA:", "ESCUELA:",
            "CURSOS Y CERTIFICACIONES", "HABILIDADES", "IDIOMAS",
        }

        def _is_keyword_line(line):
            stripped = line.strip()
            for kw in KEYWORDS:
                if stripped == kw or stripped.startswith(kw):
                    return kw
            return None

        def _normalize(t):
            t = unicodedata.normalize("NFKD", t)
            t = "".join(c for c in t if not unicodedata.combining(c))
            return re.sub(r'[^a-z0-9 ]', '', t.lower()).strip()

        def _is_same_text(src, tgt):
            return _normalize(src) == _normalize(tgt)

        # ── API 1: Google Translate (gratuito, sin API key) ──
        def _call_google(text):
            url = "https://translate.googleapis.com/translate_a/single"
            params = urllib.parse.urlencode({
                "client": "gtx", "sl": "es", "tl": "en", "dt": "t", "q": text,
            })
            req = urllib.request.Request(f"{url}?{params}", headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            })
            with urllib.request.urlopen(req, timeout=25) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                parts = []
                for chunk in data[0]:
                    if chunk[0]:
                        parts.append(chunk[0])
                return "".join(parts)

        # ── API 2: MyMemory (fallback) ──
        def _call_mymemory(text):
            url = ("https://api.mymemory.translated.net/get?"
                   + urllib.parse.urlencode({"q": text, "langpair": "es|en"}))
            req = urllib.request.Request(url, headers={"User-Agent": "CVGenerator/1.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                status = data.get("responseStatus", 0)
                if status != 200:
                    raise Exception(f"MyMemory status {status}: {data.get('responseDetails', '')}")
                translated = data["responseData"]["translatedText"]
                if not translated or not translated.strip():
                    raise Exception("MyMemory devolvió respuesta vacía")
                for marker in ["PLEASE CONTACT", "MYMEMORY WARNING", "QUERY LENGTH LIMIT"]:
                    if marker in translated.upper():
                        raise Exception(f"MyMemory error: {translated[:80]}")
                return translated

        # ── Traduce un batch de textos, intenta Google primero, MyMemory si falla ──
        def _translate_batch(texts):
            if not texts:
                return texts

            combined = "\n".join(texts)
            errors = []

            # Intento 1: Google Translate
            for attempt in range(2):
                try:
                    translated = _call_google(combined)
                    if translated and translated.strip():
                        parts = [p.strip() for p in translated.split("\n")]
                        if len(parts) == len(texts):
                            if not all(_is_same_text(s, t) for s, t in zip(texts, parts)):
                                return parts
                        if len(texts) == 1 and not _is_same_text(texts[0], translated.strip()):
                            return [translated.strip()]
                        if len(parts) > len(texts):
                            joined = " ".join(parts)
                            if not _is_same_text(combined, joined) and len(texts) == 1:
                                return [joined]
                    errors.append("Google: devolvió texto sin traducir")
                    break
                except urllib.error.HTTPError as e:
                    errors.append(f"Google HTTP {e.code}")
                    if e.code == 429 and attempt == 0:
                        time.sleep(2)
                        continue
                    break
                except Exception as e:
                    errors.append(f"Google: {e}")
                    if attempt == 0:
                        time.sleep(1)
                        continue
                    break

            # Intento 2: MyMemory (uno por uno para evitar problemas de split)
            try:
                results = []
                for i, txt in enumerate(texts):
                    translated = _call_mymemory(txt)
                    if _is_same_text(txt, translated):
                        errors.append(f"MyMemory: no tradujo fragmento {i+1}")
                        self._translate_error = " | ".join(errors)
                        return None
                    results.append(translated)
                    if i < len(texts) - 1:
                        time.sleep(0.5)
                return results
            except Exception as e:
                errors.append(f"MyMemory: {e}")

            self._translate_error = " | ".join(errors)
            return None

        SECTION_TRANSLATIONS = {
            "DATOS PERSONALES": "PERSONAL INFORMATION",
            "RESUMEN PROFESIONAL": "PROFESSIONAL SUMMARY",
            "LOGROS CLAVE": "KEY ACHIEVEMENTS",
            "EXPERIENCIA PROFESIONAL": "PROFESSIONAL EXPERIENCE",
            "EDUCACION": "EDUCATION",
            "CURSOS Y CERTIFICACIONES": "CERTIFICATIONS & COURSES",
            "HABILIDADES": "SKILLS",
            "IDIOMAS": "LANGUAGES",
        }
        KW_TRANSLATIONS = {
            "NOMBRE:": "NAME:",
            "CONTACTO:": "CONTACT:",
            "TITULO:": "TITLE:",
            "PUESTO:": "POSITION:",
            "EMPRESA:": "COMPANY:",
            "FECHAS:": "DATES:",
            "CARRERA:": "DEGREE:",
            "ESCUELA:": "SCHOOL:",
        }
        NO_TRANSLATE_KW = {"NOMBRE:", "CONTACTO:", "FECHAS:", "ESCUELA:"}

        self._translate_error = None
        try:
            lines = spanish_txt.split("\n")
            parsed = []
            for line in lines:
                if not line.strip():
                    parsed.append(("empty", line, None))
                    continue
                kw = _is_keyword_line(line)
                if kw:
                    en_kw = KW_TRANSLATIONS.get(kw, kw)
                    if kw in SECTION_TRANSLATIONS:
                        parsed.append(("section", SECTION_TRANSLATIONS[kw], None))
                    elif kw in NO_TRANSLATE_KW:
                        value = line.strip()[len(kw):].strip()
                        parsed.append(("keep", f"{en_kw} {value}" if value else en_kw, None))
                    else:
                        value = line.strip()[len(kw):].strip()
                        if value:
                            parsed.append(("kw_translate", en_kw, value))
                        else:
                            parsed.append(("keep", en_kw, None))
                elif line.strip().startswith("- "):
                    parsed.append(("bullet", None, line.strip()[2:]))
                else:
                    parsed.append(("text", None, line.strip()))

            to_translate = []
            indices = []
            for i, (kind, _, val) in enumerate(parsed):
                if kind in ("kw_translate", "bullet", "text") and val:
                    to_translate.append(val)
                    indices.append(i)

            if not to_translate:
                self._translate_error = "No se encontró texto para traducir"
                return None

            MAX_CHARS = 4500
            translated_vals = []
            batch = []
            batch_len = 0
            for txt in to_translate:
                added_len = len(txt) + (1 if batch else 0)
                if batch and batch_len + added_len > MAX_CHARS:
                    result = _translate_batch(batch)
                    if result is None:
                        return None
                    translated_vals.extend(result)
                    batch = []
                    batch_len = 0
                    time.sleep(0.5)
                batch.append(txt)
                batch_len += added_len
            if batch:
                result = _translate_batch(batch)
                if result is None:
                    return None
                translated_vals.extend(result)

            # Validar que al menos la mayoría se tradujo
            same_count = sum(1 for s, t in zip(to_translate, translated_vals) if _is_same_text(s, t))
            if same_count > len(to_translate) * 0.5:
                self._translate_error = f"Solo se tradujeron {len(to_translate) - same_count} de {len(to_translate)} fragmentos"
                return None

            for j, idx in enumerate(indices):
                if j < len(translated_vals):
                    kind, kw, _ = parsed[idx]
                    parsed[idx] = (kind, kw, translated_vals[j])

            result_lines = []
            for kind, a, b in parsed:
                if kind == "empty":
                    result_lines.append(a)
                elif kind in ("section", "keep"):
                    result_lines.append(a)
                elif kind == "kw_translate":
                    result_lines.append(f"{a} {b}")
                elif kind == "bullet":
                    result_lines.append(f"- {b}")
                elif kind == "text":
                    result_lines.append(b)

            return "\n".join(result_lines)
        except Exception as e:
            self._translate_error = str(e)
            return None

    # ── Paste text dialog ──

    def _paste_text_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Pegar texto del cliente")
        dlg.configure(bg=CARD)
        dlg.grab_set()

        w, h = 700, 620
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + 30
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(dlg, text="Pega aqui el texto del cliente",
                 font=("Segoe UI", 14, "bold"), fg=PRIMARY, bg=CARD).pack(
                     anchor="w", padx=20, pady=(18, 4))
        tk.Label(dlg, text="Borra el ejemplo de abajo y pega el texto con el mismo formato. Luego dale 'Cargar al formulario'.",
                 font=("Segoe UI", 9), fg=HINT, bg=CARD).pack(anchor="w", padx=20, pady=(0, 10))

        text_frame = tk.Frame(dlg, bg=CARD)
        text_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        text_box = tk.Text(text_frame, font=("Consolas", 10), bg=INPUT_BG, fg=TEXT,
                           relief="solid", bd=1, wrap="word", insertbackground=TEXT,
                           yscrollcommand=scrollbar.set)
        text_box.pack(fill="both", expand=True)
        scrollbar.config(command=text_box.yview)

        example = (
            "DATOS PERSONALES\n"
            "NOMBRE: Alexis Jehu Alvarez Sanchez\n"
            "CONTACTO: +52 334 648 8200 | alexisjehu@gmail.com | linkedin: alexisjehualvarezsanchez\n"
            "TITULO: Ingeniero Mecatrónico | Desarrollador .NET y C# | Integración MES\n"
            "\n"
            "RESUMEN PROFESIONAL\n"
            "Ingeniero Mecatrónico especializado en aplicaciones .NET y C# para sistemas\n"
            "industriales y empresariales. Sólida experiencia en automatización, conectando\n"
            "hardware y software para construir soluciones robustas y escalables. Experiencia\n"
            "en integración MES, protocolos de comunicación industrial y desarrollo de APIs REST.\n"
            "\n"
            "LOGROS CLAVE\n"
            "- Integración de máquinas AOI, SPI y DPI con sistemas MES para control de calidad SMT\n"
            "- Desarrollo de aplicaciones de escritorio con WinForms para control de automatización\n"
            "- Implementación de estándares IPC-CFX, Hermes e iTAC en integración MES\n"
            "- Diseño de sistemas de chatbot con integración backend usando Python\n"
            "- Soporte a sistemas robóticos Fanuc, Epson y Cartesianos en línea de producción\n"
            "\n"
            "EXPERIENCIA PROFESIONAL\n"
            "PUESTO: Desarrollador de Software / Ingeniero MES\n"
            "EMPRESA: Koh Young Technology\n"
            "FECHAS: Nov 2024 - Actual\n"
            "- Desarrollo y mantenimiento de scripts en C# para automatizar procesos industriales y mejorar la funcionalidad del sistema MES\n"
            "- Integración de máquinas AOI, SPI y DPI con sistemas MES para mejorar el control de calidad SMT\n"
            "- Diseño y consumo de APIs REST usando JSON y XML para interoperabilidad de sistemas\n"
            "- Soporte y mantenimiento de bases de datos SQL para asegurar integridad de datos y optimizar consultas\n"
            "- Gestión de protocolos de comunicación (TCP/IP, RS-232, RS-485) para conexiones estables entre dispositivos\n"
            "- Implementación de estándares IPC-CFX, Hermes e iTAC en integración MES\n"
            "\n"
            "PUESTO: Ingeniero de Software\n"
            "EMPRESA: New Power Technology Monterrey\n"
            "FECHAS: Mar 2024 - Nov 2024\n"
            "- Construcción de aplicaciones de escritorio con WinForms (.NET Framework y C#) para control de automatización\n"
            "- Integración con robots, PLCs y dispositivos industriales usando protocolos TCP/IP, Serial y Modbus\n"
            "- Uso de VisionPro para sistemas de inspección visual y validación en tiempo real\n"
            "- Creación de APIs REST y dashboards visuales para métricas de producción usando JSON\n"
            "- Gestión de control de versiones con Git y Azure Repos\n"
            "\n"
            "PUESTO: Desarrollador de Chatbots\n"
            "EMPRESA: Apolo Automation (Proyecto Personal)\n"
            "FECHAS: Feb 2023 - Dic 2023\n"
            "- Diseño de sistemas de chatbot con flujos lógicos e integración backend usando Python\n"
            "- Desarrollo de APIs y scripts automatizados para procesamiento de datos (CSV, JSON)\n"
            "- Aplicación de principios POO para modularizar flujos de automatización\n"
            "\n"
            "PUESTO: Técnico en Automatización\n"
            "EMPRESA: Jabil Guadalajara\n"
            "FECHAS: Ago 2022 - Feb 2023\n"
            "- Soporte a sistemas robóticos (Fanuc, Epson, Cartesianos) con configuración y mantenimiento\n"
            "- Depuración de software en C# para resolver problemas en línea de producción\n"
            "- Trabajo con sistemas de visión industrial Cognex y Keyence\n"
            "\n"
            "EDUCACION\n"
            "CARRERA: Ingeniería Mecatrónica\n"
            "ESCUELA: Instituto Tecnológico de Colima | 2018 - 2023\n"
            "\n"
            "CARRERA: Técnico Analista Programador\n"
            "ESCUELA: Universidad de Colima | 2015 - 2018\n"
            "\n"
            "CURSOS Y CERTIFICACIONES\n"
            "- Diplomado Técnico en IoT\n"
            "- Desarrollo de Software Embebido\n"
            "- Diseño Mecánico CSWA\n"
            "- Fundamentos de CAE (Ingeniería Asistida por Computadora)\n"
            "- Innovación en Ingeniería (ESSS)\n"
            "- Fundamentos de Electrónica Digital\n"
            "- Comunicaciones Inalámbricas / Electrónica / Técnico en Redes de Datos\n"
            "\n"
            "HABILIDADES\n"
            "Lenguajes: C#, C, C++, Python\n"
            "Tecnologías: .NET Framework, WinForms, REST API, SOAP, SQL Server\n"
            "Protocolos Industriales: IPC-CFX, Hermes, iTAC, TCP/IP, RS-232, RS-485, Modbus, SECS/GEM\n"
            "Herramientas y Plataformas: Visual Studio, Git, Azure DevOps/Repos, VisionPro, Android VM, Linux\n"
            "Otras Habilidades: POO/DOO, Depuración y Mantenimiento de Software, Sistemas de Control y Automatización\n"
            "\n"
            "IDIOMAS\n"
            "- Español - Nativo\n"
            "- Inglés - Profesional (técnico y de negocios)\n"
        )
        text_box.insert("1.0", example)

        def load_text():
            raw = text_box.get("1.0", "end").strip()
            if not raw:
                messagebox.showwarning("Sin texto", "Pega el texto del cliente antes de cargar.", parent=dlg)
                return
            self._load_from_text(raw)
            dlg.destroy()
            messagebox.showinfo("Datos cargados",
                "Se llenaron todos los campos con el texto del cliente.\n\n"
                "Revisa que todo este bien y dale 'Generar mi CV'.")

        btn_f = tk.Frame(dlg, bg=CARD)
        btn_f.pack(fill="x", padx=20, pady=(0, 18))
        self._flat_btn(btn_f, "Cargar al formulario", GREEN, load_text, 11).pack(side="right")
        self._flat_btn(btn_f, "Cancelar", HINT, dlg.destroy, 10).pack(side="right", padx=(0, 8))

    def _load_from_text(self, raw_text):
        data = parse_text(raw_text)
        self._fill_form(data)

    def _fill_form(self, data):
        theme = data.get("theme", "clasico")
        self.template_var.set(theme)

        self.nombre_entry.set_value(data.get("nombre", ""))
        self.contacto_entry.set_value(data.get("contacto", ""))
        self.titulo_entry.set_value(data.get("titulo", ""))

        self.resumen_text.set_value(data.get("resumen", ""))
        self.logros_text.set_value("\n".join(data.get("logros", [])))

        self.experiences = data.get("experiencia", [])
        self._refresh_exp_list()

        self.education_list = data.get("educacion", [])
        self._refresh_edu_list()

        self.cursos_text.set_value("\n".join(data.get("cursos", [])))

        hab_lines = [f"{cat}: {det}" for cat, det in data.get("habilidades", [])]
        self.habilidades_text.set_value("\n".join(hab_lines))

        self.idiomas_text.set_value("\n".join(data.get("idiomas", [])))

    def _load_if_exists(self):
        if os.path.exists(DATA_FILE):
            try:
                self._load_from_file(DATA_FILE, silent=True)
            except Exception:
                pass

    def _load_from_file(self, path, silent=False):
        from cv_generator import parse_data
        try:
            data = parse_data(path)
        except Exception:
            return
        if data["nombre"] == "NOMBRE COMPLETO DEL CANDIDATO":
            return
        self._fill_form(data)


def main():
    root = tk.Tk()
    CVGeneratorApp(root)
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
