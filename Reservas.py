import tkinter as tk
from tkinter import messagebox, simpledialog, scrolledtext, filedialog, ttk
import csv
from datetime import datetime
import os
import sys
import shutil
import pandas as pd

# =======================================
#         CONFIGURACIÓN Y COLORES
# =======================================
reservas_file = "reservas.csv"
bloques_file = "bloques.txt"
backup_dir = "backups"
observaciones_file = "observaciones.csv"
logo_path = "logo_escuela.ico"
modo_oscuro = False
logueado_encargado = False

COLOR_CLARO_BG = "#f7f7f7"
COLOR_CLARO_PANEL = "#ffffff"
COLOR_CLARO_SIDE = "#e8eaf6"
COLOR_CLARO_BTN = "#64b5f6"
COLOR_CLARO_BTN_HOVER = "#42a5f5"
COLOR_CLARO_BTN_TXT = "#fff"
COLOR_CLARO_RESERVADO = "#f8bbbb"
COLOR_CLARO_DISPONIBLE = "#a8e6cf"
COLOR_CLARO_RESALTADO = "#f6e58d"

COLOR_OSCURO_BG = "#181d24"
COLOR_OSCURO_PANEL = "#222933"
COLOR_OSCURO_SIDE = "#252c39"
COLOR_OSCURO_BTN = "#5dade2"
COLOR_OSCURO_BTN_TXT = "#f7f7f7"
COLOR_OSCURO_RESERVADO = "#e57373"
COLOR_OSCURO_DISPONIBLE = "#66c76b"
COLOR_OSCURO_RESALTADO = "#f6e58d"
COLOR_OSCURO_BTN_HOVER = "#4594c9"

FUENTE = ("Segoe UI", 11)
FUENTE_TITULO = ("Segoe UI", 16, "bold")
FUENTE_PANEL = ("Segoe UI", 13, "bold")
FUENTE_BTN = ("Segoe UI", 11)
FUENTE_SIDE = ("Segoe UI", 11, "bold")
FUENTE_PEQUENA = ("Segoe UI", 9)

# =======================================
#             AUXILIARES
# =======================================
def resource_path(relative_path):
    """
    Si es un recurso solo de lectura (ej: logo), usa _MEIPASS.
    Para archivos que deben mantenerse tras cerrar el programa, usa la carpeta del ejecutable.
    """
    if hasattr(sys, '_MEIPASS'):
        # Solo para archivos de solo lectura (imágenes, etc.)
        if relative_path.endswith(".png"):
            base_path = sys._MEIPASS
        else:
            # Para archivos de datos que deben mantenerse (csv/txt)
            base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


reservas_file = resource_path("reservas.csv")
bloques_file = resource_path("bloques.txt")
backup_dir = resource_path("backups")
observaciones_file = resource_path("observaciones.csv")
logo_path = resource_path("logo_escuela.png")

def cargar_bloques():
    if not os.path.exists(bloques_file):
        with open(bloques_file, 'w', encoding='utf-8') as f:
            f.write("08:15 - 09:45\n10:05 - 11:35\n11:50 - 13:20\n14:05 - 15:35")
    with open(bloques_file, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f.readlines() if line.strip()]

bloques = cargar_bloques()

def cargar_reservas():
    try:
        with open(reservas_file, mode='r', newline='', encoding='utf-8') as f:
            return list(csv.reader(f))
    except FileNotFoundError:
        return []

def cargar_observaciones():
    if not os.path.exists(observaciones_file):
        return []
    with open(observaciones_file, mode='r', newline='', encoding='utf-8') as f:
        return list(csv.reader(f))

def guardar_reserva(fecha, bloque, profesor, curso, asignatura, objetivo):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(reservas_file, mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([fecha, bloque, profesor.strip(), curso.strip(), asignatura.strip(), objetivo.strip(), timestamp])

def guardar_observacion(fecha, bloque, profesor, observacion):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(observaciones_file, mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([fecha, bloque, profesor, observacion.strip(), timestamp])

def obtener_observacion(fecha, bloque, profesor):
    observaciones = cargar_observaciones()
    for r in reversed(observaciones):
        if r[0] == fecha and r[1] == bloque and r[2] == profesor:
            return r[3]
    return ""

def esta_reservado(reservas, fecha, bloque):
    for r in reservas:
        if r[0] == fecha and r[1] == bloque:
            return True, r[2], r[3], r[4]
    return False, "", "", ""

def backup_diario():
    os.makedirs(backup_dir, exist_ok=True)
    hoy = datetime.now().strftime("%Y-%m-%d")
    backup_path = os.path.join(backup_dir, f"reservas_backup_{hoy}.csv")
    obs_backup_path = os.path.join(backup_dir, f"observaciones_backup_{hoy}.csv")
    if not os.path.exists(backup_path) and os.path.exists(reservas_file):
        shutil.copy(reservas_file, backup_path)
    if not os.path.exists(obs_backup_path) and os.path.exists(observaciones_file):
        shutil.copy(observaciones_file, obs_backup_path)

def add_hover_effect(widget, normal, hover):
    def on_enter(e): widget['bg'] = hover
    def on_leave(e): widget['bg'] = normal
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)

def exportar_excel():
    reservas = cargar_reservas()
    if not reservas:
        messagebox.showinfo("Exportar", "No hay datos para exportar.")
        return
    columnas = ["Fecha", "Bloque", "Profesor", "Curso", "Asignatura", "Objetivo", "Hora"]
    datos = [r[:7] for r in reservas]
    if logueado_encargado:
        columnas.append("Observación")
        obs_dict = {}
        for obs in cargar_observaciones():
            obs_dict[(obs[0], obs[1], obs[2])] = obs[3]
        datos = [row + [obs_dict.get((row[0], row[1], row[2]), "")] for row in datos]
    df = pd.DataFrame(datos, columns=columnas)
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
    if archivo:
        try:
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Exportar", "Datos exportados exitosamente.")
        except ImportError:
            messagebox.showerror("Error", "No se pudo exportar. Instala el módulo 'openpyxl'.")

def actualizar_fecha_hora():
    ahora = datetime.now().strftime("%d/%m/%Y   %I:%M %p")
    fecha_label.config(text=ahora)
    fecha_label.after(1000, actualizar_fecha_hora)

def centrar_ventana(ventana):
    ventana.update_idletasks()
    ancho = ventana.winfo_width()
    alto = ventana.winfo_height()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"+{x}+{y}")

# =======================================
#     TOGGLE SWITCH MATERIAL DESIGN
# =======================================
class ToggleSwitch(tk.Frame):
    def __init__(self, master, width=46, height=24, command=None, state=False, **kw):
        super().__init__(master, width=width, height=height, **kw)
        self["bg"] = kw.get("bg", "#fff")
        self.width = width
        self.height = height
        self.command = command
        self._on = state
        self.canvas = tk.Canvas(self, width=width, height=height, bg=self["bg"], bd=0, highlightthickness=0)
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.toggle)
        self.draw()
    def draw(self):
        self.canvas.delete("all")
        self.canvas.configure(bg=self["bg"])
        if self._on:
            color = "#3a6ea5"
            knob = self.width - 22
        else:
            color = "#ccc"
            knob = 2
        self.canvas.create_rounded_rect(2, 2, self.width-2, self.height-2, 11, fill=color, outline="")
        self.canvas.create_oval(knob, 2, knob+20, self.height-2, fill="#fff", outline="")
    def toggle(self, event=None):
        self._on = not self._on
        self.draw()
        if self.command:
            self.command(self._on)
    def get(self):
        return self._on
    def set(self, value):
        self._on = value
        self.draw()
def _create_rounded_rect(self, x1, y1, x2, y2, r=25, **kwargs):
    points = [
        x1+r, y1,
        x2-r, y1,
        x2, y1,
        x2, y1+r,
        x2, y2-r,
        x2, y2,
        x2-r, y2,
        x1+r, y2,
        x1, y2,
        x1, y2-r,
        x1, y1+r,
        x1, y1
    ]
    return self.create_polygon(points, smooth=True, **kwargs)
tk.Canvas.create_rounded_rect = _create_rounded_rect

# =======================================
#         LÓGICA PRINCIPAL
# =======================================

# ----- FORMULARIO DE RESERVA -----
def mostrar_formulario(bloque):
    reservas = cargar_reservas()
    fecha = datetime.now().strftime("%Y-%m-%d")
    reservado, _, _, _ = esta_reservado(reservas, fecha, bloque)
    if reservado:
        messagebox.showerror("Error", f"El bloque '{bloque}' ya está reservado.")
        return

    color_bg = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    color_btn = COLOR_OSCURO_BTN if modo_oscuro else COLOR_CLARO_BTN
    color_btn_hover = COLOR_OSCURO_BTN_HOVER if modo_oscuro else COLOR_CLARO_BTN_HOVER
    color_btn_txt = COLOR_OSCURO_BTN_TXT if modo_oscuro else COLOR_CLARO_BTN_TXT
    color_label = "#fff" if modo_oscuro else "#222"
    color_chk = "#fff" if modo_oscuro else "#222"
    color_reglamento = "#71a7ff" if modo_oscuro else "#2e6fbb"

    form = tk.Toplevel(root)
    form.title("Formulario de Reserva")
    form.geometry("440x350")
    form.resizable(False, False)
    form.configure(bg=color_bg)
    centrar_ventana(form)
    form.transient(root)
    form.grab_set()
    form.focus_set()

    tk.Label(form, text=f"Reservar bloque: {bloque}", font=FUENTE_PANEL, bg=color_bg, fg=color_label).pack(pady=(18, 8))

    frame = tk.Frame(form, bg=color_bg)
    frame.pack(pady=0)
    tk.Label(frame, text="Docente:", font=FUENTE, bg=color_bg, fg=color_label).grid(row=0, column=0, sticky="w", pady=5, padx=7)
    entry_profesor = tk.Entry(frame, font=FUENTE, width=26)
    entry_profesor.grid(row=0, column=1, pady=5)
    tk.Label(frame, text="Curso:", font=FUENTE, bg=color_bg, fg=color_label).grid(row=1, column=0, sticky="w", pady=5, padx=7)
    entry_curso = tk.Entry(frame, font=FUENTE, width=26)
    entry_curso.grid(row=1, column=1, pady=5)
    tk.Label(frame, text="Asignatura:", font=FUENTE, bg=color_bg, fg=color_label).grid(row=2, column=0, sticky="w", pady=5, padx=7)
    entry_asignatura = tk.Entry(frame, font=FUENTE, width=26)
    entry_asignatura.grid(row=2, column=1, pady=5)

    # NUEVO: Objetivo de la clase
    tk.Label(frame, text="Objetivo de la clase:", font=FUENTE, bg=color_bg, fg=color_label).grid(row=3, column=0, sticky="nw", pady=5, padx=7)
    entry_objetivo = tk.Entry(frame, font=FUENTE, width=26)
    entry_objetivo.grid(row=3, column=1, pady=5, sticky="w")

    acepta = tk.IntVar(value=0)
    casilla_frame = tk.Frame(form, bg=color_bg)
    casilla_frame.pack(pady=(14, 0))
    chk = tk.Checkbutton(
        casilla_frame,
        variable=acepta, bg=color_bg, highlightthickness=0,
        activebackground=color_bg, fg=color_chk, selectcolor=color_bg
    )
    chk.grid(row=0, column=0, sticky="n")
    lbl_chk = tk.Label(
        casilla_frame,
        text="He leído y acepto el Reglamento de Uso de la Sala de Enlaces",
        font=("Segoe UI", 10),
        bg=color_bg, fg=color_label, anchor="w", justify="center", wraplength=320
    )
    lbl_chk.grid(row=0, column=1, sticky="w")
    regl_btn = tk.Button(
        form, text="Leer reglamento",
        command=lambda: mostrar_reglamento(form),
        font=("Segoe UI", 10, "underline"), fg=color_reglamento,
        bg=color_bg, borderwidth=0, cursor="hand2", activebackground=color_bg,
        activeforeground=color_reglamento
    )
    regl_btn.pack(pady=(10, 0), anchor="center")

    btn_frame = tk.Frame(form, bg=color_bg)
    btn_frame.pack(pady=(10, 8))
    def reservar():
        nombre = entry_profesor.get().strip()
        curso = entry_curso.get().strip()
        asignatura = entry_asignatura.get().strip()
        objetivo = entry_objetivo.get().strip()
        if not nombre or len(nombre.split()) < 2:
            messagebox.showwarning("Faltan datos", "Debes ingresar **nombre y apellido** del docente.")
            return
        if not curso or not asignatura or not objetivo:
            messagebox.showwarning("Faltan datos", "Completa todos los campos.")
            return
        if not acepta.get():
            messagebox.showwarning("Falta aceptar", "Debes aceptar el reglamento.")
            return
        guardar_reserva(fecha, bloque, nombre, curso, asignatura, objetivo)
        messagebox.showinfo(
            "Reserva hecha",
            f"Reserva registrada para el docente **{nombre}**\n"
            f"Bloque: {bloque}\nCurso: {curso}\nAsignatura: {asignatura}\nObjetivo: {objetivo}\n\n"
            "¡Gracias por tu reserva!"
        )
        form.destroy()
        actualizar_bloques()
    reservar_btn = tk.Button(btn_frame, text="Reservar", font=FUENTE_BTN, width=12,
        bg=color_btn, fg=color_btn_txt, activebackground="#27ae60",
        relief="flat", command=reservar)
    reservar_btn.pack(side="left", padx=(0,10))
    add_hover_effect(reservar_btn, color_btn, color_btn_hover)
    cancelar_btn = tk.Button(btn_frame, text="Cancelar", font=FUENTE_BTN, width=10,
        bg="#e5e7e9", fg="#333", relief="flat", command=form.destroy)
    cancelar_btn.pack(side="left")
    form.pack_slaves()

# ----- REGLAMENTO -----
def mostrar_reglamento(parent):
    reglamento = """REGLAMENTO DE USO DE LA SALA DE ENLACES 
Escuela República de Grecia Llallauquén 

Introducción: 

El presente reglamento tiene como finalidad asegurar un uso responsable, seguro y eficiente de la Sala de Enlaces, resguardando los recursos tecnológicos y propiciando un ambiente óptimo para el aprendizaje. Las siguientes normas han sido establecidas para prevenir daños, garantizar la continuidad del servicio y promover el respeto y la colaboración entre todos los usuarios. 

1. Normas Generales de Uso 

- El uso de la Sala de Enlaces está destinado exclusivamente a actividades pedagógicas previamente planificadas y autorizadas. 
- Los(as) docentes deberán reservar la sala con anticipación mediante el sistema habilitado para ello. 
- El docente responsable debe estar presente durante toda la sesión y supervisar el comportamiento de los(as) estudiantes. 
- Es responsabilidad del docente velar por el uso adecuado del equipamiento y del espacio. 
- Está estrictamente prohibido consumir alimentos y bebidas dentro de la sala. 
- No se permite la instalación de programas ni la modificación de la configuración de los equipos sin autorización expresa. 
- Al término de la sesión, el docente deberá avisar al encargado enlaces de cualquier anomalía o incidente que pudo haber transcurrido durante la clase para ser registrado en el sistema, además deberá entregar la sala en las mismas condiciones en que la recibió, limpia y ordenada. 

2. Deberes de los(as) Usuarios(as) 
- Cuidar y hacer buen uso de los computadores, mobiliario y demás recursos disponibles. 
- Respetar estrictamente los horarios asignados para cada bloque. 
- Reportar de inmediato cualquier falla o daño al encargado de la Sala de Enlaces. 
 
3. Prohibiciones 
- Está prohibido el acceso a redes sociales, páginas inapropiadas o ajenas a la actividad pedagógica. 
- No está permitido manipular cables, conexiones o periféricos sin autorización. 
- Queda prohibido cambiar el mobiliario de lugar. 
 
4. Encargado de la Sala de Enlaces 
El encargado de la Sala de Enlaces es responsable de: 
- Supervisar el cumplimiento de este reglamento. 
- Registrar observaciones o incidentes en el sistema. 
- Coordinar el mantenimiento y correcto funcionamiento de los equipos. 
 
¡Gracias por su colaboración! 
 
Kevin Albanez Palacios 
Encargado Enlaces """
    color_bg = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    reglamento_win = tk.Toplevel(parent)
    reglamento_win.title("Reglamento de Uso de la Sala de Enlaces")
    reglamento_win.geometry("540x480")
    reglamento_win.resizable(False, False)
    centrar_ventana(reglamento_win)
    reglamento_win.configure(bg=color_bg)
    reglamento_win.transient(root)
    reglamento_win.grab_set()
    reglamento_win.focus_set()


    if os.path.exists(logo_path):
        from PIL import Image, ImageTk
        logo_img = Image.open(logo_path)
        logo_img = logo_img.resize((80, 80))
        logo = ImageTk.PhotoImage(logo_img)
        lbl_logo = tk.Label(reglamento_win, image=logo, bg=color_bg)
        lbl_logo.image = logo
        lbl_logo.pack(pady=(8, 2))

    color_txt = "#fff" if modo_oscuro else "#222"
    lbl = tk.Label(reglamento_win, text="REGLAMENTO DE USO DE LA SALA DE ENLACES", font=FUENTE_PANEL, bg=color_bg, fg=color_txt)
    lbl.pack(pady=(0, 8))
    txt = scrolledtext.ScrolledText(reglamento_win, wrap=tk.WORD, font=("Segoe UI", 10), height=21, width=62, bg="#f8f8fa")
    txt.insert(tk.END, reglamento)
    txt.config(state="disabled")
    txt.pack(padx=10, pady=(0, 6))
    tk.Button(reglamento_win, text="Cerrar", command=reglamento_win.destroy, width=14,
              bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, font=("Segoe UI", 10, "bold")).pack(pady=(2, 10))

# ----- HISTORIAL -----
def ver_historial():
    historial = cargar_reservas()
    historial = sorted(historial, key=lambda r: r[6] if len(r) > 6 else "0000-00-00 00:00:00", reverse=True)
    color_bg = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    ventana = tk.Toplevel(root)
    ventana.title("Historial de Reservas")
    ventana.geometry("1060x460")
    ventana.resizable(False, False)
    centrar_ventana(ventana)
    ventana.configure(bg=color_bg)
    color_txt = "#fff" if modo_oscuro else "#222"
    ventana.transient(root)
    ventana.grab_set()
    ventana.focus_set()

    filtro_frame = tk.Frame(ventana, bg=color_bg)
    filtro_frame.pack(fill="x", padx=10, pady=6)
    tk.Label(filtro_frame, text="Filtrar por nombre:", font=FUENTE, bg=color_bg, fg=color_txt).pack(side="left")
    filtro_entry = tk.Entry(filtro_frame, font=FUENTE)
    filtro_entry.pack(side="left", padx=5)

    # AHORA INCLUYE EL OBJETIVO
    columns = ["Fecha", "Bloque", "Profesor", "Curso", "Asignatura", "Objetivo", "Hora"]
    if logueado_encargado:
        columns.append("Observación")
    tree = ttk.Treeview(ventana, columns=columns, show="headings", height=18)
    for col in columns:
        # Ajusta el ancho de columna para el objetivo y observación
        if col == "Objetivo":
            tree.column(col, width=220)
        elif col == "Observación":
            tree.column(col, width=150)
        else:
            tree.column(col, width=110)
        tree.heading(col, text=col)
    tree.pack(expand=True, fill="both", padx=8, pady=(0, 6))

    def obtener_obs(f, b, p):
        return obtener_observacion(f, b, p) if logueado_encargado else ""

    def aplicar_filtro(*args):
        filtro = filtro_entry.get().strip().lower()
        tree.delete(*tree.get_children())
        for r in historial:
            if len(r) >= 7 and filtro in r[2].lower():
                values = r[:7]  # incluye objetivo
                if logueado_encargado:
                    obs = obtener_obs(r[0], r[1], r[2])
                    values = list(values) + [obs]
                tree.insert("", tk.END, values=values)

    filtro_entry.bind("<KeyRelease>", aplicar_filtro)
    aplicar_filtro()

    if logueado_encargado:
        frame_btns = tk.Frame(ventana, bg=color_bg)
        frame_btns.pack(pady=(0, 8))
        tk.Button(frame_btns, text="Exportar a Excel", command=exportar_excel, width=19, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT).pack(side="left", padx=7)
        tk.Button(frame_btns, text="Añadir Observación", command=lambda: agregar_obs_historial(tree), width=19, bg="#ffd54f").pack(side="left", padx=7)

def agregar_obs_historial(tree):
    selected = tree.focus()
    if not selected:
        messagebox.showinfo("Seleccionar reserva", "Seleccione una reserva del historial para agregar una observación.")
        return
    values = tree.item(selected, "values")
    if len(values) < 3:
        messagebox.showerror("Error", "No se pudo obtener la reserva seleccionada.")
        return
    fecha, bloque, profesor = values[0], values[1], values[2]

    color_bg = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    color_btn = COLOR_OSCURO_BTN if modo_oscuro else COLOR_CLARO_BTN
    color_btn_txt = COLOR_OSCURO_BTN_TXT if modo_oscuro else COLOR_CLARO_BTN_TXT
    color_txt = "#fff" if modo_oscuro else "#222"

    obs_win = tk.Toplevel(root)
    obs_win.title("Agregar Observación")
    obs_win.geometry("400x180")
    obs_win.resizable(False, False)
    obs_win.configure(bg=color_bg)
    centrar_ventana(obs_win)
    obs_win.transient(root)
    obs_win.grab_set()
    obs_win.focus_set()

    tk.Label(
        obs_win,
        text=f"Observación para:\n{profesor} / {bloque}",
        bg=color_bg,
        fg=color_txt,
        font=("Segoe UI", 11, "bold")
    ).pack(pady=(18, 8))

    nueva_obs_box = tk.Text(obs_win, height=3, width=38, font=("Segoe UI", 10), bg="#fff", fg="#222", relief="solid", bd=1)
    nueva_obs_box.pack(padx=20, pady=(0, 16))

    def guardar():
        nueva_obs = nueva_obs_box.get("1.0", "end").strip()
        if nueva_obs == "":
            messagebox.showwarning("Observación vacía", "Por favor, ingrese una observación.")
            return
        guardar_observacion(fecha, bloque, profesor, nueva_obs)
        messagebox.showinfo("Observación guardada", "La observación fue registrada correctamente.")
        obs_win.destroy()
        ver_historial()

    btn_frame = tk.Frame(obs_win, bg=color_bg)
    btn_frame.pack(pady=(0, 8))
    btn_guardar = tk.Button(
        btn_frame, text="Guardar", font=FUENTE_BTN, width=13, command=guardar,
        bg=color_btn, fg=color_btn_txt, relief="flat", bd=2
    )
    btn_guardar.pack(side="left", padx=8)
    btn_cancelar = tk.Button(
        btn_frame, text="Cancelar", font=FUENTE_BTN, width=11, command=obs_win.destroy,
        bg="#e5e7e9", fg="#222", relief="flat", bd=2
    )
    btn_cancelar.pack(side="left", padx=8)

    obs_win.transient(root)
    obs_win.grab_set()
    obs_win.focus_set()

# ----- BLOQUES (Y ÁREA RESERVAS DEL DÍA) -----
ALTURA_BLOQUE = 52

def actualizar_bloques():
    reservas = cargar_reservas()
    fecha = datetime.now().strftime("%Y-%m-%d")
    reservas_text.config(state="normal")
    reservas_text.delete(1.0, tk.END)
    color_btn = COLOR_OSCURO_BTN if modo_oscuro else COLOR_CLARO_BTN
    color_btn_hover = COLOR_OSCURO_BTN_HOVER if modo_oscuro else COLOR_CLARO_BTN_HOVER
    color_btn_txt = COLOR_OSCURO_BTN_TXT if modo_oscuro else COLOR_CLARO_BTN_TXT

    for i, bloque in enumerate(bloques):
        _, bloque_frame, _ = bloque_widgets[i]

        # 1. DESTRUIR TODOS los widgets hijos del bloque_frame (quedará vacío)
        for w in bloque_frame.winfo_children():
            w.destroy()

        # 2. Determinar si el bloque está reservado
        reservado, profesor, curso, asignatura = esta_reservado(reservas, fecha, bloque)
        if reservado:
            texto = f"❌ {bloque}  |  {profesor.upper()}  -  {curso} ({asignatura})"
            bg_frame = COLOR_OSCURO_RESERVADO if modo_oscuro else COLOR_CLARO_RESERVADO
            fg_label = "#fff" if modo_oscuro else "#c0392b"
            estado = "disabled"
        else:
            texto = f"✔️ {bloque}  |  Disponible"
            bg_frame = COLOR_OSCURO_DISPONIBLE if modo_oscuro else COLOR_CLARO_DISPONIBLE
            fg_label = "#222" if not modo_oscuro else "#fff"
            estado = "normal"

        bloque_frame.config(bg=bg_frame, height=ALTURA_BLOQUE)

        # 3. Crear NUEVO label cada vez, siempre mismo width y fuente
        btn_label = tk.Label(
            bloque_frame,
            text=texto,
            font=("Segoe UI", 10, "bold"),
            bg=bg_frame,
            fg=fg_label,
            width=60,
            anchor="w",
            justify="left",
            state=estado
        )
        btn_label.pack(side="left", padx=(18,2), pady=10, fill="y")

        # 4. Botón o espacio invisible a la derecha
        if reservado:
            spacer = tk.Frame(bloque_frame, width=120, height=32, bg=bg_frame)
            spacer.pack(side="right", padx=12, pady=8)
        else:
            reservar_btn = tk.Button(
                bloque_frame,
                text="Reservar",
                font=FUENTE_BTN,
                width=12,
                bg=color_btn,
                fg=color_btn_txt,
                relief="flat",
                command=lambda h=bloque: mostrar_formulario(h)
            )
            reservar_btn.pack(side="right", padx=12, pady=8)
            add_hover_effect(reservar_btn, color_btn, color_btn_hover)

    # Área de "Reservas del día"
    reservas_del_dia = []
    for r in reservas:
        if r[0] == fecha:
            if len(r) >= 7:
                reservas_del_dia.append(f"{r[1]}  |  {r[2]} - {r[3]} ({r[4]})  |  Objetivo: {r[5]}")
            else:
                reservas_del_dia.append(f"{r[1]}  |  {r[2]} - {r[3]} ({r[4]})")
    if reservas_del_dia:
        reservas_text.insert(tk.END, "\n".join(reservas_del_dia))
    else:
        reservas_text.insert(tk.END, "No hay reservas registradas para hoy.")

def liberar_reservas():
    fecha = datetime.now().strftime("%d/%m/%Y")
    if messagebox.askyesno(
        "Confirmar",
        f"Esto eliminará **todas** las reservas del día {fecha}.\n"
        "Esta acción no se puede deshacer.\n\n¿Estás seguro que deseas continuar?"
    ):
        fecha_guardado = datetime.now().strftime("%Y-%m-%d")
        reservas = cargar_reservas()
        nuevas_reservas = [r for r in reservas if r[0] != fecha_guardado]
        with open(reservas_file, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(nuevas_reservas)
        messagebox.showinfo("Reservas liberadas", f"Las reservas del día {fecha} han sido eliminadas.")
        actualizar_bloques()

# ----- LOGIN/ENCARGADO -----
def login():
    global logueado_encargado
    color_bg = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    color_btn = COLOR_OSCURO_BTN if modo_oscuro else COLOR_CLARO_BTN
    color_btn_hover = COLOR_OSCURO_BTN_HOVER if modo_oscuro else COLOR_CLARO_BTN_HOVER
    color_btn_txt = COLOR_OSCURO_BTN_TXT if modo_oscuro else COLOR_CLARO_BTN_TXT
    color_label = "#fff" if modo_oscuro else "#222"

    login_window = tk.Toplevel(root)
    login_window.title("Iniciar Sesión (Encargado)")
    login_window.geometry("320x210")
    login_window.configure(bg=color_bg)
    centrar_ventana(login_window)
    login_window.grab_set()
    login_window.resizable(False, False)
    login_window.transient(root)
    login_window.grab_set()
    login_window.focus_set()

    tk.Label(login_window, text="Usuario:", font=FUENTE, bg=color_bg, fg=color_label).pack(pady=(18, 4))
    user_entry = tk.Entry(login_window, font=FUENTE)
    user_entry.pack(pady=0)
    tk.Label(login_window, text="Contraseña:", font=FUENTE, bg=color_bg, fg=color_label).pack(pady=(14, 4))
    pass_entry = tk.Entry(login_window, font=FUENTE, show="*")
    pass_entry.pack(pady=0)

    def verificar():
        usuario = user_entry.get()
        contraseña = pass_entry.get()
        if usuario == "kevin" and contraseña == "2002":
            messagebox.showinfo("Acceso concedido", "Sesión iniciada como encargado.")
            login_window.destroy()
            set_encargado(True)
        else:
            messagebox.showerror("Acceso denegado", "Usuario o contraseña incorrectos.")

    iniciar_btn = tk.Button(login_window, text="Iniciar sesión", font=FUENTE_BTN, bg=color_btn, fg=color_btn_txt,
        width=13, command=verificar)
    iniciar_btn.pack(pady=(22, 6))
    add_hover_effect(iniciar_btn, color_btn, color_btn_hover)

def cerrar_sesion():
    set_encargado(False)
    messagebox.showinfo("Sesión cerrada", "Has cerrado sesión correctamente.")

def set_encargado(val):
    global logueado_encargado
    logueado_encargado = val
    btn_liberar.config(state="normal" if val else "disabled")
    if val:
        btn_configuracion.pack(side="top", fill="x", padx=12, pady=4)
        btn_login.pack_forget()
        btn_logout.pack(fill="both", expand=True)
    else:
        btn_configuracion.pack_forget()
        btn_logout.pack_forget()
        btn_login.pack(fill="both", expand=True)

# ----- TEMA OSCURO -----
def toggle_tema(val=None):
    global modo_oscuro
    modo_oscuro = not modo_oscuro if val is None else val
    toggle_switch.set(modo_oscuro)
    aplicar_tema()
    actualizar_bloques()

def aplicar_tema():
    color_bg = COLOR_OSCURO_BG if modo_oscuro else COLOR_CLARO_BG
    color_panel = COLOR_OSCURO_PANEL if modo_oscuro else COLOR_CLARO_PANEL
    color_side = COLOR_OSCURO_SIDE if modo_oscuro else COLOR_CLARO_SIDE
    color_btn = COLOR_OSCURO_BTN if modo_oscuro else COLOR_CLARO_BTN
    color_btn_hover = COLOR_OSCURO_BTN_HOVER if modo_oscuro else COLOR_CLARO_BTN_HOVER
    color_txt = COLOR_OSCURO_BTN_TXT if modo_oscuro else "#222"

    root.configure(bg=color_bg)
    header.configure(bg=color_panel)
    lbl_title.configure(bg=color_panel, fg="#222" if not modo_oscuro else "#fff")
    sidebar.configure(bg=color_side)
    for child in sidebar.winfo_children():
        try:
            child.configure(bg=color_side, fg="#222" if not modo_oscuro else "#fff")
        except:
            pass
    btn_actualizar.configure(bg=color_btn, fg=color_txt)
    btn_liberar.configure(bg=color_btn, fg=color_txt)
    btn_historial.configure(bg=color_btn, fg=color_txt)
    btn_configuracion.configure(bg=color_btn, fg=color_txt)
    btn_ayuda.configure(bg=color_panel, fg="#222" if not modo_oscuro else "#fff")
    panel_central.configure(bg=color_panel)
    lbl_bloques_title.configure(bg=color_panel, fg="#222" if not modo_oscuro else "#fff")
    reservas_text.configure(bg="#f3f8ff" if not modo_oscuro else "#323b44", fg="#222" if not modo_oscuro else "#fff")
    fecha_label.configure(bg=color_side, fg="#222" if not modo_oscuro else "#fff")
    footer.configure(bg=color_side, fg="#888" if not modo_oscuro else "#aaa")
    panel_bloques.configure(bg=color_panel)
    toggle_switch.configure(bg=color_panel)
    lbl_switch.configure(bg=color_panel, fg="#fff" if modo_oscuro else "#222")
    toggle_switch["bg"] = color_panel
    toggle_switch.draw()
    footer.configure(bg=color_side, fg="#888" if not modo_oscuro else "#aaa")
    if modo_oscuro:
        linea_vertical.configure(bg="#444")
    else:
        linea_vertical.configure(bg="#bbb")
    if os.path.exists(logo_path):
        lbl_logo.configure(bg=color_panel)
    else:
        lbl_logo.configure(bg=color_panel)
    add_hover_effect(btn_actualizar, color_btn, color_btn_hover)
    add_hover_effect(btn_liberar, color_btn, color_btn_hover)
    add_hover_effect(btn_historial, color_btn, color_btn_hover)
    add_hover_effect(btn_configuracion, color_btn, color_btn_hover)
    add_hover_effect(btn_login, color_btn, color_btn_hover)

# =======================================
#         INTERFAZ PRINCIPAL
# =======================================
root = tk.Tk()
root.title("Sistema de reservas sala enlaces")
root.geometry("1050x300")
root.minsize(900, 540)
root.resizable(False, False)
backup_diario()
centrar_ventana(root)

# ----- HEADER -----
header = tk.Frame(root, height=62, bg=COLOR_CLARO_PANEL, bd=1, relief="solid")
header.pack(fill="x", side="top")
header.pack_propagate(0)

if os.path.exists(logo_path):
    from PIL import Image, ImageTk
    logo_img = Image.open(logo_path).resize((38, 38))
    logo = ImageTk.PhotoImage(logo_img)
    lbl_logo = tk.Label(header, image=logo, bg=COLOR_CLARO_PANEL)
    lbl_logo.image = logo
    lbl_logo.pack(side="left", padx=(12,8), pady=8)
else:
    lbl_logo = tk.Label(header, text="🏫", bg=COLOR_CLARO_PANEL, font=("Arial",18))
    lbl_logo.pack(side="left", padx=(12,8), pady=8)

lbl_title = tk.Label(header, text="Sistema de reservas sala enlaces", font=FUENTE_TITULO, bg=COLOR_CLARO_PANEL)
lbl_title.pack(side="left", padx=(6,0))

# ===== FRAME FIJO PARA BOTÓN SESIÓN =====
frame_sesion = tk.Frame(header, bg=COLOR_CLARO_PANEL, width=128, height=38)
frame_sesion.pack_propagate(0)
frame_sesion.pack(side="right", padx=(0, 18), pady=9)

btn_login = tk.Button(frame_sesion, text="Iniciar Sesión", font=FUENTE_BTN, command=login,
                      bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, width=12)
btn_login.pack(fill="both", expand=True)
btn_logout = tk.Button(frame_sesion, text="Cerrar Sesión", font=FUENTE_BTN, command=cerrar_sesion,
                       bg="#f5b7b1", fg="#222", width=12)
# Solo uno visible a la vez, ver función set_encargado

toggle_switch = ToggleSwitch(header, command=lambda state: toggle_tema(state), bg=COLOR_CLARO_PANEL, state=modo_oscuro)
toggle_switch.pack(side="right", padx=(0, 18), pady=11)
lbl_switch = tk.Label(header, text="Tema Oscuro", font=("Segoe UI",10,"bold"), bg=COLOR_CLARO_PANEL)
lbl_switch.pack(side="right", padx=(0, 10), pady=17)

# ----- SIDEBAR -----
sidebar = tk.Frame(root, width=220, bg=COLOR_CLARO_SIDE)
sidebar.pack(side="left", fill="y")
sidebar.pack_propagate(0)

btn_actualizar = tk.Button(sidebar, text="Actualizar Reservas", font=FUENTE_SIDE, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, relief="flat", command=actualizar_bloques)
btn_actualizar.pack(side="top", fill="x", padx=12, pady=(22,4))
btn_liberar = tk.Button(sidebar, text="Liberar Reservas del día", font=FUENTE_SIDE, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, relief="flat", command=liberar_reservas, state="disabled")
btn_liberar.pack(side="top", fill="x", padx=12, pady=4)
btn_historial = tk.Button(sidebar, text="Ver Historial", font=FUENTE_SIDE, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, relief="flat", command=ver_historial)
btn_historial.pack(side="top", fill="x", padx=12, pady=4)
btn_configuracion = tk.Button(sidebar, text="Configuración", font=FUENTE_SIDE, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, relief="flat", command=lambda: messagebox.showinfo("Configuración", "Solo encargado."))
btn_configuracion.pack_forget()
btn_ayuda = tk.Button(sidebar, text="❓ Ayuda", font=FUENTE_BTN, bg=COLOR_CLARO_PANEL, relief="flat", command=lambda: messagebox.showinfo("Ayuda", "Contacto: kevinalbanez.r.grecia@educacionlascabras.cl"))
btn_ayuda.pack(side="top", fill="x", padx=12, pady=(26,2))

linea_vertical = tk.Frame(root, width=1, bg="#bbb")
linea_vertical.pack(side="left", fill="y")

fecha_label = tk.Label(sidebar, text="", font=("Segoe UI", 10, "italic"), bg=COLOR_CLARO_SIDE, fg="#222")
fecha_label.pack(side="bottom", pady=(0,8))
actualizar_fecha_hora()

footer = tk.Label(
    sidebar,
    text="Copyright 2025-2025 Grecia. All rights reserved",
    font=("Segoe UI", 7, "italic"),
    bg=COLOR_CLARO_SIDE,
    fg="#888",
    anchor="center",
    justify="center"
)
footer.pack(side="bottom", pady=(2, 4), fill="x")

# ----- PANEL CENTRAL (BLOQUES + RESERVAS DÍA) -----
panel_central = tk.Frame(root, bg=COLOR_CLARO_PANEL)
panel_central.pack(side="left", fill="both", expand=True)
panel_central.pack_propagate(0)

lbl_bloques_title = tk.Label(panel_central, text="Bloques disponibles hoy", font=FUENTE_PANEL, bg=COLOR_CLARO_PANEL)
lbl_bloques_title.pack(anchor="n", pady=(22,12))

panel_bloques = tk.Frame(panel_central, bg=COLOR_CLARO_PANEL)
panel_bloques.pack(padx=18, pady=(6, 6), fill="x")
bloque_widgets = []
for i, b in enumerate(bloques):
    bloque_frame = tk.Frame(panel_bloques, bg=COLOR_CLARO_DISPONIBLE, bd=0, relief="ridge", height=ALTURA_BLOQUE)
    bloque_frame.pack(side="top", fill="x", padx=4, pady=8)
    bloque_frame.pack_propagate(False)
    btn = tk.Label(bloque_frame, text=f"{b}", font=("Segoe UI", 12, "bold"), bg=COLOR_CLARO_DISPONIBLE,
               width=60, anchor="w", justify="left")
    btn.pack(side="left", padx=(18,2), pady=10, fill="y")
    reservar_btn = None
    def reservar_cierre(h=b):
        mostrar_formulario(h)
    reservar_btn = tk.Button(bloque_frame, text="Reservar", font=FUENTE_BTN, width=12, bg=COLOR_CLARO_BTN, fg=COLOR_CLARO_BTN_TXT, relief="flat", command=reservar_cierre)
    add_hover_effect(reservar_btn, COLOR_CLARO_BTN, COLOR_CLARO_BTN_HOVER)
    reservar_btn.pack(side="right", padx=12, pady=8)
    bloque_widgets.append((btn, bloque_frame, reservar_btn))

reservas_text = scrolledtext.ScrolledText(
    panel_central, height=5, width=74, state="disabled", font=("Segoe UI", 10),
    bg="#f4f8fd", relief="flat", borderwidth=1
)
reservas_text.pack(fill="x", expand=False, padx=20, pady=(20,0), anchor="s")

# ====== INICIALIZACIÓN FINAL ======
aplicar_tema()
actualizar_bloques()
root.mainloop()