# app/main_modern.py
"""
Sistema de Calificaciones - Club de Matemáticas Lobatchewsky
Interfaz moderna y minimalista
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import threading

# Importar módulos del sistema
from config import (RUTA_DATOS, MAPEO_MATERIAS, COLUMNA_NOMBRE,
                    COLUMNA_EMAIL, COLUMNA_GRUPO)
from data_loader import cargar_datos, validar_estructura_csv
from grader import procesar_calificaciones_google_forms, calcular_estadisticas_grupo
from generador_clave import GeneradorClave
from excel_consolidado import generar_reporte_consolidado

# Colores del tema Lobatchewsky
COLORS = {
    'principal': '#005DAB',  # Azul corporativo
    'principal_hover': '#004080',  # Azul más oscuro para hover
    'blanco': '#FFFFFF',
    'gris_claro': '#F2F2F2',
    'gris_medio': '#CCCCCC',
    'gris_oscuro': '#666666',
    'texto_principal': '#333333',
    'exito': '#28A745',
    'advertencia': '#FFC107',
    'error': '#DC3545'
}


class VentanaGeneradorClaveModerna(tk.Toplevel):
    """Ventana moderna para generar claves de respuestas."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Crear Clave de Respuestas")
        self.geometry("800x800")  # Aumentado el tamaño inicial
        self.configure(bg=COLORS['gris_claro'])
        self.resizable(True, True)  # Ahora es redimensionable
        self.minsize(750, 700)  # Tamaño mínimo

        self.generador = None
        self.total_preguntas = tk.IntVar(value=110)

        self.crear_interfaz()

        # Modal
        self.transient(parent)
        self.grab_set()

        # Centrar ventana DESPUÉS de crear interfaz
        self.update_idletasks()
        self.center_window()

        # Forzar actualización para asegurar visibilidad
        self.after(100, self.ajustar_tamanio)

    def ajustar_tamanio(self):
        """Ajusta el tamaño de la ventana para mostrar todo el contenido."""
        self.update_idletasks()

        # Obtener el tamaño requerido del contenido
        required_height = self.winfo_reqheight()
        required_width = self.winfo_reqwidth()

        # Asegurar que la ventana sea lo suficientemente grande
        current_width = self.winfo_width()
        current_height = self.winfo_height()

        # Si el contenido es más grande que la ventana actual, ajustar
        new_width = max(current_width, required_width + 40, 800)
        new_height = max(current_height, required_height + 40, 800)

        # Limitar al 90% del tamaño de la pantalla
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        new_width = min(new_width, int(screen_width * 0.9))
        new_height = min(new_height, int(screen_height * 0.9))

        self.geometry(f'{new_width}x{new_height}')
        self.center_window()

    def center_window(self):
        """Centra la ventana en la pantalla."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def crear_interfaz(self):
        # Canvas con scrollbar para todo el contenido
        canvas = tk.Canvas(self, bg=COLORS['gris_claro'], highlightthickness=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLORS['gris_claro'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Empaquetar canvas y scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Habilitar scroll con la rueda del mouse
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Frame principal con padding DENTRO del canvas
        main_frame = tk.Frame(scrollable_frame, bg=COLORS['gris_claro'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Título
        titulo_frame = tk.Frame(main_frame, bg=COLORS['principal'],
                                height=60, bd=0)
        titulo_frame.pack(fill=tk.X, pady=(0, 20))
        titulo_frame.pack_propagate(False)

        tk.Label(titulo_frame, text="✏️ Crear Clave de Respuestas",
                 font=('Poppins', 16, 'bold'),
                 bg=COLORS['principal'],
                 fg=COLORS['blanco']).pack(expand=True)

        # Frame de configuración
        config_frame = self.crear_frame_seccion(main_frame, "Configuración")

        inner_frame = tk.Frame(config_frame, bg=COLORS['blanco'])
        inner_frame.pack(fill=tk.X, padx=15, pady=10)

        tk.Label(inner_frame, text="Total de preguntas:",
                 font=('Open Sans', 10),
                 bg=COLORS['blanco']).pack(side=tk.LEFT, padx=5)

        tk.Spinbox(inner_frame, from_=1, to=200,
                   textvariable=self.total_preguntas,
                   width=10, font=('Open Sans', 10)).pack(side=tk.LEFT, padx=5)

        # Botón Inicializar
        btn_init = tk.Button(
            inner_frame,
            text="Inicializar",
            command=self.inicializar_generador,
            font=('Open Sans', 9, 'bold'),
            bg=COLORS['principal'],
            fg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        btn_init.pack(side=tk.LEFT, padx=15)

        # Frame de entrada
        entrada_frame = self.crear_frame_seccion(main_frame,
                                                 "Ingresar Respuestas")

        tk.Label(entrada_frame,
                 text="Pegue las respuestas (una por línea, formato flexible):",
                 font=('Open Sans', 9),
                 bg=COLORS['blanco'],
                 fg=COLORS['gris_oscuro']).pack(anchor=tk.W, padx=15, pady=(10, 5))

        self.texto_respuestas = scrolledtext.ScrolledText(
            entrada_frame, height=12,  # Reducido para dar espacio a los botones
            font=('Courier New', 10),
            wrap=tk.WORD,
            bg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS['gris_medio'])
        self.texto_respuestas.pack(fill=tk.BOTH, expand=False,
                                   padx=15, pady=(0, 10))

        # Botones de acción
        btn_frame = tk.Frame(entrada_frame, bg=COLORS['blanco'])
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        # Botón Procesar
        btn_procesar = tk.Button(
            btn_frame,
            text="📝 Procesar",
            command=self.procesar_respuestas,
            font=('Open Sans', 9, 'bold'),
            bg=COLORS['principal'],
            fg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        btn_procesar.pack(side=tk.LEFT, padx=5)

        # Botón Limpiar
        btn_limpiar = tk.Button(
            btn_frame,
            text="🗑️ Limpiar",
            command=self.limpiar,
            font=('Open Sans', 9, 'bold'),
            bg=COLORS['gris_medio'],
            fg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        btn_limpiar.pack(side=tk.LEFT, padx=5)

        # Botón Vista Previa
        btn_preview = tk.Button(
            btn_frame,
            text="👁️ Vista Previa",
            command=self.mostrar_vista_previa,
            font=('Open Sans', 9, 'bold'),
            bg=COLORS['principal'],
            fg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        btn_preview.pack(side=tk.LEFT, padx=5)

        # Frame de estado
        estado_frame = self.crear_frame_seccion(main_frame, "Estado")

        self.label_info = tk.Label(estado_frame,
                                   text="No inicializado",
                                   font=('Open Sans', 10),
                                   bg=COLORS['blanco'],
                                   fg=COLORS['gris_oscuro'],
                                   anchor=tk.W)
        self.label_info.pack(fill=tk.X, padx=15, pady=10)

        # ========== BOTONES FINALES - AHORA CON FONDO DESTACADO ==========
        final_container = tk.Frame(main_frame, bg=COLORS['gris_medio'],
                                   relief=tk.RAISED, bd=2)
        final_container.pack(fill=tk.X, pady=(20, 10))

        final_frame = tk.Frame(final_container, bg=COLORS['gris_medio'])
        final_frame.pack(fill=tk.X, padx=10, pady=10)

        # Botón Generar CSV - MÁS GRANDE Y DESTACADO
        self.btn_generar = tk.Button(
            final_frame,
            text="💾 Generar y Guardar CSV",
            command=self.generar_csv,
            font=('Open Sans', 11, 'bold'),
            bg=COLORS['exito'],
            fg=COLORS['blanco'],
            relief=tk.RAISED,
            bd=2,
            padx=25,
            pady=12,
            cursor='hand2'
        )
        self.btn_generar.pack(side=tk.LEFT, padx=10)

        # Hover para Generar
        def on_enter_generar(e):
            self.btn_generar['bg'] = self.oscurecer_color(COLORS['exito'])
            self.btn_generar['relief'] = tk.SUNKEN

        def on_leave_generar(e):
            self.btn_generar['bg'] = COLORS['exito']
            self.btn_generar['relief'] = tk.RAISED

        self.btn_generar.bind("<Enter>", on_enter_generar)
        self.btn_generar.bind("<Leave>", on_leave_generar)

        # Botón Cancelar
        self.btn_cancelar = tk.Button(
            final_frame,
            text="✖ Cancelar",
            command=self.destroy,
            font=('Open Sans', 10, 'bold'),
            bg=COLORS['error'],
            fg=COLORS['blanco'],
            relief=tk.RAISED,
            bd=2,
            padx=20,
            pady=12,
            cursor='hand2'
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=10)

        # Hover para Cancelar
        def on_enter_cancelar(e):
            self.btn_cancelar['bg'] = self.oscurecer_color(COLORS['error'])
            self.btn_cancelar['relief'] = tk.SUNKEN

        def on_leave_cancelar(e):
            self.btn_cancelar['bg'] = COLORS['error']
            self.btn_cancelar['relief'] = tk.RAISED

        self.btn_cancelar.bind("<Enter>", on_enter_cancelar)
        self.btn_cancelar.bind("<Leave>", on_leave_cancelar)

        # Nota informativa
        nota_label = tk.Label(
            final_frame,
            text="← Genere el archivo CSV cuando esté listo",
            font=('Open Sans', 9, 'italic'),
            bg=COLORS['gris_medio'],
            fg=COLORS['gris_oscuro']
        )
        nota_label.pack(side=tk.LEFT, padx=20)

    def crear_frame_seccion(self, parent, titulo):
        """Crea un frame de sección con título."""
        frame = tk.Frame(parent, bg=COLORS['blanco'],
                         relief=tk.FLAT, bd=0)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Título de sección
        titulo_label = tk.Label(frame, text=titulo,
                                font=('Poppins', 11, 'bold'),
                                bg=COLORS['blanco'],
                                fg=COLORS['principal'],
                                anchor=tk.W)
        titulo_label.pack(fill=tk.X, padx=15, pady=(10, 5))

        # Línea separadora
        separator = tk.Frame(frame, height=2,
                             bg=COLORS['principal'])
        separator.pack(fill=tk.X, padx=15, pady=(0, 5))

        return frame

    def oscurecer_color(self, color):
        """Oscurece un color hex."""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
        rgb = tuple(max(0, int(c * 0.8)) for c in rgb)
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'

    def inicializar_generador(self):
        total = self.total_preguntas.get()
        self.generador = GeneradorClave(total_preguntas=total)
        self.actualizar_info()
        messagebox.showinfo("✓ Inicializado",
                            f"Generador listo para {total} preguntas.")

    def procesar_respuestas(self):
        if self.generador is None:
            messagebox.showwarning("⚠ Advertencia",
                                   "Primero inicialice el generador.")
            return

        texto = self.texto_respuestas.get("1.0", tk.END)
        num_procesadas = self.generador.establecer_respuestas_desde_texto(texto)

        self.actualizar_info()
        messagebox.showinfo("✓ Procesado",
                            f"Respuestas procesadas: {num_procesadas}\n"
                            f"Válidas: {self.generador.contar_respuestas_validas()}")

    def limpiar(self):
        self.texto_respuestas.delete("1.0", tk.END)
        if self.generador:
            self.generador.limpiar_respuestas()
            self.actualizar_info()

    def mostrar_vista_previa(self):
        if self.generador is None:
            messagebox.showwarning("⚠ Advertencia",
                                   "Primero inicialice el generador.")
            return

        ventana = tk.Toplevel(self)
        ventana.title("Vista Previa")
        ventana.geometry("700x550")
        ventana.configure(bg=COLORS['gris_claro'])

        texto = scrolledtext.ScrolledText(ventana,
                                          font=('Courier New', 9),
                                          bg=COLORS['blanco'])
        texto.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        reporte = self.generador.generar_reporte_texto()
        texto.insert("1.0", reporte)
        texto.config(state=tk.DISABLED)

        btn_cerrar = tk.Button(
            ventana,
            text="Cerrar",
            command=ventana.destroy,
            font=('Open Sans', 9, 'bold'),
            bg=COLORS['principal'],
            fg=COLORS['blanco'],
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        btn_cerrar.pack(pady=10)

    def actualizar_info(self):
        if self.generador is None:
            self.label_info.config(text="❌ No inicializado")
            return

        total = self.generador.total_preguntas
        validas = self.generador.contar_respuestas_validas()
        faltantes = total - validas
        porcentaje = (validas / total * 100) if total > 0 else 0

        texto = (f"Total: {total} | Válidas: {validas} | "
                 f"Faltantes: {faltantes} | Completado: {porcentaje:.1f}%")
        self.label_info.config(text=texto)

    def generar_csv(self):
        if self.generador is None:
            messagebox.showwarning("⚠ Advertencia",
                                   "Primero inicialice el generador.")
            return

        if self.generador.contar_respuestas_validas() == 0:
            messagebox.showwarning("⚠ Advertencia",
                                   "No hay respuestas válidas.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"clave_respuestas_{self.total_preguntas.get()}.csv"
        )

        if filename:
            if self.generador.generar_csv(filename, formato='n.'):
                messagebox.showinfo("✓ Éxito",
                                    f"Archivo generado:\n{filename}")
                self.destroy()
            else:
                messagebox.showerror("✗ Error",
                                     "No se pudo generar el archivo.")


class SistemaCalificacionesLobatchewsky:
    """Sistema principal de calificaciones con interfaz moderna."""

    def __init__(self, root):
        self.root = root
        self.root.title("Club de Matemáticas Lobatchewsky - Sistema de Calificaciones")
        self.root.geometry("1400x800")
        self.root.configure(bg=COLORS['gris_claro'])

        # Variables
        self.ruta_clave = tk.StringVar()
        self.ruta_respuestas = tk.StringVar()
        self.resultados = None
        self.procesando = False

        # Fuentes
        self.font_titulo = ('Poppins', 14, 'bold')
        self.font_subtitulo = ('Poppins', 11, 'bold')
        self.font_normal = ('Open Sans', 10)
        self.font_boton = ('Open Sans', 10, 'bold')

        self.crear_interfaz()
        self.centrar_ventana()

    def centrar_ventana(self):
        """Centra la ventana en la pantalla."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def crear_interfaz(self):
        # Frame principal
        main_container = tk.Frame(self.root, bg=COLORS['gris_claro'])
        main_container.pack(fill=tk.BOTH, expand=True)

        # Panel lateral izquierdo
        self.crear_panel_lateral(main_container)

        # Área principal derecha
        self.crear_area_principal(main_container)

    def crear_panel_lateral(self, parent):
        """Crea el panel lateral con controles."""
        panel = tk.Frame(parent, bg=COLORS['principal'], width=350)
        panel.pack(side=tk.LEFT, fill=tk.Y)
        panel.pack_propagate(False)

        # Logo y título
        header_frame = tk.Frame(panel, bg=COLORS['principal'], height=80)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        header_frame.pack_propagate(False)

        tk.Label(header_frame,
                 text="🎓",
                 font=('Arial', 32),
                 bg=COLORS['principal'],
                 fg=COLORS['blanco']).pack(pady=(10, 0))

        tk.Label(header_frame,
                 text="LOBATCHEWSKY",
                 font=('Poppins', 12, 'bold'),
                 bg=COLORS['principal'],
                 fg=COLORS['blanco']).pack()

        # Sección: Cargar Archivos
        self.crear_seccion_lateral(panel, "📁 CARGAR ARCHIVOS")

        # Botón para cargar clave de respuestas
        self.crear_boton_lateral(panel, "📂 Cargar Clave de Respuestas",
                                 lambda: self.buscar_archivo('clave'),
                                 padx=20, pady=(10, 5))

        # Mostrar archivo cargado (clave)
        self.label_clave = tk.Label(panel,
                                    text="Sin archivo",
                                    font=('Open Sans', 8),
                                    bg=COLORS['principal'],
                                    fg=COLORS['gris_claro'],
                                    anchor=tk.W,
                                    wraplength=310)
        self.label_clave.pack(fill=tk.X, padx=20, pady=(0, 5))

        self.crear_boton_lateral(panel, "➕ Crear Nueva Clave",
                                 self.abrir_generador_clave,
                                 color=COLORS['exito'],
                                 padx=20, pady=(0, 15))

        # Botón para cargar respuestas de alumnos
        self.crear_boton_lateral(panel, "📂 Cargar Respuestas de Alumnos",
                                 lambda: self.buscar_archivo('respuestas'),
                                 padx=20, pady=(5, 5))

        # Mostrar archivo cargado (respuestas)
        self.label_respuestas = tk.Label(panel,
                                         text="Sin archivo",
                                         font=('Open Sans', 8),
                                         bg=COLORS['principal'],
                                         fg=COLORS['gris_claro'],
                                         anchor=tk.W,
                                         wraplength=310)
        self.label_respuestas.pack(fill=tk.X, padx=20, pady=(0, 15))

        # Sección: Procesamiento
        self.crear_seccion_lateral(panel, "⚙️ PROCESAMIENTO")

        self.crear_boton_lateral(panel, "🚀 CALIFICAR",
                                 self.procesar,
                                 color=COLORS['exito'],
                                 padx=20, pady=(10, 15),
                                 height=2)

        # Sección: Reportes
        self.crear_seccion_lateral(panel, "📊 REPORTES")

        self.btn_mostrar = self.crear_boton_lateral(panel, "👁️ Mostrar Reporte",
                                                    self.mostrar_resultados,
                                                    padx=20, pady=(10, 5),
                                                    state=tk.DISABLED)

        self.btn_excel = self.crear_boton_lateral(panel, "📄 Generar Excel",
                                                  self.generar_excel_consolidado,
                                                  padx=20, pady=(0, 15),
                                                  state=tk.DISABLED,
                                                  color=COLORS['exito'])

        # Estado
        self.status_label = tk.Label(panel,
                                     text="Listo",
                                     font=('Open Sans', 9),
                                     bg=COLORS['principal'],
                                     fg=COLORS['blanco'],
                                     anchor=tk.W,
                                     wraplength=310)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X,
                               padx=20, pady=20)

    def crear_seccion_lateral(self, parent, titulo):
        """Crea un título de sección en el panel lateral."""
        frame = tk.Frame(parent, bg=COLORS['principal'])
        frame.pack(fill=tk.X, padx=20, pady=(20, 5))

        tk.Label(frame, text=titulo,
                 font=('Poppins', 10, 'bold'),
                 bg=COLORS['principal'],
                 fg=COLORS['blanco'],
                 anchor=tk.W).pack(side=tk.LEFT)

        # Línea
        separator = tk.Frame(frame, height=2, bg=COLORS['blanco'])
        separator.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

    def crear_boton_lateral(self, parent, texto, comando,
                            color=None, width=None, height=1,
                            padx=0, pady=0, state=tk.NORMAL):
        """Crea un botón en el panel lateral."""
        bg_color = color or COLORS['principal_hover']

        btn = tk.Button(parent, text=texto,
                        command=comando,
                        font=self.font_boton,
                        bg=bg_color,
                        fg=COLORS['blanco'],
                        relief=tk.FLAT,
                        bd=0,
                        cursor='hand2',
                        state=state,
                        height=height)

        if width:
            btn.config(width=width)

        btn.pack(fill=tk.X, padx=padx, pady=pady)

        # Efecto hover
        def on_enter(e):
            if btn['state'] == tk.NORMAL:
                btn['bg'] = self.oscurecer_color(bg_color)

        def on_leave(e):
            btn['bg'] = bg_color

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        return btn

    def crear_area_principal(self, parent):
        """Crea el área principal derecha."""
        area = tk.Frame(parent, bg=COLORS['gris_claro'])
        area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                  padx=20, pady=20)

        # Título
        titulo_frame = tk.Frame(area, bg=COLORS['blanco'],
                                relief=tk.FLAT, bd=0)
        titulo_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(titulo_frame,
                 text="📊 Resultados y Métricas",
                 font=('Poppins', 16, 'bold'),
                 bg=COLORS['blanco'],
                 fg=COLORS['principal'],
                 anchor=tk.W).pack(side=tk.LEFT, padx=20, pady=15)

        # Área de contenido
        contenido_frame = tk.Frame(area, bg=COLORS['blanco'],
                                   relief=tk.FLAT, bd=0)
        contenido_frame.pack(fill=tk.BOTH, expand=True)

        # Mensaje inicial
        self.frame_inicial = tk.Frame(contenido_frame, bg=COLORS['blanco'])
        self.frame_inicial.pack(fill=tk.BOTH, expand=True)

        tk.Label(self.frame_inicial,
                 text="👈 Comience cargando los archivos",
                 font=('Poppins', 14),
                 bg=COLORS['blanco'],
                 fg=COLORS['gris_medio']).pack(expand=True)

        # Frame de resultados (oculto inicialmente)
        self.frame_resultados = tk.Frame(contenido_frame, bg=COLORS['blanco'])

        # Área de texto con scroll
        self.texto_resultados = scrolledtext.ScrolledText(
            self.frame_resultados,
            font=('Courier New', 9),
            bg=COLORS['blanco'],
            fg=COLORS['texto_principal'],
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0)
        self.texto_resultados.pack(fill=tk.BOTH, expand=True,
                                   padx=20, pady=20)

    def oscurecer_color(self, color):
        """Oscurece un color hex."""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
        rgb = tuple(max(0, int(c * 0.8)) for c in rgb)
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'

    def abrir_generador_clave(self):
        """Abre la ventana del generador de claves."""
        VentanaGeneradorClaveModerna(self.root)

    def buscar_archivo(self, tipo):
        """Busca un archivo CSV."""
        filename = filedialog.askopenfilename(
            title=f'Seleccionar archivo de {tipo}',
            filetypes=[('CSV', '*.csv'), ('Todos', '*.*')],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )

        if filename:
            nombre_archivo = os.path.basename(filename)

            if tipo == 'clave':
                self.ruta_clave.set(filename)
                self.label_clave.config(
                    text=f"✓ {nombre_archivo}",
                    fg=COLORS['blanco'])
            else:
                self.ruta_respuestas.set(filename)
                self.label_respuestas.config(
                    text=f"✓ {nombre_archivo}",
                    fg=COLORS['blanco'])

            self.status_label.config(
                text=f"✓ Archivo cargado: {nombre_archivo}")

    def procesar(self):
        """Procesa las calificaciones."""
        if self.procesando:
            return

        if not self.ruta_clave.get() or not self.ruta_respuestas.get():
            messagebox.showerror("Error",
                                 "Por favor seleccione ambos archivos")
            return

        self.procesando = True
        self.status_label.config(text="⏳ Procesando...")
        self.root.update()

        # Ejecutar en hilo separado
        thread = threading.Thread(target=self._procesar_thread)
        thread.daemon = True
        thread.start()

    def _procesar_thread(self):
        """Procesamiento en hilo separado."""
        try:
            clave_df = cargar_datos(self.ruta_clave.get())
            respuestas_df = cargar_datos(self.ruta_respuestas.get())

            if clave_df is None or respuestas_df is None:
                self.root.after(0, self._procesar_error,
                                "Error al cargar archivos")
                return

            valido_clave, _ = validar_estructura_csv(clave_df, es_clave=True)
            valido_resp, _ = validar_estructura_csv(respuestas_df, es_clave=False)

            if not valido_clave or not valido_resp:
                self.root.after(0, self._procesar_error,
                                "Formato de archivo incorrecto")
                return

            resultados = procesar_calificaciones_google_forms(
                clave_df, respuestas_df, MAPEO_MATERIAS,
                COLUMNA_NOMBRE, COLUMNA_EMAIL, COLUMNA_GRUPO
            )

            if not resultados:
                self.root.after(0, self._procesar_error,
                                "No se procesaron calificaciones")
                return

            self.root.after(0, self._procesar_completado, resultados)

        except Exception as e:
            self.root.after(0, self._procesar_error, str(e))

    def _procesar_completado(self, resultados):
        """Callback cuando el procesamiento termina."""
        self.resultados = resultados
        self.procesando = False

        self.status_label.config(
            text=f"✅ Procesamiento completado: {len(resultados)} alumnos")

        # Habilitar botones
        self.btn_mostrar.config(state=tk.NORMAL)
        self.btn_excel.config(state=tk.NORMAL)

        messagebox.showinfo("✓ Éxito",
                            f"Se procesaron {len(resultados)} alumnos correctamente")

    def _procesar_error(self, mensaje):
        """Callback cuando hay un error."""
        self.procesando = False
        self.status_label.config(text=f"✗ Error: {mensaje}")
        messagebox.showerror("Error", mensaje)

    def mostrar_resultados(self):
        """Muestra los resultados en el área principal."""
        if not self.resultados:
            return

        # Ocultar mensaje inicial y mostrar resultados
        self.frame_inicial.pack_forget()
        self.frame_resultados.pack(fill=tk.BOTH, expand=True)

        # Limpiar y mostrar
        self.texto_resultados.delete(1.0, tk.END)

        # Generar reporte
        reporte = self._generar_reporte_visual()
        self.texto_resultados.insert(1.0, reporte)

        self.status_label.config(text="✓ Resultados mostrados")

    def _generar_reporte_visual(self):
        """Genera un reporte visual de los resultados."""
        lineas = []
        lineas.append("=" * 90)
        lineas.append("CLUB DE MATEMÁTICAS LOBATCHEWSKY".center(90))
        lineas.append(f"Reporte de Calificaciones - {datetime.now().strftime('%d/%m/%Y %H:%M')}".center(90))
        lineas.append("=" * 90)
        lineas.append("")

        # Estadísticas generales
        estadisticas = calcular_estadisticas_grupo(self.resultados)

        lineas.append("📊 ESTADÍSTICAS GENERALES")
        lineas.append("-" * 90)
        lineas.append(f"  Total de Alumnos:        {estadisticas['total_alumnos']}")
        lineas.append(
            f"  Promedio del Grupo:      {estadisticas['promedio_calificacion']:.2f}/10 ({estadisticas['promedio_global']:.2f}%)")
        lineas.append(
            f"  Calificación Máxima:     {estadisticas['porcentaje_maximo']:.2f}% - {estadisticas['mejor_alumno']}")
        lineas.append(
            f"  Calificación Mínima:     {estadisticas['porcentaje_minimo']:.2f}% - {estadisticas['peor_alumno']}")
        lineas.append("")

        # Distribución
        excelente = sum(1 for r in self.resultados if r['porcentaje_global'] >= 90)
        muy_bien = sum(1 for r in self.resultados if 80 <= r['porcentaje_global'] < 90)
        bien = sum(1 for r in self.resultados if 70 <= r['porcentaje_global'] < 80)
        regular = sum(1 for r in self.resultados if 60 <= r['porcentaje_global'] < 70)
        necesita = sum(1 for r in self.resultados if r['porcentaje_global'] < 60)

        lineas.append("📈 DISTRIBUCIÓN DE CALIFICACIONES")
        lineas.append("-" * 90)
        lineas.append(f"  Excelente (90-100%):      {excelente:3d} {'█' * (excelente * 3)}")
        lineas.append(f"  Muy bien (80-89%):        {muy_bien:3d} {'█' * (muy_bien * 3)}")
        lineas.append(f"  Bien (70-79%):            {bien:3d} {'█' * (bien * 3)}")
        lineas.append(f"  Regular (60-69%):         {regular:3d} {'█' * (regular * 3)}")
        lineas.append(f"  Necesita mejorar (<60%):  {necesita:3d} {'█' * (necesita * 3)}")
        lineas.append("")

        # Promedio por materia
        lineas.append("📚 PROMEDIO POR MATERIA")
        lineas.append("-" * 90)
        for materia, datos in sorted(estadisticas['materias'].items()):
            barra = self._crear_barra_ascii(datos['promedio'], 40)
            lineas.append(f"  {materia:20s} {datos['promedio']:6.2f}% {barra}")
        lineas.append("")

        # Detalles por alumno
        lineas.append("👥 CALIFICACIONES POR ALUMNO")
        lineas.append("-" * 90)
        lineas.append(f"  {'#':<4} {'Nombre':<30} {'Grupo':<10} {'Total':<10} {'Cal.':<8} {'%':<8}")
        lineas.append("-" * 90)

        for idx, resultado in enumerate(sorted(self.resultados,
                                               key=lambda x: x['porcentaje_global'],
                                               reverse=True), 1):
            nombre = resultado['nombre'][:28]
            grupo = str(resultado.get('grupo', ''))[:8]
            total = f"{resultado['total_aciertos']}/{resultado['total_preguntas']}"
            cal = f"{resultado['calificacion_global']:.2f}"
            porc = f"{resultado['porcentaje_global']:.1f}%"

            lineas.append(f"  {idx:<4} {nombre:<30} {grupo:<10} {total:<10} {cal:<8} {porc:<8}")

        lineas.append("=" * 90)

        return "\n".join(lineas)

    def _crear_barra_ascii(self, porcentaje, longitud=40):
        """Crea una barra de progreso ASCII."""
        lleno = int(porcentaje / 100 * longitud)
        return f"[{'█' * lleno}{' ' * (longitud - lleno)}]"

    def generar_excel_consolidado(self):
        """Genera el reporte Excel consolidado."""
        if not self.resultados:
            messagebox.showwarning("Advertencia",
                                   "Primero procese las calificaciones")
            return

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"reporte_lobatchewsky_{timestamp}.xlsx"
        )

        if filename:
            try:
                self.status_label.config(text="⏳ Generando Excel consolidado...")
                self.root.update()

                generar_reporte_consolidado(self.resultados, filename)

                self.status_label.config(text="✅ Excel generado correctamente")
                messagebox.showinfo("✓ Éxito",
                                    f"Reporte Excel consolidado generado:\n\n{filename}\n\n"
                                    "El archivo contiene múltiples hojas:\n"
                                    "• Resumen General\n"
                                    "• Calificaciones Detalladas\n"
                                    "• Hojas por Grupo\n"
                                    "• Análisis de Errores\n"
                                    "• Preguntas Difíciles\n"
                                    "• Matriz Visual")

            except Exception as e:
                self.status_label.config(text=f"✗ Error: {str(e)}")
                messagebox.showerror("Error", f"No se pudo generar el Excel:\n{str(e)}")


def main():
    """Función principal."""
    root = tk.Tk()

    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')

    # Crear aplicación
    app = SistemaCalificacionesLobatchewsky(root)

    # Iniciar
    root.mainloop()


if __name__ == "__main__":
    main()