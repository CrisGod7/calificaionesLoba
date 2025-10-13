# main_moderno.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
# Usamos ttkbootstrap para una interfaz moderna
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from config import (RUTA_DATOS, MAPEO_MATERIAS, COLUMNA_NOMBRE,
                    COLUMNA_EMAIL, COLUMNA_GRUPO)
from data_loader import cargar_datos, validar_estructura_csv
from grader import procesar_calificaciones_google_forms
from reporter import exportar_a_csv, exportar_a_excel
from generador_clave import GeneradorClave


# --- VENTANA DEL GENERADOR DE CLAVE (REDiseñada) ---
class VentanaGeneradorClave(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Generador de Clave de Respuestas")
        # Geometría inicial adaptable
        self.geometry("750x700")
        self.minsize(600, 500)

        self.generador = None
        self.total_preguntas = tk.IntVar(value=110)

        self.crear_interfaz()
        self.transient(parent)
        self.grab_set()

    def crear_interfaz(self):
        # Tema 'cosmo' de ttkbootstrap
        style = ttk.Style(theme='cosmo')

        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título con el color principal de la marca
        ttk.Label(main_frame, text="Generador de Clave de Respuestas",
                  font=('Helvetica', 18, 'bold'), foreground="#0077B6").pack(pady=(0, 20))

        # --- Frame de Configuración ---
        config_frame = ttk.Labelframe(main_frame, text=" 1. Configuración Inicial ", padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(config_frame, text="Total de preguntas:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Spinbox(config_frame, from_=1, to=200, textvariable=self.total_preguntas,
                    width=8).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_frame, text="Inicializar Generador",
                   command=self.inicializar_generador, bootstyle="primary").pack(side=tk.LEFT, padx=20)

        # --- Frame de Entrada ---
        entrada_frame = ttk.Labelframe(main_frame, text=" 2. Ingresar Respuestas ", padding="15")
        entrada_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        entrada_frame.columnconfigure(0, weight=1)
        entrada_frame.rowconfigure(1, weight=1)

        ttk.Label(entrada_frame,
                  text="Pegue o escriba las respuestas (una por línea, numeradas, con viñetas, etc.):").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        self.texto_respuestas = scrolledtext.ScrolledText(entrada_frame, height=15,
                                                          font=('Courier New', 10), relief=tk.SOLID, bd=1)
        self.texto_respuestas.grid(row=1, column=0, columnspan=2, sticky="nsew")

        # --- Botones de Acción ---
        botones_frame = ttk.Frame(entrada_frame)
        botones_frame.grid(row=2, column=0, columnspan=2, pady=(15, 5), sticky=tk.W)

        ttk.Button(botones_frame, text="Procesar Texto",
                   command=self.procesar_respuestas, bootstyle="info").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(botones_frame, text="Vista Previa",
                   command=self.mostrar_vista_previa, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(botones_frame, text="Limpiar Todo",
                   command=self.limpiar, bootstyle="danger-outline").pack(side=tk.LEFT)

        # --- Frame de Estado ---
        self.info_frame = ttk.Labelframe(main_frame, text=" Estado del Proceso ", padding="10")
        self.info_frame.pack(fill=tk.X, pady=(0, 20))

        self.label_info = ttk.Label(self.info_frame, text="Aún no se ha inicializado el generador.",
                                    font=('Helvetica', 10, 'italic'))
        self.label_info.pack()

        # --- Botones Finales ---
        final_frame = ttk.Frame(main_frame)
        final_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(final_frame, text="Generar y Guardar CSV",
                   command=self.generar_csv,
                   bootstyle="success").pack(side=tk.LEFT, padx=5, ipady=5)  # Botón más grande
        ttk.Button(final_frame, text="Cerrar",
                   command=self.destroy, bootstyle="secondary").pack(side=tk.RIGHT, padx=5)

    def inicializar_generador(self):
        total = self.total_preguntas.get()
        self.generador = GeneradorClave(total_preguntas=total)
        self.actualizar_info()
        messagebox.showinfo("Proceso Completado", f"Generador listo para recibir {total} preguntas.")

    def procesar_respuestas(self):
        if not self.generador:
            messagebox.showwarning("Acción Requerida", "Por favor, inicialice el generador primero.")
            return

        texto = self.texto_respuestas.get("1.0", tk.END)
        num_procesadas = self.generador.establecer_respuestas_desde_texto(texto)
        self.actualizar_info()
        validas = self.generador.contar_respuestas_validas()
        messagebox.showinfo("Procesado",
                            f"Se leyeron {num_procesadas} respuestas del texto.\n"
                            f"Total de respuestas válidas ahora: {validas}")

    def limpiar(self):
        self.texto_respuestas.delete("1.0", tk.END)
        if self.generador:
            self.generador.limpiar_respuestas()
            self.actualizar_info()

    def mostrar_vista_previa(self):
        if not self.generador:
            messagebox.showwarning("Acción Requerida", "Por favor, inicialice el generador.")
            return

        ventana_preview = tk.Toplevel(self)
        ventana_preview.title("Vista Previa de la Clave")
        ventana_preview.geometry("600x500")

        texto_preview = scrolledtext.ScrolledText(ventana_preview, font=('Courier New', 10), wrap=tk.NONE)
        texto_preview.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        reporte = self.generador.generar_reporte_texto()
        texto_preview.insert("1.0", reporte)
        texto_preview.config(state=tk.DISABLED)

        ttk.Button(ventana_preview, text="Cerrar Vista",
                   command=ventana_preview.destroy, bootstyle="primary").pack(pady=10)

    def actualizar_info(self):
        if not self.generador:
            self.label_info.config(text="Aún no se ha inicializado el generador.")
            return

        total = self.generador.total_preguntas
        validas = self.generador.contar_respuestas_validas()
        faltantes = total - validas
        porcentaje = (validas / total * 100) if total > 0 else 0

        texto = f"Total de Preguntas: {total} | Respuestas Válidas: {validas} | Faltantes: {faltantes} | Progreso: {porcentaje:.1f}%"
        self.label_info.config(text=texto)

    def generar_csv(self):
        if not self.generador:
            messagebox.showwarning("Acción Requerida", "Por favor, inicialice el generador.")
            return

        if self.generador.contar_respuestas_validas() == 0:
            messagebox.showwarning("Sin Datos", "No hay respuestas válidas para generar el archivo CSV.")
            return

        faltantes = self.generador.obtener_respuestas_faltantes()
        if faltantes:
            if not messagebox.askyesno("Confirmación",
                                       f"Hay {len(faltantes)} preguntas sin respuesta.\n"
                                       f"El archivo se generará con espacios vacíos para estas preguntas.\n"
                                       f"¿Desea continuar?"):
                return

        # Guardar directamente con el formato más común 'n.'
        formato = 'n.'

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"clave_respuestas_{self.total_preguntas.get()}_{formato}.csv"
        )

        if filename:
            if self.generador.generar_csv(filename, formato=formato):
                messagebox.showinfo("Éxito", f"Archivo de clave generado con éxito en:\n{filename}")
                self.destroy()
            else:
                messagebox.showerror("Error Inesperado", "No se pudo generar el archivo CSV.")


# --- APLICACIÓN PRINCIPAL (REDiseñada) ---
class CalificadorApp:
    def __init__(self, root):
        self.root = root
        # Usamos el tema 'cosmo' de ttkbootstrap, que es profesional y limpio
        self.style = ttk.Style(theme='cosmo')
        self.root.title("Lobatchevsky - Sistema Calificador de Exámenes")

        # Geometría adaptable a la pantalla
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = int(screen_width * 0.6)
        height = int(screen_height * 0.7)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(800, 600)

        self.ruta_clave = tk.StringVar()
        self.ruta_respuestas = tk.StringVar()
        self.resultados = None

        self.crear_interfaz()

    def crear_interfaz(self):
        main_frame = ttk.Frame(self.root, padding="15 20")
        main_frame.pack(fill=BOTH, expand=True)

        # Configuración de la grilla responsiva
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # --- Título y Encabezado ---
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        # El logo de Lobatchevsky iría aquí si se carga la imagen
        ttk.Label(header_frame, text="Sistema Calificador de Exámenes",
                  font=('Helvetica', 22, 'bold'), foreground="#0077B6").pack(side=LEFT)

        # --- Panel de Control (Izquierda) ---
        control_panel = ttk.Frame(main_frame)
        control_panel.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(0, 20))
        control_panel.rowconfigure(5, weight=1)  # Espacio para crecer

        # 1. Selección de Archivos
        files_frame = ttk.Labelframe(control_panel, text=" 1. Cargar Archivos ", padding="15")
        files_frame.grid(row=0, column=0, sticky="ew")
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="Clave de Respuestas:").grid(row=0, column=0, sticky=tk.W, pady=(5, 2))
        clave_entry = ttk.Entry(files_frame, textvariable=self.ruta_clave, state="readonly")
        clave_entry.grid(row=1, column=0, columnspan=2, padx=(0, 5), sticky="ew")
        ttk.Button(files_frame, text="Buscar...", command=lambda: self.buscar_archivo('clave'),
                   bootstyle="info-outline").grid(row=1, column=2)

        ttk.Button(files_frame, text="Crear Nueva Clave", command=self.abrir_generador_clave, bootstyle="primary",
                   width="20").grid(row=2, column=0, columnspan=3, pady=(10, 5))

        ttk.Separator(files_frame, orient=HORIZONTAL).grid(row=3, column=0, columnspan=3, pady=15, sticky="ew")

        ttk.Label(files_frame, text="Respuestas de Alumnos:").grid(row=4, column=0, sticky=tk.W, pady=(5, 2))
        resp_entry = ttk.Entry(files_frame, textvariable=self.ruta_respuestas, state="readonly")
        resp_entry.grid(row=5, column=0, columnspan=2, padx=(0, 5), sticky="ew")
        ttk.Button(files_frame, text="Buscar...", command=lambda: self.buscar_archivo('respuestas'),
                   bootstyle="info-outline").grid(row=5, column=2)

        # 2. Acciones
        actions_frame = ttk.Labelframe(control_panel, text=" 2. Acciones Principales ", padding="15")
        actions_frame.grid(row=1, column=0, sticky="new", pady=20)
        actions_frame.columnconfigure(0, weight=1)

        ttk.Button(actions_frame, text="Procesar y Calificar", command=self.procesar, bootstyle="success",
                   width="20").pack(pady=5, ipady=5)

        # 3. Reportes
        reports_frame = ttk.Labelframe(control_panel, text=" 3. Exportar Reportes ", padding="15")
        reports_frame.grid(row=2, column=0, sticky="new")
        reports_frame.columnconfigure(0, weight=1)

        ttk.Button(reports_frame, text="Reporte General (CSV)", command=lambda: self.exportar('csv'),
                   bootstyle="secondary").pack(fill=X, pady=4)
        ttk.Button(reports_frame, text="Reporte General (Excel)", command=lambda: self.exportar('excel'),
                   bootstyle="secondary").pack(fill=X, pady=4)
        ttk.Button(reports_frame, text="Análisis de Errores", command=self.analizar_errores, bootstyle="warning").pack(
            fill=X, pady=4)
        ttk.Button(reports_frame, text="Métricas por Grupo", command=self.generar_metricas, bootstyle="warning").pack(
            fill=X, pady=4)

        # --- Área de Resultados (Derecha) ---
        results_frame = ttk.Labelframe(main_frame, text=" Consola de Resultados ", padding="15")
        results_frame.grid(row=1, column=1, sticky="nsew")
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

        self.resultado_text = scrolledtext.ScrolledText(results_frame, height=20,
                                                        font=('Courier New', 9), wrap=tk.WORD, relief=tk.SOLID, bd=1)
        self.resultado_text.grid(row=0, column=0, sticky="nsew")

        # --- Barra de Estado ---
        self.status_bar = ttk.Label(main_frame, text="Listo para iniciar. Por favor, cargue los archivos.",
                                    anchor=tk.W, bootstyle="inverse-primary", padding=5)
        self.status_bar.grid(row=3, column=0, columnspan=2, sticky="sew", pady=(10, 0))

    def abrir_generador_clave(self):
        VentanaGeneradorClave(self.root)

    def buscar_archivo(self, tipo):
        filename = filedialog.askopenfilename(
            title=f'Seleccionar archivo de {tipo}',
            filetypes=[('Archivos CSV', '*.csv'), ('Todos los archivos', '*.*')],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )
        if filename:
            if tipo == 'clave':
                self.ruta_clave.set(filename)
                self.status_bar.config(text=f"Archivo de clave cargado: {os.path.basename(filename)}")
            else:
                self.ruta_respuestas.set(filename)
                self.status_bar.config(text=f"Archivo de respuestas cargado: {os.path.basename(filename)}")

    def procesar(self):
        self.resultado_text.delete(1.0, tk.END)
        self.resultados = None

        if not self.ruta_clave.get() or not self.ruta_respuestas.get():
            messagebox.showerror("Error de Archivos",
                                 "Debe seleccionar tanto el archivo de clave como el de respuestas.")
            return

        self.status_bar.config(text="Procesando... por favor espere.")
        self.root.update_idletasks()

        clave_df = cargar_datos(self.ruta_clave.get())
        respuestas_df = cargar_datos(self.ruta_respuestas.get())

        if clave_df is None or respuestas_df is None:
            messagebox.showerror("Error de Carga",
                                 "Uno o ambos archivos no se pudieron cargar. Verifique la consola para más detalles.")
            return

        self.resultados = procesar_calificaciones_google_forms(
            clave_df, respuestas_df, MAPEO_MATERIAS,
            COLUMNA_NOMBRE, COLUMNA_EMAIL, COLUMNA_GRUPO
        )

        if not self.resultados:
            messagebox.showerror("Error de Procesamiento",
                                 "No se pudieron procesar las calificaciones. Revise que los formatos de los archivos sean correctos.")
            self.status_bar.config(text="Error durante el procesamiento.")
            return

        self.mostrar_resultados()
        self.status_bar.config(text=f"Procesamiento completado. Se calificaron {len(self.resultados)} alumnos.")

    def mostrar_resultados(self):
        if not self.resultados: return

        self.resultado_text.insert(tk.END, "=" * 80 + "\n")
        self.resultado_text.insert(tk.END, "RESUMEN DE CALIFICACIONES\n".center(80))
        self.resultado_text.insert(tk.END, "=" * 80 + "\n\n")

        for idx, res in enumerate(self.resultados, 1):
            linea = f"#{idx:<3} {res['nombre']:<40} | Aciertos: {res['total_aciertos']}/{res['total_preguntas']} ({res['porcentaje_global']}%)"
            self.resultado_text.insert(tk.END, linea + "\n")
        self.resultado_text.insert(tk.END, "\n" + "=" * 80 + "\n")
        messagebox.showinfo("Proceso Terminado", f"Se han calificado {len(self.resultados)} alumnos exitosamente.")

    def exportar(self, formato):
        if not self.resultados:
            messagebox.showwarning("Sin Datos", "Primero debe procesar las calificaciones para poder exportar.")
            return

        ext, tipo_archivo = (".csv", "CSV") if formato == 'csv' else (".xlsx", "Excel")
        filename = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[(f"Archivos {tipo_archivo}", f"*{ext}")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"Resultados_Generales_{formato}"
        )

        if filename:
            try:
                self.status_bar.config(text=f"Exportando a {formato.upper()}...")
                self.root.update_idletasks()
                if formato == 'csv':
                    exportar_a_csv(self.resultados, filename)
                else:
                    exportar_a_excel(self.resultados, filename)
                messagebox.showinfo("Éxito", f"Archivo exportado correctamente en:\n{filename}")
                self.status_bar.config(text=f"Reporte {formato.upper()} exportado con éxito.")
            except Exception as e:
                messagebox.showerror("Error de Exportación", f"Ocurrió un error al exportar:\n{e}")
                self.status_bar.config(text="Error durante la exportación.")

    def analizar_errores(self):
        if not self.resultados:
            messagebox.showwarning("Sin Datos", "Primero procese las calificaciones para analizar los errores.")
            return

        try:
            from analisis_errores import generar_todos_reportes_errores
        except ImportError:
            messagebox.showerror("Error de Módulo", "El archivo 'analisis_errores.py' no se encuentra.")
            return

        carpeta = filedialog.askdirectory(
            title="Seleccione una carpeta para guardar los reportes de errores",
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )
        if carpeta:
            try:
                self.status_bar.config(text="Generando análisis de errores...")
                self.root.update_idletasks()
                generar_todos_reportes_errores(self.resultados, carpeta)
                messagebox.showinfo("Éxito", f"Reportes de errores generados en la carpeta:\n{carpeta}")
                self.status_bar.config(text="Análisis de errores completado.")
            except Exception as e:
                messagebox.showerror("Error en Análisis", f"Ocurrió un error:\n{e}")
                self.status_bar.config(text="Error generando análisis.")

    def generar_metricas(self):
        if not self.resultados:
            messagebox.showwarning("Sin Datos", "Primero procese las calificaciones para generar métricas.")
            return
        try:
            from metricas_grupos import generar_excel_metricas_grupos
        except ImportError:
            messagebox.showerror("Error de Módulo", "El archivo 'metricas_grupos.py' no se encuentra.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile="Metricas_por_Grupo.xlsx"
        )
        if filename:
            try:
                self.status_bar.config(text="Generando métricas por grupo...")
                self.root.update_idletasks()
                generar_excel_metricas_grupos(self.resultados, filename)
                messagebox.showinfo("Éxito", f"Reporte de métricas generado en:\n{filename}")
                self.status_bar.config(text="Métricas generadas con éxito.")
            except Exception as e:
                messagebox.showerror("Error en Métricas", f"Ocurrió un error:\n{e}")
                self.status_bar.config(text="Error generando métricas.")


def main():
    # Asegurarse de que el directorio de datos existe
    if not os.path.exists(RUTA_DATOS):
        os.makedirs(RUTA_DATOS)

    root = ttk.Window(themename="cosmo")
    app = CalificadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()