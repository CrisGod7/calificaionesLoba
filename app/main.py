# app/main.py
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
from config import (RUTA_DATOS, MAPEO_MATERIAS, COLUMNA_NOMBRE,
                    COLUMNA_EMAIL, COLUMNA_GRUPO)
from data_loader import cargar_datos, validar_estructura_csv
from grader import procesar_calificaciones_google_forms
from reporter import exportar_a_csv, exportar_a_excel
from generador_clave import GeneradorClave


class VentanaGeneradorClave(tk.Toplevel):
    """Ventana para generar archivos de clave de respuestas."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Generador de Clave de Respuestas")
        self.geometry("700x650")  # Aumentado para que se vean todos los botones
        self.configure(bg='#f0f0f0')

        self.generador = None
        self.total_preguntas = tk.IntVar(value=120)

        self.crear_interfaz()

        # Hacer la ventana modal
        self.transient(parent)
        self.grab_set()

    def crear_interfaz(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        ttk.Label(main_frame, text="Generador de Clave de Respuestas",
                  font=('Arial', 14, 'bold')).pack(pady=(0, 20))

        # Frame de configuración
        config_frame = ttk.LabelFrame(main_frame, text="Configuración", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(config_frame, text="Total de preguntas:").pack(side=tk.LEFT, padx=5)
        ttk.Spinbox(config_frame, from_=1, to=200, textvariable=self.total_preguntas,
                    width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_frame, text="Inicializar",
                   command=self.inicializar_generador).pack(side=tk.LEFT, padx=20)

        # Frame de entrada de respuestas
        entrada_frame = ttk.LabelFrame(main_frame, text="Ingresar Respuestas", padding="10")
        entrada_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Label(entrada_frame,
                  text="Pegue las respuestas aquí (una por línea, soporta viñetas, números, etc.):").pack(anchor=tk.W,
                                                                                                          pady=(0, 5))

        self.texto_respuestas = scrolledtext.ScrolledText(entrada_frame, height=15,
                                                          font=('Courier', 10), wrap=tk.WORD)
        self.texto_respuestas.pack(fill=tk.BOTH, expand=True)

        # Botones de acción
        botones_frame = ttk.Frame(entrada_frame)
        botones_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(botones_frame, text="Procesar Respuestas",
                   command=self.procesar_respuestas).pack(side=tk.LEFT, padx=5)
        ttk.Button(botones_frame, text="Limpiar",
                   command=self.limpiar).pack(side=tk.LEFT, padx=5)
        ttk.Button(botones_frame, text="Vista Previa",
                   command=self.mostrar_vista_previa).pack(side=tk.LEFT, padx=5)

        # Frame de información
        self.info_frame = ttk.LabelFrame(main_frame, text="Estado", padding="10")
        self.info_frame.pack(fill=tk.X, pady=(0, 10))

        self.label_info = ttk.Label(self.info_frame, text="No inicializado",
                                    font=('Arial', 9))
        self.label_info.pack()

        # Botones finales
        final_frame = ttk.Frame(main_frame)
        final_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(final_frame, text="Generar CSV",
                   command=self.generar_csv,
                   style='Accent.TButton').pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(final_frame, text="Cerrar",
                   command=self.destroy).pack(side=tk.RIGHT, padx=5, pady=5)

    def inicializar_generador(self):
        """Inicializa el generador con el número de preguntas especificado."""
        total = self.total_preguntas.get()
        self.generador = GeneradorClave(total_preguntas=total)
        self.actualizar_info()
        messagebox.showinfo("Inicializado", f"Generador inicializado para {total} preguntas.")

    def procesar_respuestas(self):
        """Procesa el texto de respuestas ingresado."""
        if self.generador is None:
            messagebox.showwarning("Advertencia", "Primero debe inicializar el generador.")
            return

        texto = self.texto_respuestas.get("1.0", tk.END)
        num_procesadas = self.generador.establecer_respuestas_desde_texto(texto)

        self.actualizar_info()

        messagebox.showinfo("Procesado",
                            f"Se procesaron {num_procesadas} respuestas.\n"
                            f"Total válidas: {self.generador.contar_respuestas_validas()}")

    def limpiar(self):
        """Limpia el texto y las respuestas."""
        self.texto_respuestas.delete("1.0", tk.END)
        if self.generador:
            self.generador.limpiar_respuestas()
            self.actualizar_info()

    def mostrar_vista_previa(self):
        """Muestra una vista previa de las respuestas."""
        if self.generador is None:
            messagebox.showwarning("Advertencia", "Primero debe inicializar el generador.")
            return

        # Crear ventana de vista previa
        ventana_preview = tk.Toplevel(self)
        ventana_preview.title("Vista Previa de Respuestas")
        ventana_preview.geometry("600x500")

        texto_preview = scrolledtext.ScrolledText(ventana_preview, font=('Courier', 9))
        texto_preview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        reporte = self.generador.generar_reporte_texto()
        texto_preview.insert("1.0", reporte)
        texto_preview.config(state=tk.DISABLED)

        ttk.Button(ventana_preview, text="Cerrar",
                   command=ventana_preview.destroy).pack(pady=10)

    def actualizar_info(self):
        """Actualiza la información mostrada."""
        if self.generador is None:
            self.label_info.config(text="No inicializado")
            return

        total = self.generador.total_preguntas
        validas = self.generador.contar_respuestas_validas()
        faltantes = total - validas
        porcentaje = (validas / total * 100) if total > 0 else 0

        texto = f"Total: {total} | Válidas: {validas} | Faltantes: {faltantes} | Completado: {porcentaje:.1f}%"
        self.label_info.config(text=texto)

    def generar_csv(self):
        """Genera el archivo CSV."""
        if self.generador is None:
            messagebox.showwarning("Advertencia", "Primero debe inicializar el generador.")
            return

        if self.generador.contar_respuestas_validas() == 0:
            messagebox.showwarning("Advertencia", "No hay respuestas válidas para generar el CSV.")
            return

        # Preguntar si desea continuar con respuestas faltantes
        faltantes = self.generador.obtener_respuestas_faltantes()
        if faltantes:
            respuesta = messagebox.askyesno("Respuestas Faltantes",
                                            f"Hay {len(faltantes)} preguntas sin respuesta.\n"
                                            f"¿Desea continuar de todos modos?")
            if not respuesta:
                return

        # Seleccionar formato
        ventana_formato = tk.Toplevel(self)
        ventana_formato.title("Seleccionar Formato")
        ventana_formato.geometry("300x200")
        ventana_formato.transient(self)
        ventana_formato.grab_set()

        ttk.Label(ventana_formato, text="Seleccione el formato de columnas:",
                  font=('Arial', 10, 'bold')).pack(pady=20)

        formato_var = tk.StringVar(value='pregunta_n')

        ttk.Radiobutton(ventana_formato, text="pregunta_1, pregunta_2, ...",
                        variable=formato_var, value='pregunta_n').pack(anchor=tk.W, padx=40)
        ttk.Radiobutton(ventana_formato, text="P1, P2, P3, ...",
                        variable=formato_var, value='Pn').pack(anchor=tk.W, padx=40)
        ttk.Radiobutton(ventana_formato, text="1., 2., 3., ...",
                        variable=formato_var, value='n.').pack(anchor=tk.W, padx=40)

        def confirmar():
            ventana_formato.formato_seleccionado = formato_var.get()
            ventana_formato.destroy()

        ttk.Button(ventana_formato, text="Continuar",
                   command=confirmar).pack(pady=20)

        ventana_formato.wait_window()

        if not hasattr(ventana_formato, 'formato_seleccionado'):
            return

        formato = ventana_formato.formato_seleccionado

        # Seleccionar archivo de salida
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"clave_respuestas_{self.total_preguntas.get()}.csv"
        )

        if filename:
            if self.generador.generar_csv(filename, formato=formato):
                messagebox.showinfo("Éxito", f"Archivo generado correctamente:\n{filename}")
                self.destroy()
            else:
                messagebox.showerror("Error", "No se pudo generar el archivo CSV.")


class CalificadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Calificaciones")
        self.root.geometry("900x600")
        self.root.configure(bg='#f0f0f0')

        # Variables
        self.ruta_clave = tk.StringVar()
        self.ruta_respuestas = tk.StringVar()
        self.resultados = None

        self.crear_interfaz()

    def crear_interfaz(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

        # Título
        ttk.Label(main_frame, text="Sistema de Calificaciones",
                  font=('Arial', 16, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Archivo de clave
        ttk.Label(main_frame, text="Archivo de Clave:",
                  font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.ruta_clave, width=60).grid(
            row=1, column=1, padx=5, sticky=(tk.W, tk.E))

        # Frame para botones de clave
        btn_clave_frame = ttk.Frame(main_frame)
        btn_clave_frame.grid(row=1, column=2, padx=5)
        ttk.Button(btn_clave_frame, text="Buscar",
                   command=lambda: self.buscar_archivo('clave')).pack(side=tk.LEFT)
        ttk.Button(btn_clave_frame, text="Crear Nueva",
                   command=self.abrir_generador_clave).pack(side=tk.LEFT, padx=(5, 0))

        # Archivo de respuestas
        ttk.Label(main_frame, text="Respuestas de Alumnos:",
                  font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.ruta_respuestas, width=60).grid(
            row=2, column=1, padx=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Buscar",
                   command=lambda: self.buscar_archivo('respuestas')).grid(row=2, column=2, padx=5)

        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)

        ttk.Button(button_frame, text="Procesar",
                   command=self.procesar).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="CSV",
                   command=lambda: self.exportar('csv')).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="Excel",
                   command=lambda: self.exportar('excel')).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="Errores",
                   command=self.analizar_errores).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="Métricas",
                   command=self.generar_metricas).pack(side=tk.LEFT, padx=3)

        # Área de resultados
        ttk.Label(main_frame, text="Resultados:",
                  font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky=(tk.W, tk.N), pady=(0, 5))

        self.resultado_text = scrolledtext.ScrolledText(main_frame, height=20, width=90,
                                                        font=('Courier', 9), wrap=tk.WORD)
        self.resultado_text.grid(row=5, column=0, columnspan=3, pady=5,
                                 sticky=(tk.W, tk.E, tk.N, tk.S))

        # Barra de estado
        self.status_bar = ttk.Label(main_frame, text="Listo",
                                    relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

    def abrir_generador_clave(self):
        """Abre la ventana del generador de claves."""
        VentanaGeneradorClave(self.root)

    def buscar_archivo(self, tipo):
        filename = filedialog.askopenfilename(
            title=f'Seleccionar archivo de {tipo}',
            filetypes=[('CSV', '*.csv'), ('All', '*.*')],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )
        if filename:
            if tipo == 'clave':
                self.ruta_clave.set(filename)
            else:
                self.ruta_respuestas.set(filename)
            self.status_bar.config(text=f"Archivo: {os.path.basename(filename)}")

    def procesar(self):
        self.resultado_text.delete(1.0, tk.END)
        self.resultados = None

        if not self.ruta_clave.get() or not self.ruta_respuestas.get():
            messagebox.showerror("Error", "Seleccione ambos archivos")
            return

        self.status_bar.config(text="Procesando...")
        self.root.update()

        clave_df = cargar_datos(self.ruta_clave.get())
        respuestas_df = cargar_datos(self.ruta_respuestas.get())

        if clave_df is None or respuestas_df is None:
            messagebox.showerror("Error", "Error al cargar archivos")
            return

        valido_clave, _ = validar_estructura_csv(clave_df, es_clave=True)
        valido_resp, _ = validar_estructura_csv(respuestas_df, es_clave=False)

        if not valido_clave or not valido_resp:
            messagebox.showerror("Error", "Formato incorrecto")
            return

        self.resultados = procesar_calificaciones_google_forms(
            clave_df, respuestas_df, MAPEO_MATERIAS,
            COLUMNA_NOMBRE, COLUMNA_EMAIL, COLUMNA_GRUPO
        )

        if not self.resultados:
            messagebox.showerror("Error", "No se procesaron calificaciones")
            return

        self.mostrar_resultados()
        self.status_bar.config(text=f"Completado: {len(self.resultados)} alumnos")

    def mostrar_resultados(self):
        if not self.resultados:
            return

        self.resultado_text.insert(tk.END, "=" * 80 + "\n")
        self.resultado_text.insert(tk.END, "REPORTE DE CALIFICACIONES\n".center(80))
        self.resultado_text.insert(tk.END, "=" * 80 + "\n\n")

        for idx, resultado in enumerate(self.resultados, 1):
            self.resultado_text.insert(tk.END, f"Alumno {idx}: {resultado['nombre']}\n")
            self.resultado_text.insert(tk.END, f"Email: {resultado.get('email', 'N/A')}\n")
            self.resultado_text.insert(tk.END,
                                       f"Total: {resultado['total_aciertos']}/{resultado['total_preguntas']}\n")
            self.resultado_text.insert(tk.END, "-" * 80 + "\n")

            for materia, datos in sorted(resultado['calificaciones'].items()):
                self.resultado_text.insert(tk.END,
                                           f"  {materia:15s}: {datos['aciertos']:2d}/{datos['total']:2d}\n")

            self.resultado_text.insert(tk.END, "\n")

        self.resultado_text.insert(tk.END, "=" * 80 + "\n")

    def exportar(self, formato):
        if not self.resultados:
            messagebox.showwarning("Advertencia", "Procese las calificaciones primero")
            return

        ext = ".csv" if formato == 'csv' else ".xlsx"
        filename = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[("CSV" if formato == 'csv' else "Excel", f"*{ext}")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile=f"resultados{ext}"
        )

        if filename:
            try:
                if formato == 'csv':
                    exportar_a_csv(self.resultados, filename)
                else:
                    exportar_a_excel(self.resultados, filename)
                messagebox.showinfo("Éxito", f"Exportado: {filename}")
                self.status_bar.config(text="Exportado correctamente")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def analizar_errores(self):
        if not self.resultados:
            messagebox.showwarning("Advertencia", "Procese las calificaciones primero")
            return

        try:
            from analisis_errores import generar_todos_reportes_errores
        except ImportError:
            messagebox.showerror("Error", "Módulo analisis_errores.py no encontrado")
            return

        carpeta = filedialog.askdirectory(
            title="Carpeta para reportes",
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )

        if carpeta:
            try:
                generar_todos_reportes_errores(self.resultados, carpeta)
                messagebox.showinfo("Éxito", f"Reportes en: {carpeta}")
                self.status_bar.config(text="Análisis generado")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def generar_metricas(self):
        if not self.resultados:
            messagebox.showwarning("Advertencia", "Procese las calificaciones primero")
            return

        try:
            from metricas_grupos import generar_excel_metricas_grupos
        except ImportError:
            messagebox.showerror("Error", "Módulo metricas_grupos.py no encontrado")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.',
            initialfile="metricas_grupos.xlsx"
        )

        if filename:
            try:
                generar_excel_metricas_grupos(self.resultados, filename)
                messagebox.showinfo("Éxito", f"Métricas en: {filename}")
                self.status_bar.config(text="Métricas generadas")
            except Exception as e:
                messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    ttk.Style().theme_use('clam')
    CalificadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()