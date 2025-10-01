# app/main.py
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from config import (RUTA_DATOS, MAPEO_MATERIAS, COLUMNA_NOMBRE,
                    COLUMNA_EMAIL, COLUMNA_GRUPO)
from data_loader import cargar_datos, validar_estructura_csv
from grader import procesar_calificaciones_google_forms
from reporter import exportar_a_csv, exportar_a_excel


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
        ttk.Button(main_frame, text="Buscar",
                  command=lambda: self.buscar_archivo('clave')).grid(row=1, column=2, padx=5)

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
        ttk.Button(button_frame, text="Metricas",
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

        self.resultado_text.insert(tk.END, "="*80 + "\n")
        self.resultado_text.insert(tk.END, "REPORTE DE CALIFICACIONES\n".center(80))
        self.resultado_text.insert(tk.END, "="*80 + "\n\n")

        for idx, resultado in enumerate(self.resultados, 1):
            self.resultado_text.insert(tk.END, f"Alumno {idx}: {resultado['nombre']}\n")
            self.resultado_text.insert(tk.END, f"Email: {resultado.get('email', 'N/A')}\n")
            self.resultado_text.insert(tk.END,
                f"Total: {resultado['total_aciertos']}/{resultado['total_preguntas']}\n")
            self.resultado_text.insert(tk.END, "-"*80 + "\n")

            for materia, datos in sorted(resultado['calificaciones'].items()):
                self.resultado_text.insert(tk.END,
                    f"  {materia:15s}: {datos['aciertos']:2d}/{datos['total']:2d}\n")

            self.resultado_text.insert(tk.END, "\n")

        self.resultado_text.insert(tk.END, "="*80 + "\n")

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
                messagebox.showinfo("Exito", f"Exportado: {filename}")
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
            messagebox.showerror("Error", "Modulo analisis_errores.py no encontrado")
            return

        carpeta = filedialog.askdirectory(
            title="Carpeta para reportes",
            initialdir=RUTA_DATOS if os.path.exists(RUTA_DATOS) else '.'
        )

        if carpeta:
            try:
                generar_todos_reportes_errores(self.resultados, carpeta)
                messagebox.showinfo("Exito", f"Reportes en: {carpeta}")
                self.status_bar.config(text="Analisis generado")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def generar_metricas(self):
        if not self.resultados:
            messagebox.showwarning("Advertencia", "Procese las calificaciones primero")
            return

        try:
            from metricas_grupos import generar_excel_metricas_grupos
        except ImportError:
            messagebox.showerror("Error", "Modulo metricas_grupos.py no encontrado")
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
                messagebox.showinfo("Exito", f"Metricas en: {filename}")
                self.status_bar.config(text="Metricas generadas")
            except Exception as e:
                messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    ttk.Style().theme_use('clam')
    CalificadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()