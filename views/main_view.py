import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class MainView(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Cargador de Roster")
        self.geometry("900x600")

        # --- Barra superior con botones y contadores ---
        frame_top = tk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=5)

        # Botón cargar
        self.btn_cargar = tk.Button(frame_top, text="Seleccionar Excel", command=self.on_cargar_excel)
        self.btn_cargar.pack(side="left", padx=10)

        # Botón consolidar
        self.btn_consolidar = tk.Button(frame_top, text="Generar consolidado", command=self.on_generar_consolidado)
        self.btn_consolidar.pack(side="left", padx=10)

        # Contadores
        self.lbl_total_cargados = tk.Label(frame_top, text="Total archivos cargados: 0")
        self.lbl_total_validados = tk.Label(frame_top, text="Total archivos validados: 0")
        self.lbl_total_errores = tk.Label(frame_top, text="Total archivos con errores: 0")

        self.lbl_total_cargados.pack(side="left", padx=20)
        self.lbl_total_validados.pack(side="left", padx=20)
        self.lbl_total_errores.pack(side="left", padx=20)

        # --- Frame principal ---
        frame_main = tk.Frame(self)
        frame_main.pack(expand=True, fill="both", padx=10, pady=10)

        # Tabla de archivos
        frame_archivos = tk.Frame(frame_main)
        frame_archivos.pack(expand=True, fill="both", pady=(0,10))

        self.tree_archivos = ttk.Treeview(frame_archivos, columns=("N°", "Archivo", "Estado"), show="headings", height=8)
        self.tree_archivos.heading("N°", text="N°")
        self.tree_archivos.heading("Archivo", text="Archivo")
        self.tree_archivos.heading("Estado", text="Estado")

        scroll_y1 = ttk.Scrollbar(frame_archivos, orient="vertical", command=self.tree_archivos.yview)
        self.tree_archivos.configure(yscrollcommand=scroll_y1.set)

        self.tree_archivos.pack(side="left", expand=True, fill="both")
        scroll_y1.pack(side="right", fill="y")

        # Tabla secundaria de errores
        frame_errores = tk.Frame(frame_main)
        frame_errores.pack(expand=True, fill="both")

        self.tree_errores = ttk.Treeview(frame_errores, columns=("Error",), show="headings", height=6)
        self.tree_errores.heading("Error", text="Errores detectados")

        scroll_y2 = ttk.Scrollbar(frame_errores, orient="vertical", command=self.tree_errores.yview)
        self.tree_errores.configure(yscrollcommand=scroll_y2.set)

        self.tree_errores.pack(side="left", expand=True, fill="both")
        scroll_y2.pack(side="right", fill="y")

        # Evento al seleccionar archivo
        self.tree_archivos.bind("<<TreeviewSelect>>", self.on_select)

        # Diccionario para guardar errores por archivo
        self.errores_por_archivo = {}

    # --- Métodos de interacción ---
    def on_cargar_excel(self):
        rutas = filedialog.askopenfilenames(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if rutas:
            self.controller.cargar_archivos(rutas)

    def on_generar_consolidado(self): 
        """Invoca al controlador para generar el consolidado""" 
        ruta_salida = filedialog.asksaveasfilename( 
            defaultextension=".xlsx", 
            filetypes=[("Archivos Excel", "*.xlsx")], 
            title="Guardar consolidado como..." ) 
        if ruta_salida: 
            self.controller.generar_consolidado(ruta_salida)

    def mostrar_resultados(self, resultados):
        # Limpiar tablas
        for item in self.tree_archivos.get_children():
            self.tree_archivos.delete(item)
        for item in self.tree_errores.get_children():
            self.tree_errores.delete(item)

        self.errores_por_archivo.clear()

        # Insertar resultados
        total_cargados = len(resultados)
        total_validados = 0
        total_con_errores = 0

        for fila in resultados:
            num, archivo, estado, errores = fila
            self.tree_archivos.insert("", "end", values=(num, archivo, estado))
            self.errores_por_archivo[archivo] = errores

            if estado == "Validado OK":
                total_validados += 1
            if errores:
                total_con_errores += 1

        # Actualizar contadores
        self.lbl_total_cargados.config(text=f"Total archivos cargados: {total_cargados}")
        self.lbl_total_validados.config(text=f"Total archivos validados: {total_validados}")
        self.lbl_total_errores.config(text=f"Total archivos con errores: {total_con_errores}")

    def mostrar_mensaje(self, mensaje):
        """Muestra un mensaje emergente al usuario"""
        messagebox.showinfo("Información", mensaje)

    def on_select(self, event):
        """Mostrar errores del archivo seleccionado en la tabla secundaria"""
        seleccion = self.tree_archivos.selection()
        if seleccion:
            valores = self.tree_archivos.item(seleccion[0], "values")
            archivo = valores[1]
            errores = self.errores_por_archivo.get(archivo, [])
            # limpiar tabla de errores
            for item in self.tree_errores.get_children():
                self.tree_errores.delete(item)
            # insertar errores
            if errores:
                for err in errores:
                    self.tree_errores.insert("", "end", values=(err,))
            else:
                self.tree_errores.insert("", "end", values=("Sin errores. Archivo validado OK.",))
