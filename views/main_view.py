import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES

class MainView(TkinterDnD.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Cargador de Roster")
        self.geometry("800x500")

        # Botón para cargar archivos
        self.btn_cargar = tk.Button(self, text="Seleccionar Excel", command=self.on_cargar_excel)
        self.btn_cargar.pack(pady=5)

        # Frame principal
        frame_main = tk.Frame(self)
        frame_main.pack(expand=True, fill="both", padx=10, pady=10)

        # --- Tabla de archivos ---
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

        # --- Tabla secundaria de errores ---
        frame_errores = tk.Frame(frame_main)
        frame_errores.pack(expand=True, fill="both")

        self.tree_errores = ttk.Treeview(frame_errores, columns=("Error",), show="headings", height=6)
        self.tree_errores.heading("Error", text="Errores detectados")

        scroll_y2 = ttk.Scrollbar(frame_errores, orient="vertical", command=self.tree_errores.yview)
        self.tree_errores.configure(yscrollcommand=scroll_y2.set)

        self.tree_errores.pack(side="left", expand=True, fill="both")
        scroll_y2.pack(side="right", fill="y")

        # Drag & Drop en tabla de archivos
        self.tree_archivos.drop_target_register(DND_FILES)
        self.tree_archivos.dnd_bind('<<Drop>>', self.on_drop)

        # Evento al seleccionar archivo
        self.tree_archivos.bind("<<TreeviewSelect>>", self.on_select)

        # Diccionario para guardar errores por archivo
        self.errores_por_archivo = {}

    def on_cargar_excel(self):
        rutas = filedialog.askopenfilenames(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if rutas:
            self.controller.cargar_archivos(rutas)

    def on_drop(self, event):
        rutas = self.tk.splitlist(event.data)
        self.controller.cargar_archivos(rutas)

    def mostrar_resultados(self, resultados):
        # Limpiar tablas
        for item in self.tree_archivos.get_children():
            self.tree_archivos.delete(item)
        for item in self.tree_errores.get_children():
            self.tree_errores.delete(item)

        self.errores_por_archivo.clear()

        # Insertar resultados
        for fila in resultados:
            num, archivo, estado, errores = fila
            self.tree_archivos.insert("", "end", values=(num, archivo, estado))
            self.errores_por_archivo[archivo] = errores

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
