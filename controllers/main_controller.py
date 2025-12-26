import pandas as pd
from models.roster_model import RosterModel

class MainController:
    def __init__(self):
        self.view = None
        # listas para acumular DataFrames validados por grupo
        self.dfs_agents, self.dfs_sups, self.dfs_acms, self.dfs_ccms = [], [], [], []

    def set_view(self, view):
        """Conecta la vista al controlador"""
        self.view = view

    def cargar_archivos(self, rutas):
        """
        Orquesta la carga y validación de múltiples archivos Excel.
        - Recorre cada archivo seleccionado/arrastrado.
        - Crea un modelo (RosterModel) para cada archivo.
        - Carga todas las hojas y aplica las validaciones.
        - Clasifica y acumula los DataFrames validados.
        - Devuelve resultados a la vista.
        """
        resultados = []
        for i, ruta in enumerate(rutas, start=1):
            modelo = RosterModel(ruta)
            try:
                # Cargar todas las hojas del Excel
                modelo.cargar_excel()

                # Ejecutar validaciones
                errores = modelo.validar()

                # Estado según resultado
                if errores:
                    estado = "Errores detectados"
                else:
                    estado = "Validado OK"

                    # Clasificar y acumular en listas
                    for nombre_hoja, df in modelo.dataframes.items():
                        categoria, df_clasificado = modelo.clasificar_archivo(nombre_hoja, df)
                        if categoria == "Agents":
                            self.dfs_agents.append(df_clasificado)
                        elif categoria == "Sups":
                            self.dfs_sups.append(df_clasificado)
                        elif categoria == "ACMs":
                            self.dfs_acms.append(df_clasificado)
                        elif categoria == "CCMs":
                            self.dfs_ccms.append(df_clasificado)

                # Guardar resultado (N°, nombre archivo, estado, lista de errores)
                resultados.append((i, ruta.split("/")[-1], estado, errores))

            except Exception as e:
                # Si ocurre un error al cargar el archivo
                resultados.append(
                    (i, ruta.split("/")[-1], f"Error al cargar: {e}", [str(e)])
                )

        # Actualizar la vista con los resultados
        if self.view:
            self.view.mostrar_resultados(resultados)

    def generar_consolidado(self, salida="consolidado.xlsx"):
        if not any([self.dfs_agents, self.dfs_sups, self.dfs_acms, self.dfs_ccms]):
            if self.view:
                self.view.mostrar_mensaje("No hay archivos validados para consolidar.")
            return

        # Consolidar directamente desde las listas acumuladas
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            if self.dfs_agents:
                pd.concat(self.dfs_agents, ignore_index=True).to_excel(writer, sheet_name="Agents", index=False)
            if self.dfs_sups:
                pd.concat(self.dfs_sups, ignore_index=True).to_excel(writer, sheet_name="Sups", index=False)
            if self.dfs_acms:
                pd.concat(self.dfs_acms, ignore_index=True).to_excel(writer, sheet_name="ACMs", index=False)
            if self.dfs_ccms:
                pd.concat(self.dfs_ccms, ignore_index=True).to_excel(writer, sheet_name="CCMs", index=False)

        # Mostrar mensaje en la vista
        if self.view:
            self.view.mostrar_mensaje(f"✅ Consolidado generado en {salida}")

        # 🔹 Limpiar listas para evitar duplicados en futuras consolidaciones
        self.dfs_agents.clear()
        self.dfs_sups.clear()
        self.dfs_acms.clear()
        self.dfs_ccms.clear()

