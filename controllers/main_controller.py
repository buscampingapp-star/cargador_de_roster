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
        resultados = []
        for i, ruta in enumerate(rutas, start=1):
            modelo = RosterModel(ruta)
            try:
                modelo.cargar_excel()
                errores = modelo.validar()

                if errores:
                    estado = "Errores detectados"
                else:
                    estado = "Validado OK"
                    nombre_archivo = ruta.split("/")[-1]

                    # Clasificar y acumular en listas
                    for nombre_hoja, df in modelo.dataframes.items():
                        categoria, df_clasificado = modelo.clasificar_archivo(nombre_hoja, df)
                        if categoria:
                            # 🔹 Agregar columna con nombre del archivo
                            df_clasificado = df_clasificado.copy()
                            df_clasificado["ArchivoOrigen"] = nombre_archivo

                            if categoria == "Agents":
                                self.dfs_agents.append(df_clasificado)
                            elif categoria == "Sups":
                                self.dfs_sups.append(df_clasificado)
                            elif categoria == "ACMs":
                                self.dfs_acms.append(df_clasificado)
                            elif categoria == "CCMs":
                                self.dfs_ccms.append(df_clasificado)

                resultados.append((i, ruta.split("/")[-1], estado, errores))

            except Exception as e:
                resultados.append(
                    (i, ruta.split("/")[-1], f"Error al cargar: {e}", [str(e)])
                )

        if self.view:
            self.view.mostrar_resultados(resultados)


    def generar_consolidado(self, salida="consolidado.xlsx"):
        if not any([self.dfs_agents, self.dfs_sups, self.dfs_acms, self.dfs_ccms]):
            if self.view:
                self.view.mostrar_mensaje("No hay archivos validados para consolidar.")
            return

        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            if self.dfs_agents:
                pd.concat(self.dfs_agents, ignore_index=True).to_excel(writer, sheet_name="Agents", index=False)
            if self.dfs_sups:
                pd.concat(self.dfs_sups, ignore_index=True).to_excel(writer, sheet_name="Sups", index=False)
            if self.dfs_acms:
                pd.concat(self.dfs_acms, ignore_index=True).to_excel(writer, sheet_name="ACMs", index=False)
            if self.dfs_ccms:
                pd.concat(self.dfs_ccms, ignore_index=True).to_excel(writer, sheet_name="CCMs", index=False)

        if self.view:
            self.view.mostrar_mensaje(f"✅ Consolidado generado en {salida}")

        # Limpiar listas para evitar duplicados
        self.dfs_agents.clear()
        self.dfs_sups.clear()
        self.dfs_acms.clear()
        self.dfs_ccms.clear()


