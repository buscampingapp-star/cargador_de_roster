import pandas as pd

class RosterModel:
    def __init__(self, filepath):
        self.filepath = filepath
        self.data = None

    def cargar_excel(self):
        """Carga el archivo Excel en un DataFrame"""
        self.data = pd.read_excel(self.filepath)
        return self.data

    def validar(self):
        """Ejemplo de validaciones básicas"""
        errores = []

        if self.data is None:
            errores.append("No se cargó ningún archivo.")
            return errores

        # Validar columnas obligatorias
        columnas_obligatorias = ["Nombre", "Posición"]
        for col in columnas_obligatorias:
            if col not in self.data.columns:
                errores.append(f"Falta columna obligatoria: {col}")

        # Validar que no haya valores vacíos en 'Nombre'
        if "Nombre" in self.data.columns:
            if self.data["Nombre"].isnull().any():
                errores.append("Hay registros sin nombre.")

        return errores
