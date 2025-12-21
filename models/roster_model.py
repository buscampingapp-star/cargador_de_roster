import pandas as pd
import re

class RosterModel:
    def __init__(self, filepath):
        self.filepath = filepath
        self.dataframes = {}

    def cargar_excel(self):
        """Carga todas las hojas del archivo Excel en un dict de DataFrames"""
        self.dataframes = pd.read_excel(self.filepath, sheet_name=None)
        return self.dataframes

    def validar_codigo_archivo(self):
        """Extrae y valida el código de semana en el nombre del archivo"""
        nombre_archivo = self.filepath.split("/")[-1]
        match = re.search(r"\d{4}-\d{4}", nombre_archivo)
        if not match:
            return None, ["El nombre del archivo no contiene un código de semana válido (####-####)."]
        return match.group(0), []

    def validar(self):
        errores = []

        # 1. Extraer código del archivo
        codigo_semana, errores_codigo = self.validar_codigo_archivo()
        errores.extend(errores_codigo)

        if codigo_semana:
            hojas_obligatorias = ["Agents", "Sups", "ACMs", "CCMs"]
            hojas_encontradas = list(self.dataframes.keys())

            # 2. Validar existencia y consistencia de hojas obligatorias
            for hoja in hojas_obligatorias:
                coincidencias = [h for h in hojas_encontradas if h.startswith(hoja)]
                if not coincidencias:
                    # No existe ninguna hoja con ese prefijo
                    errores.append(f"Falta la hoja obligatoria: {hoja} PV {codigo_semana}")
                else:
                    # Existe hoja con ese prefijo, validar código
                    match_hoja = re.search(r"\d{4}-\d{4}", coincidencias[0])
                    if match_hoja and match_hoja.group(0) != codigo_semana:
                        errores.append(
                            f"El código del archivo ({codigo_semana}) no coincide con la hoja '{coincidencias[0]}'."
                        )

        return errores
