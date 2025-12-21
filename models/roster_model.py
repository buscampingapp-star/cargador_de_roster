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
        """
        Extrae y valida el indicador (PV/Final Roster) y el código de semana en el nombre del archivo.
        Ejemplo: 'GTD Ventas PV 1229-0104.xlsx' → indicador='PV', codigo='1229-0104'
        """
        nombre_archivo = self.filepath.split("/")[-1]
        match = re.search(r"(PV|Final Roster)\s+(\d{4}-\d{4})", nombre_archivo)
        if not match:
            return None, None, ["El nombre del archivo no contiene un indicador válido (PV/Roster Final) y un código de semana (####-####)."]
        indicador = match.group(1)
        codigo_semana = match.group(2)
        return indicador, codigo_semana, []

    def validar_columnas(self, indicador, codigo_semana):
        """Valida que las hojas tengan las columnas correctas"""
        errores = []

        # Diccionario de cabeceras esperadas por hoja
        columnas_esperadas = {
            "Agents": ["ACM", "Supervisor", "Last Name", "First Name", "First Name", "Monday", "Tuesday", "Wednesday",
                        "Thursday", "Friday", "Saturday", "Sunday", "Wave  / Ingreso", "Sunday", "Comments", "Type", "Mon Hrs.", "Tue Hrs.",
                        "Wed Hrs.", "Thu Hrs.", "Fri Hrs.", "Sat Hrs.", "Sun Hrs.", "Scheduled Hrs.", "ACDID", "Agent´s Payroll Number", "CCMSID",
                        "LOB", "Secondary Skill", "Set Skill Programado", "Equipo Específico", "Detalle de función", "Supervisor´s Payroll Number",
                        "ACM´s Payroll Number", "NT Loguin", "TPSouth Hire Date", "Project Hire Date", "Production Hire Date", "Tenure", "Site",
                        "Status", "WAHA", "VM", "Disponibilidad de horarios", "RUT / DNI", "RUT 2 / CUIL", "Telephone Number", "Mail", "Address",
                        "Municipality", "Log ID", "Extension", "CRM RRSS", "BackOffice", "Tools #1", "Tools #2", "Tools #3", "Tools #4", "Skill #1",
                        "Skill #2", "Skill #3", "Skill Detail", "Start of Vacation", "End of Vacations", "Advance Payment"],
            "Sups":   ["ACM", "Supervisor", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday", "Supervisor´s Payroll Number",
                       "CCMSID", "LOB", "WAHA", "RUT / DNI", "RUT 2 / CUIL", "Mail", "Log ID", "User Tool", "Skill", "Equipo Específico", "Detalle de función", 
                       "NT Login", "Start of Vacation", "End of Vacations", "Advance Payment", "Comments", "Disponibilidad de horarios"],
            "ACMs":   ["CCM", "ACM","Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday","ACM´s Payroll Number",
                       "CCMSID", "LOB", "WAHA", "RUT / DNI", "RUT 2 / CUIL", "Mail", "Log ID", "User Tool", "Skill", "Equipo Específico", "Detalle de función", 
                       "NT Login", "Start of Vacation", "End of Vacations", "Advance Payment", "Comments", "Disponibilidad de horarios"],
            "CCMs":   ["Gerente", "CCM", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday","CCM´s Payroll Number",
                       "CCMSID", "LOB", "WAHA", "RUT / DNI", "RUT 2 / CUIL", "Mail", "Log ID", "User Tool", "Skill", "Equipo Específico", "Detalle de función", 
                       "NT Login"]
        }

        for hoja_base, columnas in columnas_esperadas.items():
            nombre_hoja = f"{hoja_base} {indicador} {codigo_semana}"
            if nombre_hoja in self.dataframes:
                df = self.dataframes[nombre_hoja]
                columnas_reales = list(df.columns)

                # Validar columnas esperadas
                for esperado in columnas:
                    if esperado not in columnas_reales:
                        similares = [col for col in columnas_reales if col.lower().strip() == esperado.lower()]
                        if similares:
                            errores.append(
                                f"En hoja '{nombre_hoja}' la columna '{similares[0]}' tiene un nombre distinto al esperado: '{esperado}'."
                            )
                        else:
                            errores.append(f"En hoja '{nombre_hoja}' falta columna obligatoria: {esperado}")

                # Validar columnas inesperadas
                for col in columnas_reales:
                    if col not in columnas and all(col.lower().strip() != esperado.lower() for esperado in columnas):
                        errores.append(f"En hoja '{nombre_hoja}' existe columna inesperada: {col}")
            else:
                errores.append(f"No se encontró la hoja: {nombre_hoja}")

        return errores

    def validar(self):
        errores = []

        # 1. Extraer indicador y código del archivo
        indicador, codigo_semana, errores_codigo = self.validar_codigo_archivo()
        errores.extend(errores_codigo)

        if indicador and codigo_semana:
            # 2. Validar hojas obligatorias y consistencia de código
            hojas_obligatorias = ["Agents", "Sups", "ACMs", "CCMs"]
            hojas_encontradas = list(self.dataframes.keys())

            for hoja in hojas_obligatorias:
                coincidencias = [h for h in hojas_encontradas if h.startswith(hoja)]
                if not coincidencias:
                    errores.append(f"Falta la hoja obligatoria: {hoja} {indicador} {codigo_semana}")
                else:
                    match_hoja = re.search(r"\d{4}-\d{4}", coincidencias[0])
                    if match_hoja and match_hoja.group(0) != codigo_semana:
                        errores.append(
                            f"El código del archivo ({codigo_semana}) no coincide con la hoja '{coincidencias[0]}'."
                        )

            # 3. Validar columnas de cada hoja
            errores.extend(self.validar_columnas(indicador, codigo_semana))

        return errores
