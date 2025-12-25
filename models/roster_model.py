import pandas as pd
import re

class RosterModel:
    # Diccionario centralizado de columnas esperadas por hoja
    COLUMNAS_ESPERADAS = {
        "Agents": [
            "ACM", "Supervisor", "Last Name", "First Name", "Monday", "Tuesday", "Wednesday",
            "Thursday", "Friday", "Saturday", "Sunday", "Wave  / Ingreso", "Comments", "Type",
            "Mon Hrs.", "Tue Hrs.", "Wed Hrs.", "Thu Hrs.", "Fri Hrs.", "Sat Hrs.", "Sun Hrs.",
            "Scheduled Hrs.", "ACDID", "Agent´s Payroll Number", "CCMSID", "LOB", "Secondary Skill",
            "Set Skill Programado", "Equipo Específico", "Detalle de función", "Supervisor´s Payroll Number",
            "ACM´s Payroll Number", "NT Loguin", "TPSouth Hire Date", "Project Hire Date", "Production Hire Date",
            "Tenure", "Site", "Status", "WAHA", "VM", "Disponibilidad de horarios", "RUT / DNI", "RUT 2 / CUIL",
            "Telephone Number", "Mail", "Address", "Municipality", "Log ID", "Extension", "CRM RRSS", "BackOffice",
            "Tools #1", "Tools #2", "Tools #3", "Tools #4", "Skill #1", "Skill #2", "Skill #3", "Skill Detail",
            "Start of Vacation", "End of Vacations", "Advance Payment"
        ],
        "Sups": [
            "ACM", "Supervisor", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday",
            "Supervisor´s Payroll Number", "CCMSID", "LOB", "Waha", "RUT / DNI", "RUT2 / CUIL", "Mail", "Log ID",
            "User Tool", "Skill", "Equipo Específico", "Detalle de función", "NT Login", "Start of Vacation",
            "End of Vacations", "Advance Payment", "Comments", "Disponibilidad de horarios"
        ],
        "ACMs": [
            "CCM", "ACM","Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday",
            "ACM´s Payroll Number", "CCMSID", "LOB", "Waha", "DNI / RUT", "CUIL / RUT2", "Mail", "Log ID",
            "User Tool", "Skill", "Equipo Específico", "Detalle de función", "NT Login", "Start of Vacation",
            "End of Vacations", "Advance Payment", "Comments", "Disponibilidad de horarios"
        ],
        "CCMs": [
            "Gerente", "CCM", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday", "Sunday",
            "CCM´s Payroll Number", "CCMSID", "LOB", "Waha", "DNI / RUT", "CUIL / RUT2", "Mail", "Log ID",
            "User Tool", "Skill", "Equipo Específico", "Detalle de función", "NT Login"
        ]
    }

    def __init__(self, filepath):
        self.filepath = filepath
        self.dataframes = {}

    def cargar_excel(self):
        """Carga todas las hojas del archivo Excel en un dict de DataFrames"""
        self.dataframes = pd.read_excel(self.filepath, sheet_name=None)
        return self.dataframes

    def validar_codigo_archivo(self):
        """Extrae y valida el indicador (PV/Final Roster) y el código de semana en el nombre del archivo."""
        nombre_archivo = self.filepath.split("/")[-1]
        match = re.search(r"(PV|Final Roster)\s+(\d{4}-\d{4})", nombre_archivo)
        if not match:
            return None, None, ["El nombre del archivo no contiene un indicador válido (PV/Final Roster) y un código de semana (####-####)."]
        indicador = match.group(1)
        codigo_semana = match.group(2)
        return indicador, codigo_semana, []

    def validar_columnas(self, indicador, codigo_semana):
        """Valida que las hojas tengan las columnas correctas"""
        errores = []
        for hoja_base, columnas in self.COLUMNAS_ESPERADAS.items():
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
                        errores.append(f"En hoja '{nombre_hoja}' existe columna con un nombre distinto al esperado: {col}")
            else:
                errores.append(f"No se encontró la hoja: {nombre_hoja}")
        return errores

    def validar_celdas_vacias(self, indicador, codigo_semana):
        """Valida que no existan celdas vacías en las columnas obligatorias"""
        errores = []
        for hoja_base, columnas in self.COLUMNAS_ESPERADAS.items():
            nombre_hoja = f"{hoja_base} {indicador} {codigo_semana}"
            if nombre_hoja in self.dataframes:
                df = self.dataframes[nombre_hoja]
                for col in columnas:
                    if col in df.columns:
                        filas_vacias = df[df[col].isnull() | (df[col].astype(str).str.strip() == "")]
                        for idx in filas_vacias.index:
                            errores.append(
                                f"En hoja '{nombre_hoja}' la celda en fila {idx+2}, columna '{col}' está vacía."
                            )
        return errores

    def validar_horarios(self, indicador, codigo_semana):
        """Valida que los valores en el rango de columnas entre Monday y Sunday tengan formato correcto."""
        errores = []
        valores_permitidos = {"LOA", "VAC", "Approved OFF", "OFF", "Training Transfer", "Training NH"}
        patron_hora = re.compile(r"^(?:[01]\d|2[0-3]):[0-5]\d\s-\s(?:[01]\d|2[0-3]):[0-5]\d$")

        for hoja_base in ["Agents", "Sups", "ACMs", "CCMs"]:
            nombre_hoja = f"{hoja_base} {indicador} {codigo_semana}"
            if nombre_hoja in self.dataframes:
                df = self.dataframes[nombre_hoja]
                columnas = list(df.columns)
                try:
                    inicio = columnas.index("Monday")
                    fin = columnas.index("Sunday")
                    columnas_dias = columnas[inicio:fin+1]
                except ValueError:
                    continue

                for dia in columnas_dias:
                    for idx in df.index:
                        if idx < 1:
                            continue
                        valor = df.at[idx, dia]
                        if pd.isnull(valor) or str(valor).strip() == "":
                            continue
                        texto = str(valor).strip()
                        if texto not in valores_permitidos and not patron_hora.match(texto):
                            errores.append(
                                f"Formato inválido en hoja '{nombre_hoja}', fila {idx+2}, columna '{dia}': '{texto}'."
                            )
            else:
                errores.append(f"No se encontró la hoja: {nombre_hoja}")
        return errores

    def validar_horas_programadas(self, indicador, codigo_semana):
        """
        Valida que la suma de Mon Hrs. a Sun Hrs. sea igual al valor de Scheduled Hrs.
        en cada fila de la hoja 'Agents'.
        """
        errores = []
        nombre_hoja = f"Agents {indicador} {codigo_semana}"

        if nombre_hoja in self.dataframes:
            df = self.dataframes[nombre_hoja]
            columnas_dias = ["Mon Hrs.", "Tue Hrs.", "Wed Hrs.", "Thu Hrs.", "Fri Hrs.", "Sat Hrs.", "Sun Hrs."]

            if all(col in df.columns for col in columnas_dias) and "Scheduled Hrs." in df.columns:
                for idx in df.index:
                    suma_dias = df.loc[idx, columnas_dias].apply(pd.to_numeric, errors="coerce").sum()
                    total_programado = pd.to_numeric(df.at[idx, "Scheduled Hrs."], errors="coerce")

                    if not pd.isnull(total_programado) and not pd.isnull(suma_dias):
                        if suma_dias != total_programado:
                            errores.append(
                                f"En hoja '{nombre_hoja}', fila {idx+2}: la suma de horas ({suma_dias}) "
                                f"no coincide con 'Scheduled Hrs.' ({total_programado})."
                            )
            else:
                errores.append(f"En hoja '{nombre_hoja}' faltan columnas de horas o 'Scheduled Hrs.'")
        else:
            errores.append(f"No se encontró la hoja: {nombre_hoja}")

        return errores

    def validar(self):
        errores = []
        indicador, codigo_semana, errores_codigo = self.validar_codigo_archivo()
        errores.extend(errores_codigo)

        if indicador and codigo_semana:
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

            # Validaciones principales
            errores.extend(self.validar_columnas(indicador, codigo_semana))
            errores.extend(self.validar_celdas_vacias(indicador, codigo_semana))
            errores.extend(self.validar_horarios(indicador, codigo_semana))
            errores.extend(self.validar_horas_programadas(indicador, codigo_semana))  # << NUEVA VALIDACIÓN

        return errores
