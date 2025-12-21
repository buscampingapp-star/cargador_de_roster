from models.roster_model import RosterModel

class MainController:
    def __init__(self):
        self.view = None

    def set_view(self, view):
        """Conecta la vista al controlador"""
        self.view = view

    def cargar_archivos(self, rutas):
        """
        Orquesta la carga y validación de múltiples archivos Excel.
        - Recorre cada archivo seleccionado/arrastrado.
        - Crea un modelo (RosterModel) para cada archivo.
        - Carga todas las hojas y aplica las validaciones.
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
