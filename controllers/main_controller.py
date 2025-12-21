from models.roster_model import RosterModel

class MainController:
    def __init__(self):
        self.view = None

    def set_view(self, view):
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

                resultados.append((i, ruta.split("/")[-1], estado, errores))
            except Exception as e:
                resultados.append((i, ruta.split("/")[-1], f"Error al cargar: {e}", [str(e)]))

        # Actualizar vista
        if self.view:
            self.view.mostrar_resultados(resultados)
