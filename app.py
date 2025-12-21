from controllers.main_controller import MainController
from views.main_view import MainView

def main():
    controller = MainController() # 1. Crear instancia del controlador
    view = MainView(controller) # 2. Crear la vista y pasarle el controlador
    controller.set_view(view) # 3. Asociar la vista al controlador (opcional pero útil)
    view.mainloop() # 4. Iniciar el bucle principal de Tkinter

if __name__ == "__main__":
    main()