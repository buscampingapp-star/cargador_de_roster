class MainController:
    def __init__(self):
        self.view = None

    def set_view(self, view):
        """Permite al controlador tener referencia a la vista""" 
        self.view = view
    
    