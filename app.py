from controllers.main_controller import MainController
from views.main_view import MainView

def main():
    controller = MainController()
    view = MainView(controller)
    controller.set_view(view)
    view.mainloop()

if __name__ == "__main__":
    main()
