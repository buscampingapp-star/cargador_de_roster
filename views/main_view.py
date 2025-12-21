import tkinter as tk
from tkinter import ttk

class MainView(tk.Tk):
    def __init__(self, controller): 
        super().__init__()
        self.controller = controller
        self.title("Cargador de Roster")
        self.geometry("400x300")

