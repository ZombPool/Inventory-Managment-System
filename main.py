import tkinter as tk
from gui.pestanas import crear_pestanas

def main():
    root = tk.Tk()
    root.title("Sistema de Gestión en Almacén de Equipos y Herramientas")
    root.geometry("1000x600")

    archivo_excel = r"C:\Users\Paulo\Documents\DOCUMENTS EPCOMM\Proyectos de automatización\P-6. Administración a Almacén de equipos\data\Sistema de Gestion en Almacén de Equipos y Herramientas.xlsx"  # Ajusta la ruta según tu archivo
    crear_pestanas(root, archivo_excel)

    root.mainloop()

if __name__ == "__main__":
    main()
