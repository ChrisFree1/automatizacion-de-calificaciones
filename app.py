import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Variables para almacenar los nombres de estudiantes de ambos archivos
nombres_profesora = []
nombres_institucional = []

# Variables para almacenar las columnas de calificaciones
columnas_calificaciones_profesora = ["Primer Parcial", "Segundo Parcial", "Nota Final"]
columnas_calificaciones_institucional = []

def normalizar_nombre(nombre):
    # Esta función normaliza el nombre eliminando espacios en blanco adicionales
    # y convirtiéndolo a minúsculas
    return nombre.strip().lower()

def cargar_archivo_profesora():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        global nombres_profesora, columnas_calificaciones_profesora
        # Leer el archivo Excel de la profesora y extraer los nombres
        df = pd.read_excel(archivo)
        nombres_profesora = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        # Al cargar el archivo de la profesora, habilitar el botón "Subir Excel Institucional"
        columnas_calificaciones_profesora = df.columns
        subir_institucional_button["state"] = "normal"
        
def cargar_archivo_institucional():
    mensaje_advertencia = "En el Excel que subió anteriormente, los nombres de sus estudiantes deben estar en el Excel institucional para poder transferir las calificaciones. \nCaso contrario no se procederá a transferir."
    mensaje_adicional = ""
    mensaje_recordatorio = "¡Recuerde! Los nombres de sus estudiantes deben ser los mismos que están en el Excel Institucional."
    
    messagebox.showinfo("Dato Importante", mensaje_advertencia, icon='info', default='ok')
    
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        global nombres_institucional, columnas_calificaciones_institucional
        # Leer el archivo Excel institucional y extraer los nombres
        df = pd.read_excel(archivo)
        nombres_institucional = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        # Al cargar el archivo institucional, habilitar el botón "Transferir Calificaciones" y verificar las columnas
        columnas_calificaciones_institucional = df.columns
        if not list(columnas_calificaciones_profesora) == list(columnas_calificaciones_institucional):
            mensaje_adicional = "Las siguientes columnas de calificaciones no coinciden con el Excel institucional:\n"
            columnas_no_coincidentes = set(columnas_calificaciones_profesora).difference(columnas_calificaciones_institucional)
            mensaje_adicional += f"- {', '.join(columnas_no_coincidentes)}\nPor favor, verifique los nombres de las columnas."
            ventana_no_encontrados(mensaje_adicional, [])

        else:
            transferir_calificaciones_button["state"] = "normal"


def ventana_no_encontrados(mensaje, columnas_no_coincidentes):
    def on_cierre():
        deshabilitar_botones()
        ventana_emergente.destroy()

    ventana_emergente = tk.Toplevel(ventana)
    ventana_emergente.title("Advertencia")

    # Calcula el ancho y alto de la pantalla
    ancho_pantalla = ventana_emergente.winfo_screenwidth()
    alto_pantalla = ventana_emergente.winfo_screenheight()

    # Calcula las coordenadas para centrar la ventana emergente
    x = (ancho_pantalla - 400) // 2  # 400 es el ancho de tu ventana emergente
    y = (alto_pantalla - 200) // 2   # 200 es la altura de tu ventana emergente

    # Calcula la altura de la ventana emergente en función del número de columnas no coincidentes
    altura_ventana = max(200, 100 + len(columnas_no_coincidentes) * 20)

    # Establece las dimensiones y la posición de la ventana emergente
    ventana_emergente.geometry(f"400x{altura_ventana}+{x}+{y}")
    ventana_emergente.protocol("WM_DELETE_WINDOW", on_cierre)

    # Configurar un estilo para la ventana emergente
    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 14))

    # Etiqueta para mostrar el mensaje de advertencia
    label = ttk.Label(ventana_emergente, text=mensaje, foreground="black", wraplength=380)
    label.pack(pady=20)

# Resto del código sin cambios

def transferir_calificaciones():
    if not nombres_profesora or not nombres_institucional:
        messagebox.showerror("Error", "Primero cargue los archivos de la profesora y el institucional.")
        return

    if not list(columnas_calificaciones_profesora) == list(columnas_calificaciones_institucional):
        mensaje_adicional = "Las siguientes columnas de calificaciones no coinciden con el Excel institucional:\n"
        columnas_no_coincidentes = set(columnas_calificaciones_profesora).difference(columnas_calificaciones_institucional)
        mensaje_adicional += f"- {', '.join(columnas_no_coincidentes)}\nPor favor, verifique los nombres de las columnas."
        ventana_no_encontrados(mensaje_adicional, columnas_no_coincidentes)
    else:
        # Normalizar los nombres para asegurarse de que coincidan
        nombres_profesora = [normalizar_nombre(nombre) for nombre in nombres_profesora]
        nombres_institucional = [normalizar_nombre(nombre) for nombre in nombres_institucional]

        # Crear un diccionario para mapear los nombres de la profesora a los nombres institucionales
        mapeo_nombres = dict(zip(nombres_profesora, nombres_institucional))

        # Crear un DataFrame para las calificaciones
        df_profesora = pd.read_excel(archivo_profesora)
        df_institucional = pd.read_excel(archivo_institucional)

        # Actualizar las notas en el DataFrame institucional
        for columna in columnas_calificaciones_profesora:
            df_institucional[columna] = df_profesora["Nombre"].map(mapeo_nombres).map(df_profesora[columna])

        # Guardar el archivo actualizado
        ubicacion_nuevo_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        if ubicacion_nuevo_excel:
            df_institucional.to_excel(ubicacion_nuevo_excel, index=False)
            mensaje_exitoso = "Las calificaciones han sido transferidas exitosamente."
            messagebox.showinfo("Transferencia Exitosa", mensaje_exitoso)
            deshabilitar_botones()


# Función para deshabilitar los botones y mostrar el mensaje
def deshabilitar_botones():
    subir_profesora_button["state"] = "disabled"
    subir_institucional_button["state"] = "disabled"
    transferir_calificaciones_button["state"] = "disabled"
    mensaje_final.config(text="La transferencia de calificaciones ha sido completada.\nPor favor, verifique y cierre la aplicación.")

# --------------------------------------- Interfaces de mi aplicacion -----------------------------------

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Transferencia de Calificaciones")

# Calcula el ancho y alto de la pantalla
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()

# Calcula las coordenadas para centrar la ventana principal
x = (ancho_pantalla - 400) // 2  # 400 es el ancho de tu ventana principal
y = (alto_pantalla - 200) // 2   # 200 es la altura de tu ventana principal

# Establece las dimensiones y la posición de la ventana principal
ventana.geometry(f"400x200+{x}+{y}")

# Configurar un estilo para hacer que la GUI sea más moderna
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12))
style.configure("TLabel", font=("Helvetica", 14))

# Restaurar la apariencia de los mensajes a negro
style.configure("TLabel", foreground="black")

# Botón para cargar el archivo Excel de la profesora
subir_profesora_button = ttk.Button(ventana, text="Subir Excel Profesora", command=cargar_archivo_profesora)
subir_profesora_button.pack(pady=20)

# Botón para cargar el archivo Excel institucional (inicialmente deshabilitado)
subir_institucional_button = ttk.Button(ventana, text="Subir Excel Institucional", command=cargar_archivo_institucional, state="disabled")
subir_institucional_button.pack()

# Botón para transferir calificaciones (inicialmente deshabilitado)
transferir_calificaciones_button = ttk.Button(ventana, text="Transferir Calificaciones", command=transferir_calificaciones, state="disabled")
transferir_calificaciones_button.pack()

# Etiqueta para mostrar el resultado o la alerta
mensaje_final = ttk.Label(ventana, text="", foreground="black")
mensaje_final.pack()

# Personalizar la apariencia de la ventana principal
ventana.resizable(False, False)  # Ventana no redimensionable

# Iniciar la aplicación
ventana.mainloop()
