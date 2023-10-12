import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Variables para almacenar los nombres de estudiantes de ambos archivos

nombres_profesora = []
nombres_institucional = []

def normalizar_nombre(nombre):
    # Esta función normaliza el nombre eliminando espacios en blanco adicionales
    # y convirtiéndolo a minúsculas
    return nombre.strip().lower()

# sistema de interfas del sistema de como tener en claro para todo el sintema de inferzad  de omomc de tener de 
def cargar_archivo_profesora():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        global nombres_profesora
        # Leer el archivo Excel de la profesora y extraer los nombres
        df = pd.read_excel(archivo)
        nombres_profesora = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        # Habilitar el botón "Subir Excel Institucional"
        subir_institucional_button["state"] = "normal"


def cargar_archivo_institucional():
    mensaje_advertencia = "En el Excel que subió anteriormente, los nombres de sus estudiantes, deben estar en el Excel institucional para poder transferir las calificaciones. \nCaso contrario no se preocederá a transferir."
    mensaje_adicional = ""
    mensaje_recordatorio = "¡Recuerde! Los nombres de sus estudiantes, deben ser los mismos que estan en el Excel Institucional"
    
    messagebox.showinfo("Dato Importante", mensaje_advertencia, icon='info', default='ok')
    
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        global nombres_institucional
        # Leer el archivo Excel institucional y extraer los nombres
        df = pd.read_excel(archivo)
        nombres_institucional = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        # Comparar los nombres normalizados
        nombres_no_encontrados = set(nombres_profesora) - set(nombres_institucional)
        if nombres_no_encontrados:
            # Mostrar un mensaje de ventana emergente con la opción de descargar el listado en PDF
            mensaje_adicional = f"Total de estudiantes no encontrados: {len(nombres_no_encontrados)}\n\n{mensaje_recordatorio}"
            ventana_no_encontrados(mensaje_adicional, list(nombres_no_encontrados))


def ventana_no_encontrados(mensaje, estudiantes):

    def on_cierre():
        deshabilitar_botones()
        ventana_emergente.destroy()

    ventana_emergente = tk.Toplevel(ventana)
    ventana_emergente.title("Estudiantes no encontrados")

    # Calcula el ancho y alto de la pantalla
    ancho_pantalla = ventana_emergente.winfo_screenwidth()
    alto_pantalla = ventana_emergente.winfo_screenheight()

    # Calcula las coordenadas para centrar la ventana emergente
    x = (ancho_pantalla - 400) // 2  # 400 es el ancho de tu ventana emergente
    y = (alto_pantalla - 200) // 2   # 200 es la altura de tu ventana emergente

    # Establece las dimensiones y la posición de la ventana emergente
    ventana_emergente.geometry(f"400x200+{x}+{y}")
    ventana_emergente.protocol("WM_DELETE_WINDOW", on_cierre)

    # Configurar un estilo para la ventana emergente
    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 14))

    # Etiqueta para mostrar el mensaje de estudiantes no encontrados
    label = ttk.Label(ventana_emergente, text=mensaje, foreground="black", wraplength=380)
    label.pack(pady=20)

    # Botón para generar y descargar el PDF
    boton_pdf = ttk.Button(ventana_emergente, text="Generar PDF", command=lambda: generar_pdf_estudiantes_no_encontrados(estudiantes, ventana_emergente))
    boton_pdf.pack()



def generar_pdf_estudiantes_no_encontrados(estudiantes, ventana_emergente):
    ubicacion_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")])
    if ubicacion_pdf:
        c = canvas.Canvas(ubicacion_pdf, pagesize=letter)
        c.drawString(72, 750, "Estudiantes no encontrados en el Excel Institucional:")
        y = 730
        for estudiante in estudiantes:
            c.drawString(72, y, estudiante)
            y -= 20
        c.save()
        ventana_emergente.destroy()  # Cerrar la ventana emergente
        messagebox.showinfo("Descarga completada", f"El listado de estudiantes no encontrados en el Excel Institucional. \n Se ha descargado en formato PDF en la ubicación:\n{ubicacion_pdf}")
        deshabilitar_botones()

# Función para deshabilitar los botones y mostrar el mensaje
def deshabilitar_botones():
    subir_profesora_button["state"] = "disabled"
    subir_institucional_button["state"] = "disabled"
    mensaje_final.config(text="Los nombres no coincidieron \npor favor verifique y cierre la aplicación.")

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

# Etiqueta para mostrar el resultado o la alerta
mensaje_final = ttk.Label(ventana, text="", foreground="black")
mensaje_final.pack()

# Personalizar la apariencia de la ventana principal
ventana.resizable(False, False)  # Ventana no redimensionable

# Iniciar la aplicación
ventana.mainloop()



