import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

nombres_profesora = []
nombres_institucional = []
columnas_profesora = []

def normalizar_nombre(nombre):
    return nombre.strip().lower()

def verificar_columnas_calificaciones(df_profesora, df_institucional):
    if columnas_profesora:  # Verifica si la lista de columnas de la profesora no está vacía
        return all(col in df_institucional.columns for col in columnas_profesora)
    else:
        return False


def transferir_calificaciones(df_profesora, df_institucional):
    if verificar_columnas_calificaciones(df_profesora, df_institucional):
        for col in columnas_profesora:
            df_institucional[col] = df_profesora[col]
        ubicacion_nuevo_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        df_institucional.to_excel(ubicacion_nuevo_excel, index=False)
        mensaje_final.config(text="Calificaciones transferidas y guardadas en un nuevo archivo Excel.")
    else:
        mensaje_final.config(text="Su Excel no contiene las mismas columnas que el Excel institucional.\nRevise las columnas y vuelva a intentar.")



def cargar_archivo_profesora():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        global nombres_profesora
        df = pd.read_excel(archivo)
        nombres_profesora = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        subir_institucional_button["state"] = "normal"

def mostrar_mensaje_error(columnas_faltantes):
    mensaje_emergente = tk.Toplevel(ventana)
    mensaje_emergente.title("Error")
    
    mensaje = "Su Excel no contiene las siguientes columnas que están en el Excel institucional:\n"
    mensaje += "\n".join(columnas_faltantes)
    
    etiqueta = ttk.Label(mensaje_emergente, text=mensaje, wraplength=300)
    etiqueta.pack(padx=20, pady=20)

def cargar_archivo_institucional():
    global nombres_institucional
    if not nombres_profesora:
        messagebox.showinfo("Advertencia", "Primero suba el archivo Excel de la profesora.", icon='warning', default='ok')
        return

    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        df = pd.read_excel(archivo)
        nombres_institucional = [normalizar_nombre(nombre) for nombre in df["Nombre"]]
        columnas_faltantes = verificar_columnas_calificaciones(columnas_profesora, df)

        if not columnas_faltantes:
            transferir_calificaciones(columnas_profesora, df)
        else:
            mostrar_mensaje_error(columnas_faltantes)

# Resto del código sin cambios


def ventana_no_encontrados(mensaje, estudiantes):
    def on_cierre():
        deshabilitar_botones()
        ventana_emergente.destroy()

    ventana_emergente = tk.Toplevel(ventana)
    ventana_emergente.title("Estudiantes no encontrados")
    ancho_pantalla = ventana_emergente.winfo_screenwidth()
    alto_pantalla = ventana_emergente.winfo_screenheight()
    x = (ancho_pantalla - 400) // 2
    y = (alto_pantalla - 200) // 2
    ventana_emergente.geometry(f"400x200+{x}+{y}")
    ventana_emergente.protocol("WM_DELETE_WINDOW", on_cierre)
    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 14))
    label = ttk.Label(ventana_emergente, text=mensaje, foreground="black", wraplength=380)
    label.pack(pady=20)
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
        ventana_emergente.destroy()
        messagebox.showinfo("Descarga completada", f"El listado de estudiantes no encontrados en el Excel Institucional.\nSe ha descargado en formato PDF en la ubicación:\n{ubicacion_pdf}")
        deshabilitar_botones()

def deshabilitar_botones():
    subir_profesora_button["state"] = "disabled"
    subir_institucional_button["state"] = "disabled"
    mensaje_final.config(text="Los nombres no coincidieron\nPor favor verifique y cierre la aplicación.")

ventana = tk.Tk()
ventana.title("Transferencia de Calificaciones")
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()
x = (ancho_pantalla - 400) // 2
y = (alto_pantalla - 200) // 2
ventana.geometry(f"400x200+{x}+{y}")
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12))
style.configure("TLabel", font=("Helvetica", 14))
style.configure("TLabel", foreground="black")
subir_profesora_button = ttk.Button(ventana, text="Subir Excel Profesora", command=cargar_archivo_profesora)
subir_profesora_button.pack(pady=20)
subir_institucional_button = ttk.Button(ventana, text="Subir Excel Institucional", command=cargar_archivo_institucional, state="disabled")
subir_institucional_button.pack()
mensaje_final = ttk.Label(ventana, text="", foreground="black")
mensaje_final.pack()
ventana.resizable(False, False)
ventana.mainloop()