import tkinter as tk
from tkinter import filedialog, messagebox, Label, PhotoImage
import os
import importar
import new as datos
import config
import traceback
from procesamiento_en_word import generar_docx
import logging
import warnings
import glob  # Este módulo es útil para buscar archivos que coincidan con un patrón específico

# Configuración de las advertencias para evitar mostrarlas innecesariamente
warnings.filterwarnings("ignore")


def seleccionar_archivo() -> None:
    """
    Función que se ejecuta al hacer clic en el botón de seleccionar archivo.
    Se encarga de abrir el diálogo de selección de archivo, procesar el archivo seleccionado y generar el resumen.
    """
    global df, ruta_del_archivo

    # Configuración del logger
    logging.basicConfig(filename='resumen_exportaciones.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    try:
        tipos_de_archivo = [("Archivos de Excel", ("*.xls", "*.xlsx", "*.xlsm", "*.xlsb")), ("Archivos de Texto", "*.txt"), ("Archivos CSV", "*.csv")]
        ruta_del_archivo = filedialog.askopenfilename(title="Selecciona el archivo de Excel", filetypes=tipos_de_archivo)
        
        if not ruta_del_archivo:
            logging.info("El usuario canceló la selección de archivo.")
            return

        df = importar.import_data_from_excel(ruta_del_archivo)
        logging.info("Archivo importado con éxito.")
        
        boton_seleccionar.config(state=tk.DISABLED)
        
        variables_desde_mes_ano = datos.mes_ano(df)
        logging.info("Año y mes extraídos con éxito.")
        
        variables_desde_totales = datos.totales(df, variables_desde_mes_ano)
        variables_desde_no_mineras = datos.no_mineras(df, variables_desde_totales, variables_desde_mes_ano)
        logging.info("Datos procesados con éxito.")
        
        resumen = generar_docx(variables_desde_totales, variables_desde_no_mineras, variables_desde_mes_ano)
        ruta_de_salida = os.path.join(os.path.dirname(ruta_del_archivo), f"Resumen Exportaciones Enero - {variables_desde_mes_ano['mes']}.docx")
        resumen.save(ruta_de_salida)
        logging.info(f"Documento generado con éxito y guardado en {ruta_de_salida}.")

        messagebox.showinfo("Éxito", f"Se creó el documento. Puedes encontrarlo en: {ruta_de_salida}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió el siguiente error: {str(e)}")
        traza_del_error = traceback.format_exc()
        logging.error(f"Ocurrió un error: {str(e)}\n{traza_del_error}")
        print(e)
        print(traza_del_error)

# Configuración de la ventana principal
ventana_principal = tk.Tk()
ventana_principal.title("Automatización Resumen Exportaciones")
ventana_principal.geometry("1200x400")
ventana_principal.configure(bg="#FFFFFF")

# Directorio actual donde está almacenada la imagen
directorio_imagenes = "/Users/franciscomontalvo/Documents/Resumen_Exportaciones/resumen_exportaciones-1/"

# Descomenta la siguiente línea cuando el usuario final clone el repositorio para que busque en el directorio del script
# directorio_imagenes = os.path.dirname(os.path.abspath(__file__))

imagenes_en_directorio = glob.glob(os.path.join(directorio_imagenes, "*.PNG"))
# Si hay imágenes en el directorio, seleccionar y cargar la primera
if imagenes_en_directorio:
    ruta_imagen = imagenes_en_directorio[0]
    try:
        imagen = PhotoImage(file=ruta_imagen)
        etiqueta_imagen = tk.Label(ventana_principal, image=imagen, bg="white")
        etiqueta_imagen.pack(pady=20)
    except Exception as e:
        logging.error(f"Ocurrió un error al cargar la imagen: {str(e)}")
        messagebox.showerror("Error", f"Ocurrió un error al cargar la imagen: {str(e)}")
else:
    logging.warning("No se encontraron imágenes en el directorio.")
    messagebox.showwarning("Advertencia", "No se encontraron imágenes en el directorio.")

etiqueta = tk.Label(ventana_principal, text="Haz clic en el botón para seleccionar un archivo y generar el resumen.", font=("Arial", 12), bg="white")
etiqueta.pack(pady=20)

boton_seleccionar = tk.Button(ventana_principal, text="Seleccionar archivo", command=seleccionar_archivo, bg="#0000FF", fg="white", font=("Arial", 12))
boton_seleccionar.pack(pady=20)

ventana_principal.mainloop()
