import tkinter as tk
from tkinter import filedialog, messagebox, Label
import os
import importar
import new as datos
import config
import traceback
from procesamiento_en_word import generar_docx

def select_file():
    global df, file_path
    try:
        filetypes = [("Excel Files", ("*.xls", "*.xlsx", "*.xlsm", "*.xlsb")), ("Text Files", "*.txt"), ("CSV Files", "*.csv")]
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[])
        
        # Check if the user cancelled the file dialog
        if not file_path:
            return

        # Reading the file using the import_data_from_excel function
        df = importar.import_data_from_excel(file_path)

        

        # Deshabilitar el botón después de cargar el archivo
        select_button.config(state=tk.DISABLED)

        # Extrayendo el año y el mes usando la función ano_mes
        vars_from_mes_ano = datos.mes_ano(df) 

        # Procesando los datos usando las funciones totales y no_mineras
        vars_from_totales = datos.totales(df,vars_from_mes_ano)
        vars_from_no_mineras = datos.no_mineras(df, vars_from_totales,vars_from_mes_ano)

        # Generando el documento de Word
        resumen = generar_docx(vars_from_totales, vars_from_no_mineras,vars_from_mes_ano)
        output_path = os.path.join(os.path.dirname(file_path), f"Resumen Exportaciones Enero - {vars_from_mes_ano['mes']}.docx")
        resumen.save(output_path)

        # Notificar al usuario que se ha generado el documento
        messagebox.showinfo("Éxito", f"Se creó el documento lo puede encontrar en: {output_path}")

    except Exception as e:
            messagebox.showerror("Error", f"Ocurrió el siguiente error: {str(e)}")
            error_traceback = traceback.format_exc()

            print(e)
            print(error_traceback)

    finally:
            # Cerrar la ventana de Tkinter después del procesamiento, independientemente del éxito o error
            app.destroy()


        # Edición del diseño de la interfaz gráfica:
    if __name__ == "__main__":
           # Aquí van la configuración de la ventana principal se inicia app que es el gui de tkinter
            app = tk.Tk()
            app.title("Automatización Resumen Exportaciones")
            app.geometry("1200x400")
            app.configure(bg="#87CEFA")
            
            # Aquí se pone la imagen de Procolombia
            corporate_image_path = r"D:\usuarios\Pvein2\OneDrive - PROCOLOMBIA\Escritorio\Francisco\Corrección Resumen Export (Doc) (S)\Procolombia.PNG"  # Reemplazar con la ruta de su imagen
            corporate_image = tk.PhotoImage(file=corporate_image_path)
            image_label = Label(app,bg="#87CEFA",image=corporate_image)
            image_label.pack(pady=20)  # Colocar la imagen con un poco de relleno
            
            # Finalmente se añade una etiqueta con quien desarrolló la aplicación    
            label_text = "Herramienta desarrollada por la GIC Procolombia (Coordinación de Analítica)"
            label = tk.Label(app, text=label_text,bg="#87CEFA", font=('Arial', 12))
            label.pack(side=tk.BOTTOM, pady=10) 


        # Aquí se edita el botón principal de la aplicación
            select_button = tk.Button(app, text="Selecciona la base", command=select_file, bg='#003366', fg='white', borderwidth=0, padx=20, pady=10)
            select_button.pack(pady=20)

            app.mainloop()

