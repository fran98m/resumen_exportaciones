import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import importar
import re
import new as datos
import config
import procesamiento_en_word as datos

file_path = None

def select_file():
    global df, file_path
    filetypes = [("Excel Files", ("*.xls", "*.xlsx", "*.xlsm", "*.xlsb")), ("Text Files", "*.txt"), ("CSV Files", "*.csv")]
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[])
    ano_mes()

    # Reading the file using the import_data_from_excel function
    df = importar.import_data_from_excel(file_path)
    
    # Extracting year and month using the ano_mes function
    datos.ano_mes(df)
    vars_from_mes_ano = datos.mes_ano(df) 
    # Processing the data using the totales and no_mineras functions
    vars_from_totales = datos.totales(df,vars_from_mes_ano)
    vars_from_no_mineras = datos.no_mineras(df, vars_from_totales,vars_from_mes_ano)
    
    # Generating the Word document
    from procesamiento_en_word import generar_docx  # Importing generar_docx
    resumen = generar_docx(vars_from_totales, vars_from_no_mineras)
    resumen.save("path_where_you_want_to_save.docx")  # Modify this path as needed


def ano_mes():

    # Reading the file using the import_data_from_excel function
    df = importar.import_data_from_excel(file_path)
    
    # Extracting year and month using the ano_mes function
    datos.mes_ano(df)
    
    # Processing the data using the totales and no_mineras functions
    vars_from_totales = datos.totales(df)
    vars_from_no_mineras = datos.no_mineras(df, vars_from_totales)
    
    # Generating the Word document
    from procesamiento_en_word import generar_docx  # Importing generar_docx
    resumen = generar_docx(vars_from_totales, vars_from_no_mineras)
    resumen.save("path_where_you_want_to_save.docx")  # Modify this path as needed

app = tk.Tk()
app.title("Automatización Resumen Exportaciones")

select_button = tk.Button(app, text="Selecciona la base", command=select_file)
select_button.pack(pady=20)


app.mainloop()
