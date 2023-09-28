import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import importar
import new as datos
import config
from procesamiento_en_word import generar_docx

file_path = None

def select_file():
    global df, file_path
    filetypes = [("Excel Files", ("*.xls", "*.xlsx", "*.xlsm", "*.xlsb")), ("Text Files", "*.txt"), ("CSV Files", "*.csv")]
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[])
    
    # Reading the file using the import_data_from_excel function
    df = importar.import_data_from_excel(file_path)
    
    # Extracting year and month using the ano_mes function
    vars_from_mes_ano = datos.mes_ano(df) 
    # Processing the data using the totales and no_mineras functions
    vars_from_totales = datos.totales(df,vars_from_mes_ano)
    vars_from_no_mineras = datos.no_mineras(df, vars_from_totales,vars_from_mes_ano)
    
    # Generating the Word document
    resumen = generar_docx(vars_from_totales, vars_from_no_mineras,vars_from_mes_ano)
    output_path = os.path.join(os.path.dirname(file_path), f"Resumen Exportaciones Enero - {vars_from_mes_ano['mes']}.docx")
    resumen.save(output_path)

app = tk.Tk()
app.title("Automatización Resumen Exportaciones")

select_button = tk.Button(app, text="Selecciona la base", command=select_file)
select_button.pack(pady=20)

app.mainloop()

