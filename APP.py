import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import os
import time
import importar
import re
import new as datos
import config

file_path = None

def select_file():
    global df, file_path
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb"),("Text Files", "*.txt","*.csv")])
    ano_mes()

def ano_mes():
    if file_path:
        base_name = os.path.basename(file_path)
        filename_without_extension = os.path.splitext(base_name)[0] #Elimina la extensión del archivo
        # Extract year and month from the filename
        match = re.search(r'(\d{4})\s*\((\w+)\)', filename_without_extension) #Para buscar el año se usa d{4} y para el mes \w+ donde d es para digitos y w para letras, el 4 indica 4 digitos seguidos y el w dentro del parentesis indica palabra dentro de un parentesis
        if match:
            #Se añaden a config para poder importar estas variables en el módulo de procesamiento de datos 
            config.ano = match.group(1)  # El primer grupo es el año
            config.mes = match.group(2)  # Se va a tomar Julio sin los paréntesis
            print(config.ano)
            print(config.mes)



app = tk.Tk()
app.title("Automatización Resumen Exportaciones")

select_button = tk.Button(app, text="Selecciona la base", command=select_file)
select_button.pack(pady=20)


app.mainloop()
