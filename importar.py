import pandas as pd
import pyxlsb
"""
Correlativa
-Esta correlativa se hace con el fin de que en dado caso que cambien las posiciones de las variables se ajuste solamente el número en el código
*Nota: Python al igual que la mayoría de lenguajes de programación empiezan a contar desde 0, por lo tanto la primera posición es 0 y no 1.

-[0]: Tipo Ocupa la posición 0 en la base.
-[1]: Cadena Ocupa la posición 1 en la base.
-[2]: Sector Ocupa la posición 2 en la base.
-[3]: Subsector Ocupa la posición 3 en la base.
-[6]: NIT Ocupa la posición 6 en la base.
-[7]: Razon Social Ocupa la posición 7 en la base.
-[8]: País Destino Ocupa la posición 8 en la base.
-[14]: Departamento Origen Ocupa la posición 14 en la base. 
-[35]: 2022 USD (Ene-Jul) Ocupa la posición 35 en la base.
-[37]: 2023 USD (Ene-Jul) Ocupa la posición 37 en la base.

*Se deben ajustar las columnas de acuetrdo a la base de datos que se esté utilizando. En la sección columns_to_select
**Nota en df=... y sale skiprows=range(5), pues esta es la posición donde comienza la base al momento ajustar acorde, no olvidar que se cuenta desde 0
***Una vez se tienen estas variables seleccionadas se deben poner en el orden que sale en la correlativa porque el codigo de procesamiento va a leer en ese orden el nuevo df 
"""
def import_data_from_excel(file_path):
    # Specified columns to select
    columns_to_select_excel = [0, 1, 2, 3, 6, 7, 8, 14, 35, 37]   
    if file_path.endswith('.xlsb'):
        # Read the second sheet of the Excel binary workbook starting from the 6th row and using the specified columns
        df = pd.read_excel(file_path, sheet_name=1, engine='pyxlsb', skiprows=range(5), usecols=columns_to_select)
    elif file_path.endswith('.txt'):
        df = pd.read_csv(file_path, sep=";",thousands=".",decimal=",",encoding='latin-1')    
    elif file_path.endswith('.csv'):
        df=pd.read_csv(file_path, sep=";",thousands=".",decimal=",",encoding='latin-1')
    else:
        df = pd.read_excel(file_path, sheet_name=1, skiprows=range(5), usecols=columns_to_select_excel)
    
    print(df.head(3))
    return df


