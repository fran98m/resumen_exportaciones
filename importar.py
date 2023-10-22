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
-[35]: 2022 USD (Ene-Mes) Ocupa la posición 35 en la base.
-[37]: 2023 USD (Ene-Mes) Ocupa la posición 37 en la base.

*Se deben ajustar las columnas de acuetrdo a la base de datos que se esté utilizando. En la sección columns_to_select
**Nota en df=... y sale skiprows=range(5), pues esta es la posición donde comienza la base al momento ajustar acorde, no olvidar que se cuenta desde 0
***Una vez se tienen estas variables seleccionadas se deben poner en el orden que sale en la correlativa porque el codigo de procesamiento va a leer en ese orden el nuevo df 
"""

def import_data_from_excel(file_path:str)->pd.DataFrame:
    """
    Importar datos de un archivo de Excel, CSV o TXT. Se manejan los archivos xlsb también.

    Parámetros:
    file_path -- Ruta del archivo a importar

    Retorna:
    Un dataframe con los datos importados

    """

    # Columnas especificadas para seleccionar
    columns_to_select_excel = [0, 1, 2, 3, 6, 7, 8, 14, 35, 37]

    # Especificar el tipo de dato para la columna 4 (índice basado en 0)
    column_data_types = {0:str, #Tipo
                         1:str, #Cadena
                         2:str, #Sector
                         3:str, #Subsector
                         6:str, #NIT
                         7:str, #Razon Social
                         8:str, #País Destino
                         14:str, #Departamento Origen
                         35:float, #2022 USD (Ene-Mes)
                         37:float} #2023 USD (Ene-Mes)
    

    # Inicializar un dataframe vacío
    df = pd.DataFrame()

    try:
        if file_path.endswith('.xlsb'):
            # Leer el encabezado por separado
            df = pd.read_excel(file_path, sheet_name=1, engine='pyxlsb', header=5, 
                               usecols=columns_to_select_excel, dtype=column_data_types)


        elif file_path.endswith('.txt') or file_path.endswith('.csv'):
            df = pd.read_csv(file_path, sep=";", thousands=".", decimal=",", encoding='latin-1', dtype=column_data_types)

        else:
            # Leer el encabezado por separado
            header = pd.read_excel(file_path, sheet_name=1, nrows=1, skiprows=5, usecols=columns_to_select_excel, dtype=str)
            # Leer los datos
            data = pd.read_excel(file_path, sheet_name=1, header=None, skiprows=6, usecols=columns_to_select_excel, dtype=column_data_types)
            # Combinar el encabezado y los datos
            df = pd.concat([header, data])

        # Si se importaron los datos correctamente, mostrar mensaje de éxito
        print(f"Datos importados correctamente desde el archivo: {file_path}")

    except Exception as e:
        # Si hubo un error al importar los datos, mostrar el mensaje de error
        print(f"Error al leer el archivo: {file_path}. Error: {e}")

    return df
