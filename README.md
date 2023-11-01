# Resumen Exportaciones - ProColombia

## Descripción

Este proyecto tiene como objetivo automatizar la generación de resúmenes sobre exportaciones colombianas entre 2017 y 2023. Utiliza Python y diversas bibliotecas para procesar datos y generar un informe detallado en formato Word.

## Módulos y Funcionalidades

### 1. **importar.py**
- **Importación de Datos**: Permite la carga de archivos `.xlsb`, `.txt` y `.csv`.
- **Preprocesamiento**: Ajusta y filtra las columnas necesarias para el análisis.

### 2. **procesamiento_datos.py**
- **Extracción de Fecha**: Identifica el mes y el año de los datos a partir de las columnas.
- **Cálculos de Exportaciones**: Realiza diferentes operaciones y cálculos relacionados con las exportaciones.
- **Filtrado y Agrupación**: Agrupa y filtra los datos según diferentes criterios, como país destino o departamento origen.

### 3. **procesamiento_en_word.py**
- **Generación de Informes**: Crea informes detallados en formato Word basados en los datos procesados.
- **Formateo de Texto**: Asegura que el informe tenga un aspecto profesional, aplicando formatos y estilos específicos.

### 4. **main.py**
- **Interfaz de Usuario**: Proporciona una interacción simple con el usuario para seleccionar archivos y generar el informe.
- **Integración**: Combina las funcionalidades de los otros módulos para llevar a cabo el proceso de principio a fin.

## Instalación

#### Desde una terminal:
1. Clonar este repositorio:
   ```bash
   git clone [URL del repositorio]
    ```
2. Navegar al repositorio:
    ```bash
    cd Resumen_Exportaciones
    ```
3. Instalar las dependencias:
    ```bash
    pip install -r requirements.txt
    ```
4. Ejecutar el script:
    ```python
    python main.py  
    ```
#### Desde VSCode:
1. Clonar el repositorio desde la GUI:
    ```bash
    git clone https://github.com/fran98m/resumen_exportaciones
    ```
2. Instalar las dependencias:
    ```bash 
    pip install -r requierements.txt
    ```
3. Ejecutar main.py desde la GUI
    

# Uso:

- Ejecutas main.py y luego sigues las instruicciones de la pantalla
- Selecciona la base de datos, **ten en cuenta que el archivo se va a guardar en la carpeta de la base de datos**. 

# Licencia

- MIT consulta LICENSE para más información