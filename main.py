import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QFileDialog, QMessageBox, QWidget, QSizePolicy
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt, QTimer
import os
from importar import *
from procesamiento_datos import *
from procesamiento_en_word import generar_docx
import logging
import warnings


# Configuración de las advertencias para evitar mostrarlas innecesariamente
warnings.filterwarnings("ignore")


# Configura el logging
logger = logging.getLogger('Resumen_Exportaciones')
logger.setLevel(logging.INFO)

# Crea un manejador de archivo para escribir mensajes de log en un archivo
file_handler = logging.FileHandler('exportaciones_log.txt')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Crea un manejador de flujo para escribir mensajes de log en la consola
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Añade los manejadores al logger
logger.addHandler(file_handler)
logger.addHandler(stream_handler)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Configuración básica de la ventana principal
        self.setWindowTitle('Automatización Resumen Exportaciones')
        self.setGeometry(100, 100, 1200, 800)
        self.setAutoFillBackground(True)

        # Widget central
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        # Layout vertical
        layout = QVBoxLayout(self.central_widget)

        # Cargar y mostrar la imagen en una etiqueta
        ruta_imagen = os.path.join(os.path.dirname(__file__), "Procolombia.PNG")
        pixmap = QPixmap(ruta_imagen)
        self.lblImage = QLabel(self)
        self.lblImage.setPixmap(pixmap)
        self.lblImage.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lblImage, alignment=Qt.AlignCenter)

        # Etiqueta informativa
        self.lblInfo = QLabel("Haz clic en el botón para seleccionar un archivo y generar el resumen.", self)
        self.lblInfo.setFont(QFont("Arial", 12))
        self.lblInfo.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lblInfo)

        # Botón para seleccionar archivo
        self.btnSelect = QPushButton('Seleccionar archivo', self)
        self.btnSelect.setFont(QFont("Arial", 12))
        self.btnSelect.setStyleSheet("background-color: #0000FF; color: white; padding: 10px; border-radius: 5px;")
        self.btnSelect.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Para que el botón no se estire
        self.btnSelect.clicked.connect(self.seleccionar_archivo)
        layout.addWidget(self.btnSelect, alignment=Qt.AlignCenter)


        # Etiqueta de créditos
        self.lblCredits = QLabel("Desarrollado por la GIC Procolombia (Coordinación de analítica)", self)
        self.lblCredits.setFont(QFont("Arial", 10))
        self.lblCredits.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lblCredits, alignment=Qt.AlignBottom)  # La etiqueta se añade aquí, alineada en la parte inferior

        # Ajustes finales
        self.central_widget.setLayout(layout)
        self.show()

    def seleccionar_archivo(self):
        # Configuración del logger
        logging.basicConfig(filename='resumen_exportaciones.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

        # Usamos QFileDialog para seleccionar el archivo
        ruta_del_archivo, _ = QFileDialog.getOpenFileName(self, "Selecciona el archivo de Excel", "", "Archivos de Excel (*.xls *.xlsx *.xlsm *.xlsb);;Archivos de Texto (*.txt);;Archivos CSV (*.csv)")

        # Si el usuario cancela la selección, la ruta del archivo será una cadena vacía
        if not ruta_del_archivo:
            logging.info("El usuario canceló la selección de archivo.")
            return

        try:
            df = import_data_from_excel(ruta_del_archivo)
            logging.info("Archivo importado con éxito.")
            
            self.btnSelect.setEnabled(False)  # Deshabilita el botón después de cargar el archivo
            
            variables_desde_mes_ano = mes_ano(df)
            logging.info("Año y mes extraídos con éxito.")
            
            variables_desde_totales = totales(df)
            variables_desde_no_mineras = no_mineras(df, variables_desde_totales)
            logging.info("Datos procesados con éxito.")
            
            resumen = generar_docx(variables_desde_totales, variables_desde_no_mineras, variables_desde_mes_ano)
            ruta_de_salida = os.path.join(os.path.dirname(ruta_del_archivo), f"Resumen Exportaciones Enero - {variables_desde_mes_ano['mes']}.docx")
            resumen.save(ruta_de_salida)
            logging.info(f"Documento generado con éxito y guardado en {ruta_de_salida}.")

            # Usamos QMessageBox para mostrar un mensaje de éxito
            QMessageBox.information(self, "**Éxito**", f"Se creó el documento. Puedes encontrarlo en: {ruta_de_salida}")
            QTimer.singleShot(1000, self.close)
        except Exception as e:
            # Usamos QMessageBox para mostrar un mensaje de error
            QMessageBox.critical(self, "Error", f"Ocurrió el siguiente error: {str(e)}")
            logging.error(f"Ocurrió un error: {e}",exc_info=True)
            QTimer.singleShot(1000, self.close)


if __name__ == '__main__':
    # Crea una aplicación PyQt5
    app = QApplication(sys.argv)

    # Crea una instancia de la clase App, que es nuestra ventana principal
    ex = App()

    # Ejecuta el bucle de eventos de la aplicación hasta que se cierre
    sys.exit(app.exec_())
