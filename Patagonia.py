#   --------------------------------------------------------------------------------------------------------------
#                                           Importacion de modulos externos 
#   --------------------------------------------------------------------------------------------------------------
import logging
import os
import sys
import tkinter as tk
import time
import pandas as pd
from datetime import datetime
from tkinter import messagebox,ttk
#   --------------------------------------------------------------------------------------------------------------


# -------------------------------------------------------------------------------------------------------------- 
#                                           Importacion de modulos ROBOTI
# --------------------------------------------------------------------------------------------------------------
from PDFParticionador import CortarPdf
#  
# -------------------------------------------------------------------------------------------------------------- 



# -------------------------------------------------------------------------------------------------------------- 
#                                           Configuracion de Robot
# -------------------------------------------------------------------------------------------------------------- 

# Diccionario de configuraciones
config_bot = {
    'NombreRobot': "",
    'CarpetaInput': "",
    'CarpetaOutput': "",
    'CarpetaLogs': ""
}

# Variables booleanas de validacion
LecturaConfig = False
ProcesoIniciado = False

# Configuracion de formato de fecha
def current_date_format(date):
    months = ("Enero", "Febrero", "Marzo", "Abri", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} de {} del {}".format(day, month, year)
    return messsage

# -------------------------------------------------------------------------------------------------------------- 
#                                          Carga de variables del Settings
# -------------------------------------------------------------------------------------------------------------- 
# Selecciono path
PathSettings = os.path.join(os.getcwd(), 'data', 'Settings.xlsx')

# Lectura de archivo
try:
    # Lectura de excel "Settings.xlsx" > si existe toma los valores de la hoja "Config" y los mete en un data frame
    archivo_excel = pd.ExcelFile(PathSettings)
    df = archivo_excel.parse('Config')

    # Bucle para buscar las key del diccionario en el data frame
    for clave in config_bot:
        valor = df.loc[df['KEY'] == clave, 'VALUE'].values[0]
        config_bot[clave] = valor

    # Validacion de existencia y captura de datos del config
    LecturaConfig = True
except:
    #En caso que no exista el archivo Settings, finalizar ejecución.
    LecturaConfig = False
    
    

# Captura de horario exacto de inicio de robot
inicioRobot = time.time()
# -------------------------------------------------------------------------------------------------------------- 

# -------------------------------------------------------------------------------------------------------------- 
#                                           Configuracion de Logs
# -------------------------------------------------------------------------------------------------------------- 

# Configuraciones generales
ProcessName = config_bot['NombreRobot']
fechaActual = datetime.now()
PathLogs = os.path.join(config_bot['CarpetaLogs'], str(current_date_format(fechaActual)) + '.log')
FormatLogs = '%(asctime)s - %(levelname)s - %(message)s - ' + ProcessName

logging.basicConfig(level=0,
                    format=FormatLogs,
                    filename=PathLogs,
                    filemode='a')


# -------------------------------------------------------------------------------------------------------------- 
#                                     Finalizacion de proceso si no existe SETTINGS                       |
# -------------------------------------------------------------------------------------------------------------- 
if LecturaConfig == False:
    messagebox.showinfo(title="ERROR DE CONFIG", message='El archivo Settings no existe, por favor verificar la ruta: ' + PathSettings)
    sys.exit()
# --------------------------------------------------------------------------------------------------------------


# -------------------------------------------------------------------------------------------------------------- 
#                                           Inicio de precesamiento del Robot
# -------------------------------------------------------------------------------------------------------------- 
def EJECUTAR():
#   -------------------------------------------------- 
    logging.info("Proceso iniciado")
    # Aca colocar los llamados a las funciones
#   --------------------------------------------------
    
    # Colocar todas las funciones necesarias

    CortarPdf(archivo_entrada="C:/RPA/Roboti/Patagonia/Input/Input.pdf",carpeta_salida="C:/RPA/Roboti/Patagonia/Output",limpieza=True,Export_XLS="C:/RPA/Roboti/Patagonia/Output/Salida.xlsx")

#   --------------------------------------------------
    # Finalizacion de ROBOTI
    FIN()
#   --------------------------------------------------

# Funcion de finalización -----------

def FIN():

    # Cálculo de tiempo de ejecución
    finRobot = time.time()
    tiempoEjecucion = round(-1 * (inicioRobot - finRobot),2)

    logging.info('Fin de proceso - Tiempo de ejecución: ' + str(tiempoEjecucion))
    window.quit()

# --------------------------------------------------------------------------------------------------------------
#                                           Creacion de ventana de INICIO
# --------------------------------------------------------------------------------------------------------------
# ventana para ejecutar

window = tk.Tk()
window.title('ROBOTI')
window.resizable(0,0)

# Configuracion de tamaño y posicion inicial de ventana
ancho_ventana = 300
alto_ventana = 200
x_ventana = window.winfo_screenwidth() // 2 - ancho_ventana // 2
y_ventana = window.winfo_screenheight() // 2 - alto_ventana // 2
posicion = str(ancho_ventana) + "x" + str(alto_ventana) + "+" + str(x_ventana) + "+" + str(y_ventana)
window.geometry(posicion)

# ------------------------------------------------------------------------------------------------------------------------

# VENTANA DE EJECUCION {

#------------------ TITULO --------------------------
Titulo = ttk.Label(window,text=ProcessName)
Titulo.pack(anchor="center")
Titulo.place(x=40, y=20, width=200, height=30)

#------------------ BOTONES -------------------------
BtEjecutar = ttk.Button(window, text="Ejecutar", command=EJECUTAR)
BtEjecutar.pack()
BtEjecutar.place(x=30, y=140, width=100, height=30)

BtCancel = ttk.Button(window, text="Salir",command=FIN)
BtCancel.pack()
BtCancel.place(x=160, y=140, width=100, height=30)

#------------------ EJECUCION DE VENTANA ------------
window.mainloop()

# }


