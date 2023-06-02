import logging
import shutil
from tkinter import messagebox
import os, PyPDF2
from os import mkdir

import pandas as pd


#   FUNCION PARA CORTAR PDFS
def CortarPdf(archivo_entrada,carpeta_salida,limpieza,Export_XLS):

    cantidadPdfCreados = 0

    logging.info('Sub-Proceso: Cortado de PDF iniciado.')

    # Extraccion de datos
    DATA = []

    #Verificacion de existencia de carpeta de salida  
    if os.path.exists(carpeta_salida):
        logging.info('Verificar carpeta Ouput:' + carpeta_salida + ' - Verificacion Exitosa') 
        if limpieza==True:
            for file in os.scandir(carpeta_salida):
                if file.is_file():
                    logging.warning('Archivo a eliminar: ' + file.path)
                    os.remove(file.path)
                elif file.is_dir():
                    logging.warning('Carpeta a eliminar: ' + file.path)
                    try:
                        os.rmdir(file.path)
                    except:
                        shutil.rmtree(file.path)
    else:
        
        logging.warning('No existe carpeta: ' + carpeta_salida + ' - se procede a crearla.')
        mkdir(carpeta_salida)

    carpeta_salida = carpeta_salida + '/Cortar-PDF'
    if os.path.exists(carpeta_salida):
        pass
    else:
        mkdir(carpeta_salida)
        logging.info('Carpeta salida PDF creada :' + carpeta_salida)
    
    try:
        pdf_file = open(archivo_entrada, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        logging.info('Leyendo Archivo PDF: ' + archivo_entrada)
    except:
        logging.fatal('No se ha seleccionado un archivo válido o el archivo no existe \n path: ' + archivo_entrada)
        exit()

    for page_num in range(len(pdf_reader.pages)):

        #  LEER PDF Y PAGINARLO

        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])

    #   si la pagina esta vacia se excluye de la paginacion     ||  Validacion - PaginaBlanco
        page = pdf_reader.pages[page_num]
        if len(page.extractText()) == 0:
            continue

        # EXTRACCION DE DATOS

        text = page.extractText().split('\n')

        IMPORTE = text[14]

        ESTADO = text[12].split(':')[1]

        ADHERENTE = text[17].split(':')[2]

        CODIGO_COMPANIA = '7300'

        DATOS_EXTRAIDOS = {'CODIGO_COMPANIA': CODIGO_COMPANIA,'ADHERENTE':ADHERENTE, 'ESTADO': ESTADO, 'IMPORTE': IMPORTE} 
        
        DATA.append(DATOS_EXTRAIDOS)

        output_filename = carpeta_salida + '/{}_{}.pdf'.format(CODIGO_COMPANIA.strip(), ESTADO.strip())
        
        #  ESCRIBIR PDFS

        with open(output_filename, 'wb') as output:
            pdf_writer.write(output)
            cantidadPdfCreados += 1
            logging.info('Archivo Creado >>> ' + output_filename)
        
    if cantidadPdfCreados:
        logging.info('Pdf Creados: ' + str(cantidadPdfCreados))

    logging.info('Finalizacion de separacion de PDF')

    

    logging.info('Verificacion de existencia del archivo excel salida :' + Export_XLS)


    if os.path.exists(Export_XLS):
        try:
            logging.info("Eliminando documento :" + Export_XLS)
            os.remove(Export_XLS)
        except:                
            logging.info('Error al eliminar excel anterior (Verificar que no esté en uso)')
        
    
    try:
        
        df = pd.DataFrame(DATA) 
        df.to_excel(Export_XLS)

        logging.info('Datos exportados exitosamente')

    except:
        logging.fatal('Error al exportar datos al excel.')

    #messagebox.showinfo(title="Proceso correcto.", message="Los PDF Fueron cortados exitosamente!")

    pdf_file.close()
