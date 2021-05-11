import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
#from selenium.webdriver.support.ui import Select
#import org.openqa.selenium.support.ui.WebDriverWait;
#import org.openqa.selenium.ElementClickInterceptedException;
#import org.openqa.selenium.ElementNotInteractableException;
import os
import os.path
from os import path
import glob
# pip install configparser
from configparser import ConfigParser
# pip install pyodbc
import pyodbc
#
import logging
# pip install pytest-shutil
import shutil
# pyinstaller down_pdf_sii.py Comando para crear ejecutable
from csv import DictReader
import pandas as pd

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def nombremes(i):
    switcher = {
        1: 'Enero',
        2: 'Febrero',
        3: 'Marzo',
        4: 'Abril',
        5: 'Mayo',
        6: 'Junio',
        7: 'Julio',
        8: 'Agosto',
        9: 'Septiembre',
        10: 'Octubre',
        11: 'Noviembre',
        12: 'Diciembre'
    }
    return switcher.get(i, "")


# --- Parametros del proceso
rut = '7967830-9'
rutacarpetadescarga = "D:/Users/RYANEZ/Downloads/"
hoy = datetime.now()
mes = hoy.month
nom_mes = nombremes(mes)

agno = hoy.year
ames= str(agno) + "_" + nom_mes

if not (path.exists('config.ini')):
    #logger.critical("Archivo 'config.ini' NO encontrado.")
    quit()

config = ConfigParser()
config.read('config.ini')

db_server = config.get('DATABASE', 'DB_SERVER')
db_name = config.get('DATABASE', 'DB_NAME')
db_user = config.get('DATABASE', 'DB_USER')
db_password = config.get('DATABASE', 'DB_PASSWORD')
rutacarpetadescarga= config.get('APPLICATION', 'PATH_DOWNLOAD_FOLDER')
rutacarpetadescarga.replace("\\", "/")
rutacarpetadescarga_pdf= config.get('APPLICATION', 'PATH_DOWNLOAD_PDF_FOLDER')

rutaregistro= config.get('APPLICATION','PATH_REGISTRO')
#rutacarpetapdf="D:\Users\RYANEZ\Documents\Madesal\descarga libro Compra_Venta\PDFS\COMPRA"
rutacarpetapdf= ""

sendmail_logpdf= config.get('EMAIL', 'SENDMAIL_LOGPDF')
servermail_logpdf= config.get('EMAIL', 'SERVERMAIL')
portmail_logpdf= config.get('EMAIL', 'PORTMAIL')
usermail_logpdf= config.get('EMAIL', 'USERMAIL')
passwordmail_logpdf= config.get('EMAIL', 'PASSWORDMAIL')
destinationmail_logpdf= config.get('EMAIL', 'DESTINATIONMAIL')

archivo_log= rutaregistro +'Log_PDF_' +ames + "_" +str(hoy.day) + "_" + str(hoy.hour) + "_" + str(hoy.minute)+"_prueba.txt"
logger = logging.getLogger(archivo_log)
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(archivo_log)
fh.setLevel(logging.DEBUG)
# logger.addHandler(fh)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

logger.info(f"Conexion exitosa a BD{db_name}.")

logging.shutdown()

if (sendmail_logpdf=="YES") or (sendmail_logpdf=="SI"):
    # Creamos el objeto mensaje
    mensaje = MIMEMultipart()
    
    # Establecemos los atributos del mensaje
    asunto="Descarga PDF Libros SII concluido, " + hoy.strftime("%d/%m/%Y, %H:%M:%S")
    cuerpo = "Se adjunta archivo log para su revision."
    mensaje['From'] = usermail_logpdf
    mensaje['To'] = destinationmail_logpdf
    mensaje['Subject'] = asunto
    # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
    mensaje.attach(MIMEText(cuerpo, 'plain'))
    print(f"Agregamos el cuerpo del mensaje!")
    
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(archivo_log, 'rb')
    nombre_adjunto= path.basename(archivo_log)
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'text/plain')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    mensaje.attach(adjunto_MIME)
    
    # Creamos la conexión con el servidor
    sesion_smtp = smtplib.SMTP(servermail_logpdf, portmail_logpdf)
    print(f"Creamos la conexión con el servidor!, ehlo:{sesion_smtp.ehlo()}")
    # Ciframos la conexión
    sesion_smtp.ehlo()
    sesion_smtp.starttls()
    sesion_smtp.ehlo()
    print(f"Ciframos la conexión!, ehlo:{sesion_smtp.ehlo()}")
    # Iniciamos sesión en el servidor
    sesion_smtp.login(usermail_logpdf,passwordmail_logpdf)
    print(f"Iniciamos sesión en el servidor, usermail_logpdf:{usermail_logpdf}, passwordmail_logpdf:{passwordmail_logpdf}!")
    # Convertimos el objeto mensaje a texto
    texto = mensaje.as_string()
    print(f"Convertimos el objeto mensaje a texto!")
    # Enviamos el mensaje
    #sesion_smtp.sendmail(usermail_logpdf, destinationmail_logpdf, texto)
    sesion_smtp.send_message(mensaje)
    print(f"Enviamos el mensaje!, usermail_logpdf:{usermail_logpdf}, passwordmail_logpdf:{passwordmail_logpdf}")
    # Cerramos la conexión
    sesion_smtp.quit()
    sesion_smtp.close()

