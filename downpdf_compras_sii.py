import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
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
# pip install pyinstaller programa para crear ejecutable de python
# pyinstaller downlvc_sii.py Comando para crear ejecutable

logger = logging.getLogger('registro_Log')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('registro.log')
fh.setLevel(logging.DEBUG)
# logger.addHandler(fh)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)


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

def check_exists_by_xpath(wdrive, xpath):
    try:
        wdrive.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


def check_exists_by_ccs(wdrive, ccspath):
    try:
        wdrive.find_element(By.CSS_SELECTOR, ccspath)
    except NoSuchElementException:
        return False
    return True

# --- Parametros del proceso
rut = '7967830-9'
rutacarpetadescarga = "D:/Users/RYANEZ/Downloads/"
hoy = date.today()
mes = hoy.month
nom_mes = nombremes(mes)
agno = hoy.year
ames= str(agno) + "_" + nom_mes

if not (path.exists('config.ini')):
    logger.critical("Archivo 'config.ini' NO encontrado.")
    quit()

config = ConfigParser()
config.read('config.ini')

db_server = config.get('DATABASE', 'DB_SERVER')
db_name = config.get('DATABASE', 'DB_NAME')
db_user = config.get('DATABASE', 'DB_USER')
db_password = config.get('DATABASE', 'DB_PASSWORD')
rutaregistro= config.get('APPLICATION','PATH_REGISTRO')
rutacarpetadescarga.replace("\\", "\")
rutacarpetadescarga= config.get('APLICATION', 'PATH_DOWNLOAD_FOLDER')
rutacarpetadescarga.replace("\\", "/")
#rutacarpetapdf="D:\Users\RYANEZ\Documents\Madesal\descarga libro Compra_Venta\PDFS\COMPRA"
rutacarpetapdf= config.get('APPLICATION','PATH_PDFS')
rutacarpetapdf.replace("\\", "\")
rut_pdf="15181434-4"
password_pdf="grupom15181"

cantempresas_descargadas= 0
cantempresas_base= 0
error_proceso= False

try:
    conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                              db_server+';DATABASE='+db_name+';UID='+db_user+';PWD=' + db_password)
    # OK! conexión exitosa
    logger.info(f"Conexion exitosa a BD{db_name}.")

except Exception as e:
    error_string = str(e)
    logger.critical(f'Ocurrió un error al conectar a SQL Server:{error_string}')
#    print("Ocurrió un error al conectar a SQL Server: ", e)
    quit()

try:
    #with conexion.cursor() as cursor:
        # En este caso no necesitamos limpiar ningún dato
        cursor= conexion.cursor()
        cursor.execute("SELECT [login_SII_emp], [password_SII_emp],[ruta_destino_SII_LC_emp], [ruta_destino_SII_LC_emp] FROM [dbo].[EMPRESAS] WHERE [Descarga_Libros_SII_emp]='SI';")
        # Con fetchall traemos todas las filas
except Exception as e:
    error_string = str(e)
    logger.critical(f"Ocurrió un error al consultar: {error_string}")
    quit()

try:
    #with conexion.cursor() as cursor:
        empresas = cursor.fetchall()
        cantempresas_base= len(empresas)
        logger.info(f"Empresas a Procesar={cantempresas_base}.")
        # Recorrer y descargar
        for empresa in empresas:
            continuar= True
            cantempresas_descargadas= cantempresas_descargadas + 1
            rut= empresa[0]
            password_sii= empresa[1]
            carpeta_destino_lc= empresa[2]
            carpeta_destino_lv = empresa[3]
            if carpeta_destino_lc[-1] != "\\":
                carpeta_destino_lc= carpeta_destino_lc + "\\"
            carpeta_destino_lc= carpeta_destino_lc + ames + "\\"
            if carpeta_destino_lv[-1] != "\\":
                carpeta_destino_lv= carpeta_destino_lv + "\\"
            carpeta_destino_lv = carpeta_destino_lv + ames + "\\"
            carpeta_destino_lc.replace("\\", "/")
            carpeta_destino_lv.replace("\\", "/")
