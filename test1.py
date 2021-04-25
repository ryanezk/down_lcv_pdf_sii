import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import os
import os.path
from os import path
import glob
from configparser import ConfigParser
import pyodbc
import logging
import shutil

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

# --- Parametros del proceso
rut = '7967830-9'
rutacarpetadescarga = "D:/Users/RYANEZ/Downloads/"
hoy = date.today()
mes = hoy.month
nom_mes = nombremes(mes)
agno = hoy.year

if not (path.exists('config.ini')):
    logger.critical("Archivo 'config.ini' NO encontrado.")
    quit()

config = ConfigParser()
config.read('config.ini')

db_server = config.get('DATABASE', 'DB_SERVER')
db_name = config.get('DATABASE', 'DB_NAME')
db_user = config.get('DATABASE', 'DB_USER')
db_password = config.get('DATABASE', 'DB_PASSWORD')
rutacarpetadescarga= config.get('APLICATION', 'PATH_DOWNLOAD_FOLDER')
rutacarpetadescarga.replace("\\", "/")

try:
    conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                              db_server+';DATABASE='+db_name+';UID='+db_user+';PWD=' + db_password)
    # OK! conexión exitosa
    logger.info(f"Conexion exitosa a BD{db_name}.")

except Exception as e:
    # Atrapar error
    logger.critical('Ocurrió un error al conectar a SQL Server')
#    print("Ocurrió un error al conectar a SQL Server: ", e)

try:
    with conexion.cursor() as cursor:
        # En este caso no necesitamos limpiar ningún dato
        cursor.execute("SELECT [login_SII_emp], [password_SII_emp],[ruta_destino_SII_emp] FROM [dbo].[EMPRESAS] WHERE [Descarga_Libros_SII_emp]='SI';")
        # Con fetchall traemos todas las filas
        empresas = cursor.fetchall()
        # Recorrer y descargar
        for empresa in empresas:
            rut= empresa[0]
            password_sii= empresa[1]
            carpeta_destino= empresa[2]
            carpeta_destino.replace("\\", "/")

            # Eliminar archivos descargados previamente
            nomArchivosCompra = "RCV_COMPRA_REGISTRO_" + rut + "_" + str(agno) + str(mes).zfill(2) + "*.csv"
            ruta = rutacarpetadescarga + nomArchivosCompra
            # print(ruta)
            py_files = glob.glob(ruta)
            for py_file in py_files:
                try:
                    os.remove(py_file)
                except OSError as e:
                    logger.critical(f"Error:{e.strerror}")
            nomArchivosVenta = "RCV_VENTA_" + rut + "_" + str(agno) + str(mes).zfill(2) + "*.csv"
            ruta = rutacarpetadescarga + nomArchivosVenta
            # print(ruta)
            py_files = glob.glob(ruta)
            for py_file in py_files:
                try:
                    os.remove(py_file)
                except OSError as e:
                    logger.critical(f"Error:{e.strerror}")
            # quit()

            # Ejecuta el robot para cada empresa. Descargar libros Compra y Venta
            driver = webdriver.Chrome('../Drivers/chromedriver.exe')
            driver.set_page_load_timeout(10)
            driver.get('https://homer.sii.cl')
            vinculo = driver.find_element(By.XPATH, "//a[contains(text(),'Ingresar a Mi Sii')]")
            vinculo.click()
            time.sleep(1)
            driver.find_element_by_id('rutcntr').send_keys(rut)
            driver.find_element_by_id('clave').send_keys(password_sii)
            btn_ingresar = driver.find_element(By.XPATH, "//button[@id='bt_ingresar']/img")
            btn_ingresar.click()
            time.sleep(2)
            comando = "//div[@id='ModalEmergente']/div/div/div[3]/button"
            if check_exists_by_xpath(driver, comando):
                ventana_aviso = driver.find_element(By.XPATH, comando)
                ventana_aviso.click()
            vinculo = driver.find_element(By.XPATH, "//a[contains(text(),'Servicios online')]")
            vinculo.click()
            time.sleep(1)
            vinculo = driver.find_element(By.XPATH, "(//a[contains(text(),'Factura electrónica')])[3]")
            vinculo.click()
            time.sleep(1)
            vinculo = driver.find_element(By.XPATH, "(//a[contains(text(),'Registro de Compras y Ventas')])[2]")
            vinculo.click()
            time.sleep(1)
            vinculo = driver.find_element(By.XPATH, "//a[contains(@href, 'https://www4.sii.cl/consdcvinternetui')]")
            vinculo.click()
            time.sleep(1)
            periodoMes = driver.find_element(By.XPATH, "//div[2]/select")
            periodoMes.click()
            selperiodoMes = Select(periodoMes)
            selperiodoMes.select_by_visible_text(nom_mes)

            periodoAgno = driver.find_element(By.XPATH, "//select[2]")
            periodoAgno.click()
            selperiodoAgno = Select(periodoAgno)
            selperiodoAgno.select_by_visible_text(str(agno))

            vinculo = driver.find_element(By.XPATH, "//button[contains(.,'Consultar')]")
            vinculo.click()
            time.sleep(1)

            vinculo = driver.find_element(By.XPATH, "//button[contains(.,'Descargar Detalles')]")
            vinculo.click()
            time.sleep(1)
            vinculo = driver.find_element(By.XPATH, "//strong[contains(.,'VENTA')]")
            vinculo.click()
            time.sleep(1)
            vinculo = driver.find_element(By.XPATH, "//button[contains(.,'Descargar Detalles')]")
            vinculo.click()
            time.sleep(1)

            vinculo = driver.find_element(By.XPATH, "//a[contains(.,'Cerrar Sesión')]")
            vinculo.click()

            time.sleep(2)
            driver.close()
            driver.quit()
            logger.info(f"Libros Compra/Venta descargados, rut:{rut}")

            #Copiar archivos descargados a carpeta destino
            ruta = rutacarpetadescarga + nomArchivosCompra
            # print(ruta)
            py_files = glob.glob(ruta)
            for py_file in py_files:
                try:
                    #os.remove(py_file)
                    nomArchivoDestino= path.basename(py_file)
                    shutil.move(py_file, carpeta_destino+nomArchivoDestino)
                    logger.info(f"Archivo Compra({nomArchivoDestino}) copiado a carpeta destino ({carpeta_destino}).")

                except OSError as e:
                    logger.critical(f"Error:{e.strerror}")

                ruta = rutacarpetadescarga + nomArchivosVenta
                # print(ruta)
                py_files = glob.glob(ruta)
                for py_file in py_files:
                    try:
                        # os.remove(py_file)
                        nomArchivoDestino = path.basename(py_file)
                        shutil.move(py_file, carpeta_destino + nomArchivoDestino)
                        logger.info(f"Archivo Venta({nomArchivoDestino}) copiado a carpeta destino ({carpeta_destino}).")

                    except OSError as e:
                        logger.critical(f"Error:{e.strerror}")
except Exception as e:
    logger.critical("Ocurrió un error al consultar: ")
finally:
    conexion.close()
    logger.info("Proceso terminado!")



