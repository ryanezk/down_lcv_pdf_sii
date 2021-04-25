import time
from datetime import datetime
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
import io
# pip install configparser
from configparser import ConfigParser
# pip install pyodbc
import pyodbc
#
import logging
# pip install pytest-shutil
import shutil
import codecs
# pyinstaller downlvc_sii.py Comando para crear ejecutable

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

def copiar_csv(ArchivoOrigen, ArchivoDestino):
    try:
        f_paso= codecs.open(ArchivoOrigen, 'r', "utf-8")
        contenido = f_paso.read().splitlines()
        f2_paso = codecs.open(ArchivoDestino, 'w', "utf-8")
        for line in contenido:
            f2_paso.write(f'{line}\r\n')
    except Exception as e:
        error_string = str(e)
        return error_string
    return ""

def reemplaza_titulos_libro(nombreArchivo):
    try:
        f_paso= codecs.open(nombreArchivo, 'r', "utf-8")
        contenido = f_paso.read().splitlines()
        txt_paso= contenido[0] + ";Centro Costo;Item Gasto"
        contenido[0]= txt_paso
        f2_paso = codecs.open(nombreArchivo, 'w', "utf-8")
        for line in contenido:
            f2_paso.write(f'{line}\r\n')
    except Exception as e:
        error_string = str(e)
        return error_string
    return ""
  
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
rutaregistro= config.get('APPLICATION','PATH_REGISTRO')

ruta= rutaregistro +'Log_LCV_' +ames + "_" +str(hoy.day) + "_" + str(hoy.hour) + "_" + str(hoy.minute)
logger = logging.getLogger(ruta)
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(ruta)
fh.setLevel(logging.DEBUG)
# logger.addHandler(fh)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

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
        cursor.execute("SELECT [nombreEmpresa_emp],[login_SII_emp], [password_SII_emp],[ruta_destino_SII_LC_emp], [ruta_destino_SII_LC_emp] FROM [dbo].[EMPRESAS] WHERE [Descarga_Libros_SII_emp]='SI';")
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
            nomEmpresa= empresa[0]
            rut= empresa[1]
            password_sii= empresa[2]
            carpeta_destino_lc= empresa[3]
            carpeta_destino_lv = empresa[4]
            if carpeta_destino_lc[-1] != "\\":
                carpeta_destino_lc= carpeta_destino_lc + "\\"
            carpeta_destino_lc= carpeta_destino_lc + ames + "\\"
            if carpeta_destino_lv[-1] != "\\":
                carpeta_destino_lv= carpeta_destino_lv + "\\"
            carpeta_destino_lv = carpeta_destino_lv + ames + "\\"
            carpeta_destino_lc.replace("\\", "/")
            carpeta_destino_lv.replace("\\", "/")
            #print(carpeta_destino_lv)
            #quit()

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
            try:
                driver = webdriver.Chrome('./chromedriver.exe')
                driver.set_page_load_timeout(50)
                driver.get('https://homer.sii.cl/index.html')
                time.sleep(2)
                continuar= True
                comando= "//a[contains(text(),'Ingresar a Mi Sii')]"
                if check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                    time.sleep(2)
                    driver.find_element_by_id('rutcntr').send_keys(rut)
                    driver.find_element_by_id('clave').send_keys(password_sii)
                    btn_ingresar = driver.find_element(By.XPATH, "//button[@id='bt_ingresar']/img")
                    btn_ingresar.click()
                    time.sleep(2)
                else:
                    continuar= False
                error_proceso= False
                comando="//a[contains(text(),'Continuar')]"
                representante= False
                if continuar and check_exists_by_xpath(driver, comando):
                    representante= True
                    ventana_aviso = driver.find_element(By.XPATH, comando)
                    ventana_aviso.click()
                    time.sleep(2)

                comando = "//div[@id='ModalEmergente']/div/div/div[3]/button"
                if continuar and check_exists_by_xpath(driver, comando):
                    ventana_aviso = driver.find_element(By.XPATH, comando)
                    ventana_aviso.click()
                time.sleep(2)

                comando="//a[contains(text(),'Servicios online')]"
                if continuar and check_exists_by_xpath(driver, comando):
                    #element = driver.find_element_by_xpath(comando)
                    #driver.execute_script("arguments[0].click();", element)
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False

                time.sleep(2)
                comando= "(//a[contains(text(),'Factura electrónica')])[3]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False
                time.sleep(2)
                comando="(//a[contains(text(),'Registro de Compras y Ventas')])[2]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False
                time.sleep(2)
                comando="//a[contains(@href, 'https://www4.sii.cl/consdcvinternetui')]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False

                time.sleep(8)

                comando="//div[2]/select"
                if continuar and check_exists_by_xpath(driver, comando):
                    periodoMes = driver.find_element(By.XPATH, comando)
                    periodoMes.click()
                    selperiodoMes = Select(periodoMes)
                    selperiodoMes.select_by_visible_text(nom_mes)
                else:
                    continuar= False

                if continuar and representante:
                    comando="//select[@name='rut']"
                    if check_exists_by_xpath(driver, comando):
                        periodoRut = driver.find_element(By.XPATH, comando)
                        periodoRut.click()
                        # selperiodoRut = Select(periodoRut)
                        comando= "//select[@name='rut']/option[text()='" + str(rut) + "']"
                        driver.find_element_by_xpath(comando).click()
                        # selperiodoRut.select_by_value(str(rut))
                        # selperiodoRut.select_by_visible_text(rut)
                    else:
                        continuar= False

                comando="//select[2]"
                if continuar and check_exists_by_xpath(driver, comando):
                    periodoAgno = driver.find_element(By.XPATH, comando)
                    periodoAgno.click()
                    selperiodoAgno = Select(periodoAgno)
                    selperiodoAgno.select_by_visible_text(str(agno))
                else:
                    continuar= False
                comando="//button[contains(.,'Consultar')]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                    continuar_venta= True
                else:
                    continuar= False
                    continuar_venta= False
                time.sleep(2)
                comando="//button[contains(.,'Descargar Detalles')]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False
                time.sleep(2)

                #comando="//strong[contains(.,'VENTA')]"
                comando = ".nav-tabs > li:nth-child(2) strong"
                if continuar_venta and check_exists_by_ccs(driver, comando):
                    #vinculo = driver.find_element(By.XPATH, comando)
                    vinculo= driver.find_element_by_css_selector(comando)
                    vinculo.click()
                    continuar= True
                else:
                    continuar= False
                time.sleep(2)
                comando="//button[contains(.,'Descargar Detalles')]"
                if continuar and check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
                else:
                    continuar= False
                time.sleep(2)

                comando="//a[contains(.,'Cerrar Sesión')]"
                if check_exists_by_xpath(driver, comando):
                    vinculo = driver.find_element(By.XPATH, comando)
                    vinculo.click()
            except Exception as e:
                error_string = str(e)
                logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} ,Ocurrió un error:{error_string}")
                continuar= False
                error_proceso= True

            time.sleep(2)
            driver.close()
            driver.quit()
            #logger.info(f"Libros Compra/Venta descargados, rut:{rut}")

            #Copiar archivos descargados a carpeta destino
            ruta = rutacarpetadescarga + nomArchivosCompra
            #print(f"rut:{rut}, ruta:{ruta}")
            py_files = glob.glob(ruta)
            cantcompra=0
            try:
                os.stat(carpeta_destino_lc)
            except:
                os.mkdir(carpeta_destino_lc)
            carpeta_destino_lc= carpeta_destino_lc + "COMPRA/"
            try:
                os.stat(carpeta_destino_lc)
            except:
                os.mkdir(carpeta_destino_lc)
            for py_file in py_files:
                try:
                    #os.remove(py_file)
                    nomArchivoDestino= path.basename(py_file)
                    shutil.move(py_file, carpeta_destino_lc+nomArchivoDestino)
                    nomArchivo= carpeta_destino_lc+nomArchivoDestino                   
                    logger.info(f"Empresa:{nomEmpresa}, Rut:{rut} Archivo Compra({nomArchivoDestino}) copiado a carpeta destino ({carpeta_destino_lc}).")
                    
                    txt_paso= reemplaza_titulos_libro(nomArchivo)
                    if txt_paso != "":
                        logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Error:{txt_paso}")
                    
                    cantcompra= cantcompra + 1
                except Exception as e:
                    error_string = str(e)
                    logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Ocurrió un error:{error_string}")

            ruta = rutacarpetadescarga + nomArchivosVenta
            # print(ruta)
            py_files = glob.glob(ruta)
            cantventa=0
            try:
                os.stat(carpeta_destino_lv)
            except:
                os.mkdir(carpeta_destino_lv)
            carpeta_destino_lv = carpeta_destino_lv + "VENTA/"
            try:
                os.stat(carpeta_destino_lv)
            except:
                os.mkdir(carpeta_destino_lv)
            for py_file in py_files:
                try:
                    # os.remove(py_file)
                    nomArchivoDestino = path.basename(py_file)
                    archivo_destino= carpeta_destino_lv + nomArchivoDestino
                    error_archivo= copiar_csv(py_file, archivo_destino)
                    if error_archivo !="":
                        logger.info(f"Empresa:{nomEmpresa}, Rut:{rut} Archivo Venta({nomArchivoDestino}) NO copiado. Error:{error_archivo}.")
                    else:    
                        #shutil.move(py_file, carpeta_destino_lv + nomArchivoDestino)
                        logger.info(f"Empresa:{nomEmpresa}, Rut:{rut} Archivo Venta({nomArchivoDestino}) copiado a carpeta destino ({carpeta_destino_lv}).")
                        
                        txt_paso= reemplaza_titulos_libro(archivo_destino)
                        if txt_paso != "":
                            logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Error:{txt_paso}")

                        cantventa= cantventa + 1

                except OSError as e:
                    logger.critical(f"Error:{e.strerror}")
            if error_proceso==True:
                logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Errores en la pagina impidieron descargar archivos.")
                cantempresas_descargadas= cantempresas_descargadas - 1
            else:
                if (cantcompra==0) and (cantventa==0):
                    logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Sin datos para libros.")
                    cantempresas_descargadas= cantempresas_descargadas - 1
except Exception as e:
    error_string = str(e)
    logger.critical(f"Empresa:{nomEmpresa}, Rut:{rut} Ocurrió un error:{error_string}")
finally:
    conexion.close()
    logger.info(f"Proceso terminado! {cantempresas_descargadas} Empresas Descargadas de {cantempresas_base}.")
