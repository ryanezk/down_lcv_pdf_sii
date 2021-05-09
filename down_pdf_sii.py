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


def get_serial_number_of_physical_disk(drive_letter='C:'):
    import wmi
    
    c = wmi.WMI()
    logical_disk = c.Win32_LogicalDisk(Caption=drive_letter)[0]
    partition = logical_disk.associators()[1]
    physical_disc = partition.associators()[0]
    return physical_disc.SerialNumber

    
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


def abreviaciontipodocsii(tipo_doc_sii):
    switcher = {
        33: 'FAC',
        34: 'FEX',
        61: 'NC'
    }
    return switcher.get(tipo_doc_sii,"")

def check_exists_by_id(wdrive, idtext):
    try:
        wdrive.find_element(By.ID, idtext)
    except NoSuchElementException:
        return False
    return True


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


def check_exists_by_link(wdrive, linktext):
    try:
        wdrive.find_element(By.LINK_TEXT, linktext)
    except NoSuchElementException:
        return False
    return True


def delete_pdf_files(rutaarchivos):
    #files = [f for f in os.listdir(rutaarchivos) if os.path.isfile(f)]
    ruta_paso= rutaarchivos + '*.pdf'
    files= glob.glob(ruta_paso)
    for f in files:
        try:
            os.remove(f)
        except OSError as e:
            logger.critical(f"Error:{e.strerror}")


def mover_pdf_files(rutaorigen,carpeta_destino,subcarpeta, empresa, periodo_paso, nomArchivoDestino, rut_paso):
    carpeta_destino_paso= carpeta_destino+subcarpeta+'/'
    try:
        os.stat(carpeta_destino_paso)
    except:
        os.mkdir(carpeta_destino_paso)
    carpeta_destino_paso= carpeta_destino_paso+empresa+'/'
    try:
        os.stat(carpeta_destino_paso)
    except:
        os.mkdir(carpeta_destino_paso)
    carpeta_destino_paso= carpeta_destino_paso+periodo_paso+'/'
    try:
        os.stat(carpeta_destino_paso)
    except:
        os.mkdir(carpeta_destino_paso)
    #py_files = [f for f in os.listdir(rutaorigen) if os.path.isfile(f)]
    ruta_paso= rutaorigen + '*.pdf'
    py_files= glob.glob(ruta_paso)
    cant_copiado= 0
    for py_file in py_files:
        try:
            cant_copiado= cant_copiado + 1
            shutil.move(py_file, carpeta_destino_paso+nomArchivoDestino)
            #print(f"Destino:{carpeta_destino_paso}, Archivo:{nomArchivoDestino}")
            logger.info(f"Empresa:{empresa},Rut: {rut_paso}. Archivo:{nomArchivoDestino} copiado a carpeta destino ({carpeta_destino_paso}).")
        except OSError as e:
            logger.critical(f"Error:{e.strerror}")
    if (cant_copiado==0):
            logger.info(f"Empresa:{empresa},Rut: {rut_paso}. Archivo, desde {ruta_paso} a:{carpeta_destino_paso} con nombre {nomArchivoDestino} NO copiado!")
            

#def pdf_compras(lsrut, lsrut_pdf, lspassword_pdf):
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

chrome_options = Options()
chrome_options.add_experimental_option('prefs',  {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
    "download.default_directory": rutacarpetadescarga_pdf
    }
)
#    "plugins.plugins_disabled": ["Chrome PDF Viewer"], "download.default_directory": rutacarpetadescarga_pdf,

archivo_log= rutaregistro +'Log_PDF_' +ames + "_" +str(hoy.day) + "_" + str(hoy.hour) + "_" + str(hoy.minute)+".txt"
logger = logging.getLogger(archivo_log)
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(archivo_log)
fh.setLevel(logging.DEBUG)
# logger.addHandler(fh)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

rut_pdf=""
password_pdf=""

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
    cursor.execute("SELECT [nombreEmpresa_emp],[login_SII_emp], [password_SII_emp],[ruta_destino_SII_LC_emp], [ruta_destino_SII_LC_emp], [rut_sii_pdf_emp], [password_sii_pdf_emp],[ruta_destino_sii_pdf_emp] FROM [dbo].[EMPRESAS] WHERE [Descarga_Libros_SII_emp]='SI' AND [rut_sii_pdf_emp]<>'';")
    # Con fetchall traemos todas las filas
except Exception as e:
    error_string = str(e)
    logger.critical(f"Ocurrió un error al consultar: {error_string}")
    quit()

try:
    #with conexion.cursor() as cursor:
    empresas = cursor.fetchall()
    cantempresas_base= len(empresas)
    logger.info(f"PDF:Empresas a Procesar={cantempresas_base}.")
    # Recorrer y descargar
    for empresa in empresas:
        continuar= True
        cantempresas_descargadas= cantempresas_descargadas + 1
        nomEmpresa= empresa[0]
        rut= empresa[1]
        password_sii= empresa[2]
        carpeta_destino_lc= empresa[3]
        carpeta_destino_lv = empresa[4]
        rut_pdf= empresa[5]
        password_pdf=empresa[6]
        rutacarpetapdf= empresa[7]
        if carpeta_destino_lc[-1] != "\\":
            carpeta_destino_lc= carpeta_destino_lc + "\\"
        carpeta_destino_lc= carpeta_destino_lc + ames + "\\"
        if carpeta_destino_lv[-1] != "\\":
            carpeta_destino_lv= carpeta_destino_lv + "\\"
        if rutacarpetapdf[-1] != "\\":
            rutacarpetapdf= rutacarpetapdf + "\\"
        carpeta_destino_lv = carpeta_destino_lv + ames + "\\"
        carpeta_destino_lc.replace("\\", "/")
        carpeta_destino_lv.replace("\\", "/")
        rutacarpetapdf.replace("\\", "/")
        #print(f"carpeta_destino_lc= {carpeta_destino_lc}, carpeta_destino_lv={carpeta_destino_lv}, rutacarpetapdf={rutacarpetapdf}, rutacarpetadescarga_pdf={rutacarpetadescarga_pdf}")
        #quit()
        if rut_pdf=="":
            logger.info(f"Empresa {nomEmpresa} No tiene definido el rut del representante, NO se pueden descargar PDF Compra/Venta. Rut: {rut}")
        else:
            for operacion in [1,2,3]:
                try:
                    if (operacion ==1): 
                        nomArchivosCompra = "RCV_COMPRA_REGISTRO_" + rut + "_" + str(agno) + str(mes).zfill(2) + "*.csv"
                        #carpeta_destino_lc= carpeta_destino_lc + "COMPRA/"
                        ruta = carpeta_destino_lc + "COMPRA/" + nomArchivosCompra
                        nombreOperacion= "COMPRA"
                    elif (operacion ==2):
                        nomArchivosCompra = "RCV_COMPRA_PENDIENTE_" + rut + "_" + str(agno) + str(mes).zfill(2) + "*.csv"
                        #carpeta_destino_lc= carpeta_destino_lc + "COMPRA/"
                        ruta = carpeta_destino_lc + "COMPRA/" + nomArchivosCompra
                        nombreOperacion= "COMPRA"
                    else:
                        nomArchivosCompra = "RCV_VENTA_" + rut + "_" + str(agno) + str(mes).zfill(2) + "*.csv"
                        #carpeta_destino_lv= carpeta_destino_lv + "VENTA/"
                        ruta = carpeta_destino_lv + "VENTA/" + nomArchivosCompra
                        nombreOperacion= "VENTA"
                    nomArchivoCompra=""
                    py_files = glob.glob(ruta)
                    for py_file in py_files:
                        try:
                            #open(py_file)
                            nomArchivoCompra= py_file
                            #print(f"nomArchivoCompra: {nomArchivoCompra}")
                        except OSError as e:
                            logger.critical(f"Error:{e.strerror}")
                    if (nomArchivoCompra==""):
                        if (operacion ==1): 
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut},No se encontro archivo Libro Compra para procesar. Ruta:{ruta}")
                        elif (operacion ==2): 
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut},No se encontro archivo Libro Compra Pendiente para procesar. Ruta:{ruta}")
                        else:
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut},No se encontro archivo Libro Venta para procesar. Ruta:{ruta}")
                    else:
                        #Obtener lista de pdf. Para revisar y no descargar nuevamente.
                        ruta_paso= rutacarpetapdf + nombreOperacion + "/"+ nomEmpresa + "/"+ ames+"/*.pdf"                            
                        pdf_files= glob.glob(ruta_paso)
                        #print(f"ruta_paso:{ruta_paso}, Cant files:{str(len(pdf_files))}")
                        if (operacion ==1): 
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut}, Archivo Libro Compra:{nomArchivoCompra}")
                        elif (operacion ==2): 
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut}, Archivo Libro Compra Pendiente:{nomArchivoCompra}")
                        else:
                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut}, Archivo Libro Venta:{nomArchivoCompra}")
                        try:
                            continuar= True
                            driver2 = webdriver.Chrome('./chromedriver.exe',options = chrome_options)
                            driver2.set_page_load_timeout(100)
                            driver2.set_window_size(1280, 960)
                            driver2.get('https://homer.sii.cl/index.html')
                            time.sleep(2)
                            vinculo = driver2.find_element(By.XPATH, "//a[contains(text(),'Ingresar a Mi Sii')]")
                            vinculo.click()
                            time.sleep(2)

                            driver2.find_element_by_id('rutcntr').send_keys(rut_pdf)
                            driver2.find_element_by_id('clave').send_keys(password_pdf)
                            btn_ingresar = driver2.find_element(By.XPATH, "//button[@id='bt_ingresar']/img")
                            btn_ingresar.click()
                            time.sleep(2)

                            error_proceso= False
                            comando="//a[contains(text(),'Continuar')]"
                            representante= False
                            if check_exists_by_xpath(driver2, comando):
                                representante= True
                                ventana_aviso = driver2.find_element(By.XPATH, comando)
                                ventana_aviso.click()
                            time.sleep(2)
                            comando = "//div[@id='ModalEmergente']/div/div/div[3]/button"
                            if check_exists_by_xpath(driver2, comando):
                                ventana_aviso = driver2.find_element(By.XPATH, comando)
                                ventana_aviso.click()
                            time.sleep(6)

                            comando="//ul[@id='main-menu']/li[2]/a" 
                            #"//a[contains(text(),'Servicios online')]"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                            else:
                                continuar= False
                            #Factura Electronica
                            time.sleep(2)
                            comando= "Factura electrónica"
                            if continuar and check_exists_by_link(driver2, comando):
                                vinculo = driver2.find_element(By.LINK_TEXT, comando)
                                vinculo.click()
                            else:
                                continuar= False
                            time.sleep(2)
                            comando="//p[4]/a"
                            #//a[contains(text(),'Sistema de facturación gratuito del SII')]"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                            else:
                                continuar= False
                            time.sleep(2)
                            comando="//a[contains(.,'Historial de DTE y respuesta a documentos recibidos (*)')]"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                            else:
                                continuar= False
                            time.sleep(2)
                            if (operacion ==1) or (operacion ==2):
                                comando="//a[contains(.,'Ver documentos recibidos - Generar respuesta al emisor')]"
                            else:
                                comando="//a[contains(.,'Ver documentos emitidos')]"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                            else:
                                continuar= False
                            time.sleep(2)
                            comando="//select[@name='RUT_EMP']"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                rutempre = driver2.find_element(By.XPATH, comando)
                                rutempre.click()
                                time.sleep(1)
                                selrutempresa = Select(rutempre)
                                selrutempresa.select_by_value(rut)
                            else:
                                continuar= False
                            time.sleep(2)
                            comando="//button[contains(.,'Enviar')]"
                            if continuar and check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                            else:
                                continuar= False                    
                            time.sleep(2)
                            if continuar:
                                #--- Abrir y leer libro Compra/Venta, recorrer y descargar docs
                                df= pd.read_csv(nomArchivoCompra,sep=';',index_col = False)
                                #df.set_index('Nro')
                                for row in df.itertuples(index=False):
                                    continuar= True
                                    
                                    print(row)

                                    tipo_doc_sii= row[1]
                                    rut_prov= row[3]
                                    folio=row[5]
                                    x = rut_prov.split("-")
                                    rut_sd= x[0]
                                    nomArchivo=str(folio)+"_"+rut_sd+'_'+str(tipo_doc_sii) +'.pdf'
                                    for f_pdf in pdf_files:
                                        nomArchivo_paso= path.basename(f_pdf)
                                        if (nomArchivo_paso==nomArchivo):
                                            logger.info(f"Empresa:{nomEmpresa},Rut:{rut}, {nombreOperacion} Archivo:{nomArchivo} ya esta en la carpeta destino!")
                                            continuar=False
                                    #print(f"Rut Proveedor:{rut_prov}, Rut SD:{rut_sd}, Folio:{folio}")
                                    if continuar:
                                        delete_pdf_files(rutacarpetadescarga_pdf)
                                        continuar= True
                                        time.sleep(2)
                                        #comando="//a[contains(.,'SELECCIÓN DE DOCUMENTOS')]"
                                        comando="//a[contains(@href, '#collapseFiltro')]"
                                        if continuar and check_exists_by_xpath(driver2, comando):
                                            vinculo = driver2.find_element(By.XPATH, comando)
                                            vinculo.click()
                                        else:
                                            continuar= False
                                        time.sleep(1)
                                        if (operacion==1) or (operacion==2):
                                            comando="//input[@name='RUT_EMI']"
                                        else:
                                            comando="//input[@name='RUT_RECP']"
                                        if continuar and check_exists_by_xpath(driver2, comando):
                                            inputobject = driver2.find_element(By.XPATH, comando)
                                            inputobject.clear()
                                            inputobject.send_keys(rut_sd)
                                        else:
                                            continuar= False
                                        time.sleep(1)
                                        comando="//input[@name='FOLIO']"
                                        if continuar and check_exists_by_xpath(driver2, comando):
                                            inputobject = driver2.find_element(By.XPATH, comando)
                                            inputobject.clear()
                                            inputobject.send_keys(folio)
                                        else:
                                            continuar= False
                                        time.sleep(1)
                                        comando="//select[@name='TPO_DOC']"
                                        if continuar and check_exists_by_xpath(driver2, comando):
                                            tipodoc = driver2.find_element(By.XPATH, comando)
                                            tipodoc.click()
                                            time.sleep(1)
                                            seltipodoc = Select(tipodoc)
                                            seltipodoc.select_by_value(str(tipo_doc_sii))
                                        else:
                                            continuar= False
                                        time.sleep(1)
                                        comando="//input[@name='BTN_SUBMIT']"
                                        if continuar and check_exists_by_xpath(driver2, comando):
                                            vinculo = driver2.find_element(By.XPATH, comando)
                                            vinculo.click()
                                        else:
                                            continuar= False
                                        if (continuar==False):
                                            logger.critical(f"Empresa:{nomEmpresa},Rut:{rut}, {nombreOperacion} Error imprevisto de la pagina. No permite filtrar doc. Rut.:{rut_prov} Folio:{folio}!")
                                        else:
                                            continuar= True
                                            time.sleep(1)
                                            comando="//td/a/img"
                                            if continuar and check_exists_by_xpath(driver2, comando):
                                                vinculo = driver2.find_element(By.XPATH, comando)
                                                vinculo.click()
                                            else:
                                                continuar= False
                                            if (continuar==False):
                                                if (operacion==1) or (operacion==2):
                                                    logger.critical(f"Empresa:{nomEmpresa},Rut:{rut}, Prov.:{rut_prov} Folio:{folio} NO Encontrado!")
                                                else:
                                                    logger.critical(f"Empresa:{nomEmpresa},Rut:{rut}, Cliente.:{rut_prov} Folio:{folio} NO Encontrado!")
                                            else:
                                                if (operacion==1) or (operacion==2):
                                                    time.sleep(1)
                                                    comando="//a[contains(.,'Otros detalles documento')]"
                                                    if continuar and check_exists_by_xpath(driver2, comando):
                                                        vinculo = driver2.find_element(By.XPATH, comando)
                                                        vinculo.click()
                                                    else:
                                                        continuar= False
                                                    time.sleep(1)
                                                    comando="//a[contains(.,'VISUALIZACIÓN DOCUMENTO (pdf)')]"
                                                    if continuar and check_exists_by_xpath(driver2, comando):
                                                        vinculo = driver2.find_element(By.XPATH, comando)
                                                        vinculo.click()
                                                    else:
                                                        continuar= False
                                                    if continuar:
                                                        time.sleep(5)
                                                        #logger.info(f"Empresa:{rut}, Prov.:{rut_prov} Folio:{folio} Descargado!")                                    
                                                        mover_pdf_files(rutacarpetadescarga_pdf, rutacarpetapdf, nombreOperacion, nomEmpresa, ames, nomArchivo, rut)
                                                        #driver2.back() 

                                                    time.sleep(1)
                                                    comando="//input[@value='Volver a Revisión Documentos']"
                                                    if continuar and check_exists_by_xpath(driver2, comando):
                                                        vinculo = driver2.find_element(By.XPATH, comando)
                                                        vinculo.click()
                                                    else:
                                                        continuar= False
                                                    time.sleep(1)
                                                else:
                                                    #<button id="open-button" tabindex="1">Abrir</button>
                                                    time.sleep(2)
                                                    try:
                                                        iframe = driver2.find_elements_by_tag_name('iframe')[0]
                                                        driver2.switch_to.frame(iframe)
                                                        comando="//button[text()='Abrir']"
                                                        if continuar and check_exists_by_xpath(driver2, comando):
                                                            vinculo = driver2.find_element(By.XPATH, comando)
                                                            vinculo.click()
                                                        else:
                                                            continuar= False
                                                    finally:
                                                        driver2.switch_to.default_content()
                                                    
                                                    if continuar:
                                                        time.sleep(5)
                                                        mover_pdf_files(rutacarpetadescarga_pdf, rutacarpetapdf, nombreOperacion, nomEmpresa, ames, nomArchivo, rut)
                                                    time.sleep(1)
                                                    comando="//input[@value='Volver']"
                                                    if continuar and check_exists_by_xpath(driver2, comando):
                                                        vinculo = driver2.find_element(By.XPATH, comando)
                                                        vinculo.click()
                                                    else:
                                                        continuar= False
                                                    time.sleep(1)
                            else:
                                logger.critical(f"Empresa:{nomEmpresa},Rut:{rut}, error desconocido impidio revisar/descargar documentos {nombreOperacion}.")
                            time.sleep(2)
                            comando="//a[contains(.,'Cerrar Sesión')]"
                            if check_exists_by_xpath(driver2, comando):
                                vinculo = driver2.find_element(By.XPATH, comando)
                                vinculo.click()
                        except Exception as e:
                            error_string = str(e)
                            if (operacion ==1):
                                logger.critical(f"Empresa:{nomEmpresa},Rut:{rut} ,Ocurrió un error en descarga pdf Compra:{error_string}") 
                            elif (operacion ==2):
                                logger.critical(f"Empresa:{nomEmpresa},Rut:{rut} ,Ocurrió un error en descarga pdf Compra Pendiente:{error_string}") 
                            else:
                                logger.critical(f"Empresa:{nomEmpresa},Rut:{rut} ,Ocurrió un error en descarga pdf Venta:{error_string}") 

                        finally:
                            driver2.close()
                            driver2.quit()
                            time.sleep(2)
                except Exception as e:
                    error_string = str(e)
                    logger.critical(f"Empresa:{nomEmpresa},Rut:{rut} ,Ocurrió un error en descarga pdf Compra/Venta:{error_string}")
except Exception as e:
    error_string = str(e)
    logger.critical(f"Empresa:{nomEmpresa},Rut:{rut} ,Ocurrió un error:{error_string}")
finally:
    conexion.close()
    logger.info(f"Proceso terminado! {cantempresas_descargadas} Empresas Descargadas de {cantempresas_base}.")
    logging.shutdown()

    if (sendmail_logpdf=="YES") or (sendmail_logpdf=="SI"):
        # Creamos el objeto mensaje
        mensaje = MIMEMultipart()
        
        # Establecemos los atributos del mensaje
        asunto="Descarga PDF Libros SII concluido, " + hoy.strftime("%d/%m/%Y, %H:%M:%S")
        cuerpo = "Se adjunta archivo log para su revision. Revise las linea con aviso 'Critical'."
        mensaje['From'] = usermail_logpdf
        mensaje['To'] = destinationmail_logpdf
        mensaje['Subject'] = asunto
        # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        
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
        # Ciframos la conexión
        sesion_smtp.starttls()
        # Iniciamos sesión en el servidor
        sesion_smtp.login(usermail_logpdf,passwordmail_logpdf)
        # Convertimos el objeto mensaje a texto
        texto = mensaje.as_string()
        # Enviamos el mensaje
        sesion_smtp.sendmail(usermail_logpdf, destinationmail_logpdf, texto)
        # Cerramos la conexión
        sesion_smtp.quit()

