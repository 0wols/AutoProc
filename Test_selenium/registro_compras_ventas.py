"""

Programa de descarga de Registro Ventas SII

VERSION CON POP-UP

"""

import logging
import os
import selenium
import shutil
import pyautogui as ag
import pandas as pd
from glob import glob
from os import chdir
from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import NoSuchElementException
from secrets import empresas, direccion_para, direccion_cc
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
from datetime import datetime
from functools import wraps
from time import time, sleep
from tenacity import retry, retry_if_result, retry_if_exception_type
from win32api import GetKeyState 
from win32con import VK_CAPITAL


ag.PAUSE = 2

fecha_actual = datetime.today().strftime('%Y-%m-%d')
dia = datetime.now().strftime('%d-%m-%Y')
dia_formateado = int(dia[:2])
fecha_csv = datetime.today().strftime('%Y%m')



root = "C:\\Users\\Usuario ECM\\Desktop\\Python\\AutoProc\\Test_selenium"
rutas = "C:\\Users\\Usuario ECM\\Desktop\\Python\\AutoProc\\Test_selenium\\Historico1\\"


nombres_archivo = ('Registro Compras Resumen-' + str(dia) +'.xlsm', 'Registro Ventas Resumen-' + str(dia) + '.xlsm') 
asunto = 'Registros Compras y Ventas ' + str(dia)


logging.basicConfig(filename=r"W:\Logs\Test_selenium\Log_SII_" + fecha_actual + '.txt',
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.INFO)

logging.info('Programa iniciado, Selenium version: %s' % selenium.__version__)
currentMonth = datetime.now().month

meses = {
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

mes_actual = meses[currentMonth]
mes_anterior = meses[currentMonth - 1]

texto = 'Estimad@s:\n\nSe adjunta registro de compras y registro de ventas para el holding actualizado al: ' + fecha_actual +'\n\n'


def timing(f):
    @wraps(f)
    def wrap(*args, **kw):
        ts = time()
        result = f(*args, **kw)
        te = time()
        print('func:%r args:[%r, %r] took: %2.4f sec' % \
          (f.__name__, args, kw, te-ts))
        return result
    return wrap


@timing
def main():
    if dia_formateado < 10:
        loop1()
    else:
        loop2()
    driver.close()
    logging.info('Fin descargas, se cambian de formato los archivos.')
    generar_libro()
    logging.info('Fin conversion. Se mueven archivos xlsx')
    mover_xlsx()
    logging.info('Fin movimiento archivos. Se ejecuta Macro')
    correr_macro_compras()
    correr_macro_ventas()
    logging.info('Fin Macro. Se envia Correo')
    enviar_correo(direccion_para, direccion_cc, rutas, nombres_archivo, asunto)
    logging.info('Programa Finalizado')


def loop1():
    for i in empresas:
        logging.info('Comienza a descargar Registro Compras Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_rc_y_pend(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Registro Compras Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_rc_y_pend(i[1], mes_actual, i[2])
        logging.info('Comienza a descargar Registro Ventas Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_ventas(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Registro Ventas Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_ventas(i[1], mes_actual, i[2])




def loop2():
    for i in empresas:
        logging.info('Comienza a descargar Registro Compras Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_rc(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Registro Compras Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_rc_y_pend(i[1], mes_actual, i[2])
        logging.info('Comienza a descargar Registro Ventas Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_ventas(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Registro Ventas Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_ventas(i[1], mes_actual, i[2])


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)


options = Options()
# options.headless = True
prefs = {
    "download.default_directory" : root + "\\Descarga",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing_for_trusted_sources_enabled": False,
    "safebrowsing.enabled": False
}
options.add_experimental_option("prefs", prefs)

options.add_argument("--window-size=1920,1200")

dir = 'https://zeusr.sii.cl/AUT2000/InicioAutenticacion/IngresoRutClave.html?https://www4.sii.cl/consdcvinternetui/'

DRIVER_PATH = r'C:\Users\Usuario ECM\Downloads\chromedriver_win32\chromedriver.exe'
ser = Service(DRIVER_PATH)
driver = webdriver.Chrome(
    options=options,
    service=ser)
action = ActionChains(driver)
driver.implicitly_wait(20)

driver.get(dir)
driver.maximize_window()

wait = WebDriverWait(driver, 20)


# @retry(retry=retry_if_result(lambda result : result == "Error 500") | retry_if_result(lambda result : result == "ERR_CONNECTION_RESET") | retry_if_exception_type(WebDriverException) | retry_if_exception_type(TimeoutException))
def descargar_rc_y_pend(nombre, mes, passwd):
    """ Descargar archivos csv de pagina del SII """

    logear(nombre, mes, passwd)

    if "No hay información de Registro." in driver.page_source:
        sleep(3)
        element7 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//strong[contains(text(),'Pendientes')]")))
        element7.click()
        if "Sin información" in driver.page_source:
            logging.info("No hay información de Registro ni de Pendientes. No se descargaron archivos")
            deslogear()
        elif "No hay información de Pendientes" in driver.page_source:
            logging.info("No hay información de Registro ni de Pendientes. No se descargaron archivos")
            sleep(3)
            deslogear()
        else:
            element8 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
            element8.click()
            sleep(3)
            logging.info('Archivo descargado')
            deslogear()
    elif "Sin información" in driver.page_source:
        element7 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//strong[contains(text(),'Pendientes')]")))
        element7.click()
        sleep(5)
        if "Sin información" in driver.page_source:
            logging.info("No hay información de Registro ni de Pendientes. No se descargaron archivos")
            deslogear()
        else:
            element8 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
            element8.click()
            sleep(3)
            logging.info('Archivo descargado')
            deslogear()
    else:
        element5 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
        action.move_to_element(element5).perform()
        action.double_click(element5).perform()
        sleep(3)
        element6 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//strong[contains(text(),'Pendientes')]")))
        element6.click()
        sleep(3)
        if "Sin información" in driver.page_source:
            sleep(3)
            logging.info('Archivo descargado')
            deslogear()
        elif "No hay información de Pendientes" in driver.page_source:
            sleep(3)
            logging.info('Archivo descargado')
            deslogear()
        else:
            element7 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
            element7.click()
            sleep(3)
            logging.info('Archivo descargado')
            deslogear()


def descargar_rc(nombre, mes, passwd):
    """ Descargar archivos csv de pagina del SII """

    logear(nombre, mes, passwd)

    if "No hay información de Registro." in driver.page_source:
        sleep(3)
        deslogear()
    elif "Sin información" in driver.page_source:
        sleep(3)
        deslogear()
    else:
        element5 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
        action.move_to_element(element5).perform()
        action.double_click(element5).perform()
        sleep(3)
        deslogear()


def descargar_ventas(nombre, mes, passwd):
    logear(nombre, mes, passwd)

    element5 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//strong[contains(text(),'VENTA')]")))
    element5.click()
    sleep(5)
    sinDoc = ag.locateCenterOnScreen(imPath("sin_documentos.png"))

    if "No hay información de Ventas." in driver.page_source:
        sleep(3)
        logging.info("Sin información de registro de ventas")
        deslogear()
    elif sinDoc is not None:
        sleep(3)
        logging.info("Sin información de registro de ventas")
        deslogear()
    else:
        sleep(3)
        element6 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
        # element6.click()
        action.move_to_element(element6).perform()
        action.double_click(element6).perform()
        sleep(3)
        deslogear()


def logear(nombre, mes, passwd):
    element0 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//input[@id='rutcntr']")))
    element0.send_keys(nombre)

    element1 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//input[@id='clave']")))
    element1.send_keys(passwd)

    element2 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[@id='bt_ingresar']")))
    element2.click()
    sleep(3)

    while True:
        try:
            element3 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//select[@id='periodoMes']")))
            select = Select(element3)
            select.select_by_visible_text(mes)
        except TimeoutException as t:
            logging.info('Error por excepción de Timeout', t)
            continue
        break
    sleep(3)

    element4 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Consultar')]")))
    element4.click()
    sleep(3)


def deslogear():
    element9 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//body/div[@id='my-wrapper']/div[1]/div[1]/nav[1]/div[1]/div[1]/span[1]/a[1]")))
    element9.click()
    while True:
        link = ag.locateCenterOnScreen(imPath("link_registro.png"))
        if link is not None:
            break
    ag.click(link, clicks=2)
    element12 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//a[contains(text(),'Ingresar al Registro de Compras y Ventas')]")))
    element12.click()
    sleep(3)
    pass

custom_date_parser = lambda x: datetime.strptime(x, "%d-%m-%Y")


def generar_libro():
    """ Generar archivos xlsx a partir de los archivos csv descargados """
    for csv_file in glob(root + "\\Descarga\\*.csv"):
        print(csv_file)
        df = pd.read_csv(csv_file, sep=";", parse_dates=['Fecha Docto','Fecha Recepcion'], infer_datetime_format=True, dayfirst=True, index_col=False)
        xlsx_file = os.path.splitext(csv_file)[0] + '.xlsx'
        df.to_excel(xlsx_file, index=None, header=True)


def mover_xlsx():
    """ Mover de carpetas los archivos xlsx"""
    source = root + "\\Descarga\\"
    dest1 = root + "\\Descarga\\Registro_compra\\"
    dest2 = root + "\\Descarga\\Registro_pendientes\\"
    dest3 = root + "\\Descarga\\Registro_ventas\\"

    chdir(source)

    files = os.listdir(source)

    for f in glob("*.xlsx"):
        if f.startswith("RCV_COMPRA_REGISTRO"):
            shutil.move(f, dest1)
        elif f.startswith("RCV_COMPRA_PENDIENTE"):
            shutil.move(f, dest2)
        elif f.startswith("RCV_VENTA"):
            shutil.move(f, dest3)


def correr_macro_compras():
    """ Ejecutar archivo xlsm con la macro del Registro Compras  """
    ag.hotkey("winleft", "r")
    caps1 = GetKeyState(VK_CAPITAL)
    if caps1 == 0:
        ag.write(root +  "\\Registro Compras.xlsm")
    else:
        ag.write("c:\\uSERS\\uSUARIO ecm\\dESKTOP\\pYTHON\\aUTOpROC\\tEST_SELENIUM\\rEGISTRO cOMPRAS.XLSM")
    ag.press("enter")
    sleep(5)
    ag.hotkey("winleft", "up")
    ag.click(862, 75)
    ag.click(469, 154)
    ag.hotkey("ctrl", "u")

    sleep(8)

    while True:
        c = ag.locateCenterOnScreen(imPath('Mikasa.png'))
        if c is not None:
            ag.press("enter")
            break


def correr_macro_ventas():
    """ Ejecutar archivo xlsm con la macro del Registro Ventas  """
    ag.hotkey("winleft", "r")
    caps1 = GetKeyState(VK_CAPITAL)
    if caps1 == 0:
        ag.write(root +  "\\Registro Ventas.xlsm")
    else:
        ag.write("c:\\uSERS\\uSUARIO ecm\\dESKTOP\\pYTHON\\aUTOpROC\\tEST_SELENIUM\\rEGISTRO vENTAS.XLSM")
    ag.press("enter")
    sleep(5)
    ag.hotkey("winleft", "up")
    ag.click(862, 75)
    ag.click(469, 154)
    ag.hotkey("ctrl", "h")

    sleep(8)

    while True:
        c = ag.locateCenterOnScreen(imPath('Mikasa.png'))
        if c is not None:
            ag.press("enter")
            break


def enviar_correo(direccion_para, direccion_cc, rutas, nombres_archivo, asunto):
    """ Enviar correo de archivo consolidado """

    hora = datetime.now().strftime('%H:%M:%S')
    sleep(3)
    mouse.click(button='right', coords=(275, 1057))
    sleep(3)
    mouse.click(button='left', coords=(250, 850))
    outlook = Application().connect(
        best_match=u"Sin título - Mensaje (HTML)",
        timeout=200)
    outlook[u"Sin título - Mensaje (HTML)"].child_window(
        control_id=4099
    ).click_input()
    send_keys(direccion_para)
    outlook[u"Sin título - Mensaje (HTML)"].child_window(
        control_id=4100
    ).click_input()
    send_keys(direccion_cc)
    send_keys('{ENTER}')
    send_keys('%N')
    send_keys('ud')
    sleep(3)
    outlook[u"Insertar archivo"].child_window(control_id=1001).click_input()
    while True:
        file1 = ag.locateCenterOnScreen(imPath("historico.png"))
        if file1 is not None:
            break
    ag.click(file1, clicks=2)
    outlook[u"Insertar archivo"].child_window(
        control_id=1148,
        found_index=0).click_input()
    send_keys(nombres_archivo[0], with_spaces=True)
    send_keys('{ENTER}')
    sleep(3)
    send_keys('%N')
    send_keys('ud')
    outlook[u"Insertar archivo"].child_window(control_id=1001).click_input()
    while True:
        file2 = ag.locateCenterOnScreen(imPath("historico.png"))
        if file2 is not None:
            break
    ag.click(file2, clicks=2)
    outlook[u"Insertar archivo"].child_window(
        control_id=1148,
        found_index=0).click_input()
    send_keys(nombres_archivo[1], with_spaces=True)
    send_keys('{ENTER}')
    outlook[u"Sin título - Mensaje (HTML)"].child_window(
        control_id=4101).click_input()
    send_keys(asunto, with_spaces=True)
    outlook[u"Sin título - Mensaje (HTML)"].child_window(
        control_id=4159).click_input()
    ag.click(47, 338, duration=1)
    send_keys(
        texto,
        with_spaces=True,
        with_newlines=True)
    outlook[asunto + " - Mensaje(HTML)"].child_window(
        control_id=4256).click_input()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info('Se interrumpió el programa debido a KeyboardInterrupt')
        pass
