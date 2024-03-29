"""

Programa de descarga de Registro Compras SII

VERSION MAIN ( NO POP-UP )

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

direccion_para = "fernando.allendes@ecm.cl;jaime.arancibia@ecm.cl;cristian.coronel@ecm.cl"
direccion_cc = "alberto.allendes@ecm.cl;maximiano.coronel@ecm.cl;pablo.coronel@valsegur.cl;lorena.paredes@ecm.cl;tomas.yanez@ecm.cl"

rutas = (r'W:\Test_selenium\Historico', r'W:\Test_selenium\Descarga\Registro_ventas')
nombres_archivo = ('Registro Compras Resumen-' + str(dia) +'.xlsm', 'RCV_VENTA_89630400-3_' + str(fecha_csv) + '.csv') 
asunto = 'Registros Compras ' + str(dia)


logging.basicConfig(filename=r"C:\Users\Usuario ECM\Desktop\Python\Logs\Test_selenium\Log_SII_" + fecha_actual + '.txt',
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.INFO)

logging.info('Programa iniciado, Selenium version: %s' % selenium.__version__)


empresas = [
    ('ACHEGEO', '650210794', 'cordoba20'),
    ('APOCE', '651465125', 'huesca20'),
    ('VALPARAISO', '760094056', 'zaragoza20'),
    ('IQUIQUE', '760484032', 'bilbao20'),
    ('RECOLETA', '760888540', 'alicante20'),
    ('SANTIAGO', '761124684', 'toledo20'),
    ('CONCESIONES PROVIDENCIA', '761487582', 'sevilla20'),
    ('IMAGENOLOGIA', '761531697', 'caceres20'),
    ('LOS ANDES', '761541498', 'granada20'),
    ('COQUIMBO', '76161374k', 'almeria20'),
    ('ECM GEOTERMIA', '761684574', 'pamplona20'),
    ('CP LATINA CHILE', '761805495', 'salamanca2'),
    ('INFINERGEO', '761698176', 'murcia20'),
    ('PIEDMONTT', '761872265', 'cordoba20'),
    ('CONCESIONES VALDIVIA', '762136775', 'malaga20'),
    ('INVERREST', '763098923', 'alcala20'),
    ('CAJA EXPRESS', '763381463', 'albacete20'),
    ('VALSEGUR', '765033810', 'tarragona2'),
    ('RENAL', '765184304', 'daroca20'),
    ('POLLO ABRASADO', '767860056', 'coruna20'),
    ('CAFE DEL NORTE', '767860099', 'escalona20'),
    ('TALCA', '768298602', 'cuenca20'),
    ('CLEVERPARK', '769756159', 'aranjuez20'),
    ('PUNTA ARENAS', '772903170', 'almeria10'),
    ('DIAGNOSIS', '785037707', 'mallorca20'),
    ('MEDCONSUL', '785750306', 'altamira20'),
    ('INGENIEROS', '79558840k', 'burgos20'),
    ('SERVILAND', '798929402', 'merida20'),
    ('DENSITOSEA', '799540401', 'escorial20'),
    ('ECM', '896304003', 'madrid20'),
    ('INVERTERRA', '967010006', 'segovia20'),
    ('TOMOIMAGEN', '967529400', 'girona20'),
    ('SEC', '967786500', 'valencia20'),
    ('ESTACIONAR', '968709100', 'barcelona2'),
    ('AUTOPARQUEO', '968709208', 'huelva20'),
    ('AUTOORDEN', '968709305', 'getafe20')
]

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

texto = 'Estimad@s:\n\nSe adjunta registro de compras para el holding y registro de ventas ECM actualizado al: ' + fecha_actual +'\n\n'


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
    descargarVentasECM('896304003', mes_actual, 'madrid20')
    driver.close()
    logging.info('Fin descargas, se cambian de formato los archivos.')
    generar_libro()
    logging.info('Fin conversion. Se mueven archivos xlsx')
    mover_xlsx()
    logging.info('Fin movimiento archivos. Se ejecuta Macro')
    correr_macro()
    logging.info('Fin Macro. Se envia Correo')
    enviar_correo(direccion_para, direccion_cc, rutas, nombres_archivo, asunto)
    logging.info('Programa Finalizado')


def loop1():
    for i in empresas:
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_rc_y_pend(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_rc_y_pend(i[1], mes_actual, i[2])


def loop2():
    for i in empresas:
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargar_rc(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargar_rc_y_pend(i[1], mes_actual, i[2])


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)


options = Options()
# options.headless = True
prefs = {
    "download.default_directory" : r"W:\Test_selenium\Descarga",
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


def deslogear():
    element9 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//body/div[@id='my-wrapper']/div[1]/div[1]/nav[1]/div[1]/div[1]/span[1]/a[1]")))
    element9.click()
    element10 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//body/div[@id='my-wrapper']/div[3]/div[1]/div[1]/div[2]/div[1]/ul[1]/li[6]/a[1]")))
    element10.click()
    element11 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//a[contains(text(),'Ingresar al Registro de Compras y Ventas')]")))
    element11.click()
    sleep(3)
    pass


def descargarVentasECM(nombre, mes, passwd):
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

    element5 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//strong[contains(text(),'VENTA')]")))
    element5.click()

    sleep(3)
    element6 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Descargar Detalles')]")))
    action.move_to_element(element6).perform()
    action.click(element6).perform()
    sleep(5)


custom_date_parser = lambda x: datetime.strptime(x, "%d-%m-%Y")


def generar_libro():
    """ Generar archivos xlsx a partir de los archivos csv descargados """
    for csv_file in glob(r"W:\Test_selenium\Descarga\*.csv"):
        print(csv_file)
        df = pd.read_csv(csv_file, sep=";", parse_dates=['Fecha Docto','Fecha Recepcion'], infer_datetime_format=True, dayfirst=True, index_col=False)
        xlsx_file = os.path.splitext(csv_file)[0] + '.xlsx'
        df.to_excel(xlsx_file, index=None, header=True)


def mover_xlsx():
    """ Mover de carpetas los archivos xlsx"""
    source = 'W:\\Test_selenium\\Descarga\\'
    dest1 = 'W:\\Test_selenium\\Descarga\\Registro_compra\\'
    dest2 = 'W:\\Test_selenium\\Descarga\\Registro_pendientes\\'
    dest3 = 'W:\\Test_selenium\\Descarga\\Registro_ventas\\'

    chdir(source)

    files = os.listdir(source)

    for f in glob("*.xlsx"):
        if f.startswith("RCV_COMPRA_REGISTRO"):
            shutil.move(f, dest1)
        elif f.startswith("RCV_COMPRA_PENDIENTE"):
            shutil.move(f, dest2)
    for b in glob("*.csv"):
        if b.startswith("RCV_VENTA"):
            shutil.move(b, dest3)


def correr_macro():
    """ Ejecutar archivo xlsm con la macro  """
    ag.hotkey("winleft", "r")
    caps1 = GetKeyState(VK_CAPITAL)
    if caps1 == 0:
        ag.write(r"W:\Test_selenium\Registro Compras.xlsm")
    else:
        ag.write(r"w:\tEST_SELENIUM\rEGISTRO cOMPRAS.XLSM")
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
    send_keys(rutas[0])
    send_keys('{ENTER}')
    outlook[u"Insertar archivo"].child_window(
        control_id=1148,
        found_index=0).click_input()
    send_keys(nombres_archivo[0], with_spaces=True)
    send_keys('{ENTER}')
    sleep(3)
    send_keys('%N')
    send_keys('ud')
    outlook[u"Insertar archivo"].child_window(control_id=1001).click_input()
    send_keys(rutas[1])
    send_keys('{ENTER}')
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
