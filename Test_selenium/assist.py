import logging
import os
import selenium
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
from tenacity import retry, retry_if_result


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


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)


ag.PAUSE = 2

fecha_actual = datetime.today().strftime('%Y-%m-%d')
dia = datetime.now().strftime('%d-%m-%Y')

directorio_descargas = r'W:\Test_selenium\Descarga'

archivos_xlsx_registro = r'W:\Test_selenium\Descarga\Registro_compra'
archivos_xlsx_pendiente = r'W:\Test_selenium\Descarga\Registro_pendientes'

direccion1 = 'maximiano.coronel@ecm.cl'
direccion2 = 'alberto.allendes@ecm.cl'
ruta = r'W:\Test_selenium\Historico'
nombre_archivo = 'Nuevos Registro Compras-' + str(dia) +'.xlsx'
asunto = 'Nuevos Registros Compras ' + str(dia)


logging.basicConfig(filename='Log_SII_' + fecha_actual + '.txt',
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

texto = 'Estimados:\n\nSe adjunta registro de compras del holding para el mes ' + meses[currentMonth - 1] + ' y los nuevos registros al: ' + fecha_actual


def error500():
    """Devuelve True si el valor es Error500 """
    return "Error500" in driver.page_source



@timing
def main():
    for i in empresas:
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_anterior))
        descargarBalance(i[1], mes_anterior, i[2])
        logging.info('Comienza a descargar Empresa: {} , Mes: {}'.format(i[0], mes_actual))
        descargarBalance(i[1], mes_actual, i[2])

    #             # except WebDriverException as e:
    #             #     logging.info('Error por excepción de WebDriver', e)
    #             #     continue
    driver.close()
    logging.info('Fin descargas, se cambian de formato los archivos.')
    genLibro()
    # logging.info('Fin descargas. Se abre Macro')
    # correrMacro()
    # logging.info('Fin Macro. Se envia Correo')
    # enviarCorreo(direccion1, ruta, nombre_archivo, asunto)
    # direccion2,



# # def descargarBalance(nombre, passwd):
options = Options()
#options.headless = True
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
