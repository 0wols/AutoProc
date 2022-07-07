"""

Programa de descarga de Boletas Recibidas SII

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

root = "C:\\Users\\Usuario ECM\\Desktop\\Python\\AutoProc\\Test_selenium"

logging.basicConfig(filename=r"W:\Logs\Test_selenium\Log_Boletas_" + fecha_actual + '.txt',
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.INFO)

logging.info('Programa iniciado, Selenium version: %s' % selenium.__version__)

mesesMayus = { 0 : "ENERO",
               1 : "FEBRERO",
               2 : "MARZO",
               3 : "ABRIL",
               4 : "MAYO",
               5 : "JUNIO",
               6 : "JULIO",
               7 : "AGOSTO",
               8 : "SEPTIEMBRE",
               9 : "OCTUBRE",
               10 : "NOVIEMBRE",
               11 : "DICIEMBRE"
}


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
    logging.info("Comienza Loop descarga Boletas recibidas")
    loop()
    driver.close()
    logging.info("Programa Finalizado")


def loop():
    for empresa in empresas:
        descargar_boletas_recibidas(empresa[0], empresa[1], empresa[2])


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)


options = Options()
# options.headless = True
prefs = {
    "download.default_directory" : "C:\\Users\\Usuario ECM\\Desktop\\Python\\AutoProc\\Test_selenium\\Descarga\\Impuestos\\",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing_for_trusted_sources_enabled": False,
    "safebrowsing.enabled": False
}
options.add_experimental_option("prefs", prefs)

options.add_argument("--window-size=960,1200")

dir = 'https://zeusr.sii.cl/AUT2000/InicioAutenticacion/IngresoRutClave.html?https://loa.sii.cl/cgi_IMT/TMBCOC_MenuConsultasContribRec.cgi?dummy=1461943244650'

DRIVER_PATH = r'C:\Users\Usuario ECM\Downloads\chromedriver_win32\chromedriver.exe'
ser = Service(DRIVER_PATH)
driver = webdriver.Chrome(
    options=options,
    service=ser)
action = ActionChains(driver)
driver.implicitly_wait(20)

driver.get(dir)
# driver.maximize_window()

wait = WebDriverWait(driver, 20)


def descargar_boletas_recibidas(nombre, crut, passwd):

    logear(crut, passwd)

    element3 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//input[@id='cmdconsultar124']")))
    element3.click()

    sleep(3)

    mesDatos0 = []

    for mes, valor in mesesMayus.items():
        try:
            linkMes = driver.find_element(By.XPATH, r"//a[contains(text(),'{}')]".format(valor))
            mesDatos0.append(linkMes)
        except NoSuchElementException:

            logging.info("No hay información de Boletas recibidas para la empresa:{} en el mes:{}".format(nombre, valor))
            # mesDatos0.append(None)
            pass

    print(mesDatos0)

    mesDatos1 = dict(zip(mesesMayus.values(), mesDatos0))

    # mesDatos1 = { k : v for (k, v) in mesDatos1.items() if v != None }

    print(mesDatos1)

    for mes, valor in mesDatos1.items():

        sleep(5)

        valor.click()

        sleep(3)
        
        tab = driver.find_element(By.XPATH,r"//body/div[2]/center[1]/center[1]/form[1]/table[1]")
        tab_html = tab.get_attribute('outerHTML')
        df = pd.read_html(tab_html, skiprows=1, thousands='.', encoding="utf-8")[0]
        df.to_excel(r"C:\Users\Usuario ECM\Desktop\Python\AutoProc\Test_selenium\Descarga\Boletas_Recibidas\Boletas_Recibidas_{}_{}.xlsx".format(nombre, mes), header=False, index=False, float_format = '%.2f')

        sleep(2)

        volver = wait.until(EC.element_to_be_clickable((By.XPATH, r"//tbody/tr[1]/td[3]/input[1]")))
        volver.click()
        logging.info("Descargadas Boletas Mes:{} Empresa:{}".format(mes, nombre))


    deslogear()


    
def logear(crut, passwd):

    element0 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//input[@id='rutcntr']")))
    element0.send_keys(crut)

    element1 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//input[@id='clave']")))
    element1.send_keys(passwd)

    element2 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[@id='bt_ingresar']")))
    element2.click()

    sleep(3)


def deslogear():

    element0 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//a[contains(text(),'Cerrar Sesión')]")))
    element0.click()

    element1 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//a[contains(text(),'Emitir boleta de honorarios electrónica')]")))
    element1.click()

    element2 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//button[contains(text(),'Cerrar')]")))
    element2.click()

    sleep(3)

    while True:
        botonConsultar = ag.locateCenterOnScreen(imPath("Consultar_Boletas.png"))
        if botonConsultar is not None:
            break

    ag.click(botonConsultar) 

    element3 = wait.until(EC.element_to_be_clickable((By.XPATH, r"//a[contains(text(),'Consultar boletas recibidas')]")))
    element3.click()



if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info('Se interrumpió el programa debido a KeyboardInterrupt')
        pass
