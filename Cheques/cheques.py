"""

Programa de descarga de Cheques ECM


"""

import pyautogui as ag
import time, os, logging
import ctypes
import subprocess
import re    
import pandas as pd
from datetime import date, timedelta
from pyautogui import FailSafeException, ImageNotFoundException
from pywinauto.application import Application
from tenacity import retry, retry_if_result, retry_if_exception_type, stop_after_delay
from win32api import GetKeyState 
from win32con import VK_CAPITAL 
from pyscreeze import Box



# VARIABLES GLOBALES

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
ag.PAUSE = 2

hoy = date.today()
hoy_formato_dia = date.today().strftime("%a")

logging.basicConfig(
    filename=r"W:\Logs\Cheques\Log_Cheques_" + date.today().strftime("%d-%m-%Y") + '.txt',
    level=logging.INFO,
    format='%(asctime)s.%(msecs)03d: %(message)s',
    datefmt='%H:%M:%S')

if hoy_formato_dia == 'Mon':
    fechaFlex = hoy - timedelta(days=3)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")
elif hoy_formato_dia == 'Sun':
    fechaFlex = hoy - timedelta(days=2)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")
else:
    fechaFlex = hoy - timedelta(days=1)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")

fecha = date.today().strftime("%d-%m-%Y")

empresas = [
    ('ACHEGEO', '1102002', 'ACHEGEO.png'),
    ('ANDES', '1102002', 'ANDES.png'),
    ('APOCE', '1102002', 'APOCE.png'),
    ('AUTOORDEN', '1102008', 'AUTOORDEN.png'),
    ('CAFETAL', '1102001', 'CAFETAL.png'),
    ('CLEVER PARK SPA', '1102002', 'CLEVER PARK SPA.png'),
    ('CONCESIONES PROVIDENCIA', '5111042', 'CONCESIONES PROVIDENCIA.png'),
    ('CONCESIONES PUNTA ARENAS', '1102004', 'CONCESIONES PUNTA ARENAS.png'),
    ('CONCESIONES SANTIAGO', '1102003', 'CONCESIONES SANTIAGO.png'),
    ('CONCESIONES RECOLETA', '1102002', 'CONCESIONES RECOLETA.png'),
    ('CONSORCIO VALPARAISO', '1102004', 'CONSORCIO VALPARAISO.png'),
    ('COQUIMBO', '1102001', 'COQUIMBO.png'),
    ('DENSITOSEA', '1102004', 'DENSITOSEA.png'),
    ('ESTACIONAR', '1102003', 'ESTACIONAR.png'),
    ('INFINERGEO', '1110132', 'INFINERGEO.png'),
    ('INVERREST', '1102008', 'INVERREST.png'),
    ('INVERTERRA', '1102004', 'INVERTERRA.png'),
    ('IQUIQUE', '1102004', 'IQUIQUE.png'),
    ('PIEDMONT', '1102005', 'PIEDMONT.png'),
    ('POLLO ABRAZADO', '1102001', 'POLLO ABRAZADO.png'),
    ('RENAL', '1102003', 'RENAL.png'),
    ('SEC', '1102004', 'SEC.png'),
    ('SERVILAND', '1102008', 'SERVILAND.png'),
    ('TALCA', '1102002', 'TALCA.png'),
    ('TOMO IMAGEN', '1102005', 'TOMO IMAGEN.png'),
    ('VALDIVIA', '1102001', 'VALDIVIA.png'),
    ('VALSEGUR', '1102005', 'VALSEGUR.png'),
    ('VALSEGUR-ETV', '1102005', 'VALSEGUR-ETV.png')
]


def main():


    logging.info('Programa iniciado. Presione Ctrl-C para abortar.')

    miDicc = dict()

    for i in empresas:
        miDicc[i[0]] = buscarVentana(i[2], i[1], fechaFlex, fecha, i[0])

    valorTest = Box(left=737, top=522, width=357, height=39) 

    keyList = [ k for k in miDicc.keys() if miDicc[k] == valorTest ]

    logging.info("Programa Finalizado. Empresas con Flex tomado : {}".format(keyList))


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)


def get_processes_running():
    """ Takes tasklist output and parses the table into a dict"""
    tasks = subprocess.check_output(['tasklist']).decode('cp866', 'ignore').split("\r\n")
    p = []
    for task in tasks:
        m = re.match("(.+?) +(\\d+) (.+?) +(\\d+) +(\\d+.* K).*",task)
        if m is not None:
            p.append({"image":m.group(1),
                        "pid":m.group(2),
                        "session_name":m.group(3),
                        "session_num":m.group(4),
                        "mem_usage":m.group(5)
                        })
    return p


def buscarVentana(coords, cta_final, fecha_flex, fecha, nom_modulo):
    """Loop principal donde se descarga el Balance de cada empresa """

    while True:
        empresa2 = ag.locateCenterOnScreen(imPath(coords))
        if empresa2 is not None:
            break

    ag.click(empresa2, clicks=2)

    logging.info('Se hizo click en el Modulo: {}'.format(nom_modulo))

    while True:
        user_box = ag.locateCenterOnScreen(imPath('Aceptar_Cancelar.png'))
        if user_box is not None:
            print(user_box)
            break

    # Usuario
    ag.click(user_box[0] - 160, user_box[1] - 4)  
    caps1 = GetKeyState(VK_CAPITAL)
    if caps1 == 0:
        ag.write('CCORR')
    else:
        ag.write('ccorr')
    # Password
    ag.click(user_box[0] - 160, user_box[1] + 30)
    caps2 = GetKeyState(VK_CAPITAL)
    if caps2 == 0:
        ag.write('CCORR')
    else:
        ag.write('ccorr')
    ag.click(user_box[0] + 1, user_box[1] - 13)
    
    while True:
        flex_cuentas = ag.locateCenterOnScreen(imPath('Flex_Cuentas.png'))
        if flex_cuentas is not None:
            break

    ag.click(flex_cuentas)

    while True:
        flex_cuentas_ficha = ag.locateCenterOnScreen(imPath('Flex_Cuenta-ficha.png'))
        if flex_cuentas_ficha is not None:
            break

    ag.click(flex_cuentas_ficha)
    time.sleep(3)
    ag.click(206, 127)
    ag.write('1', interval=.25)
    ag.click(385, 128)
    ag.write(cta_final, interval=.25)
    ag.click(205, 265)
    ag.write(fecha_flex, interval=.25)
    ag.click(384, 261)
    ag.write(fecha_flex, interval=.25)
    ag.click(49, 182)
    ag.click(63, 149)

    ag.click(195, 392)

    logging.info('Buscando ventana Excel')


    time.sleep(90)
    b = ag.locateCenterOnScreen(imPath('Barra_Excel_Error.png'))

    if b is not None:
        logging.info("Se encontro Barra de Error: {}".format(b))
        ag.click(1081, 604)
        time.sleep(5)
        checkFlex = ag.locateOnScreen(imPath("Flex_ocupado.png"))
        if checkFlex == None:
            proc = conectarVentana()
            excel = Application().connect(process=int(proc))
            excel[u"excflx.txt - Excel"].maximize()
            sinDatosEnConsulta()
        else:
            ag.click(1137, 549)
            ag.click(1904, 13)
            ag.click(272, 401)
            ag.click(41, 38)
            ag.click(41, 125)
    else:
        time.sleep(15)
        logging.info("No se encontro Barra de Error: {}".format(b))
        checkFlex = ag.locateOnScreen(imPath("Flex_ocupado.png"))
        if checkFlex == None:
            proc = conectarVentana()
            excel = Application().connect(process=int(proc))
            excel[u"excflx.txt - Excel"].maximize()  
            ventanaEncontrada(nom_modulo, fecha)
        else:
            ag.click(1137, 549)
            ag.click(1904, 13)
            ag.click(272, 401)
            ag.click(41, 38)
            ag.click(41, 125)
    return checkFlex


def ventanaEncontrada(nom_modulo, fecha):
    # logging.info('Ventana encontrada, coordenadas: {}'.format(coords))
    time.sleep(3)
    ag.click(717, 19)
    ag.press("alt")
    ag.press("a")
    ag.press("v")
    ag.press("1")
    ag.click(763, 42)
    ag.write(r"W:\Cuentas-ficha\Cheques")
    ag.press("enter")
    ag.click(770, 904)
    caps3 = GetKeyState(VK_CAPITAL)
    if caps3 == 0:
        ag.write('Cheques_' + nom_modulo + '_' + fecha + '.xlsx')
    else:
        ag.write('cHEQUES_' + nom_modulo.lower() + '_' + fecha + '.XLSX')
    ag.click(771, 925)
    ag.write("ll")
    ag.press("enter")
    ag.click(1748, 1008)
    ag.click(1904, 14)
    ag.click(274, 399)
    ag.click(40, 39)
    ag.click(43, 125)
    logging.info("Archivo descargado, empresa: {}".format(nom_modulo))


def sinDatosEnConsulta():
    ag.moveTo(580, 1062, duration=1)
    ag.moveTo(580, 989, duration=1)
    ag.click(580, 989)
    ag.hotkey("winleft", "up")
    ag.hotkey("alt", "f4")
    ag.click(270, 395)
    ag.click(37, 32)
    ag.click(36, 119)
    logging.info("No hay datos en la consulta seleccionada")


def conectarVentana():
    df = get_processes_running()
    df = pd.DataFrame(df)
    proc = df.loc[df["image"] == "EXCEL.EXE"]["pid"].values[0]
    return proc



if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logging.info("Programa Finalizado por KeyboardInterrupt")
    except FailSafeException as f:
        logging.info("Programa Finalizado por FailSafe,", f)
        pass
