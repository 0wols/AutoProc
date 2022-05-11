"""

Programa de descarga de Cuentas Ficha ECM


"""

import pyautogui as ag
import time, os, logging
import ctypes
import subprocess
import re    
import pandas as pd
from datetime import date
from pywinauto.application import Application
from win32api import GetKeyState 
from win32con import VK_CAPITAL
from pyscreeze import Box



# VARIABLES GLOBALES

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
ag.PAUSE = 2

fecha = date.today().strftime("%d-%m-%Y")

logging.basicConfig(
    filename=r"W:\Logs\Cuentas-ficha\Log_CuentaFicha_" + str(fecha) + '.txt',
    level=logging.INFO,
    format='%(asctime)s.%(msecs)03d: %(message)s',
    datefmt='%H:%M:%S')

nombres = [
    ('ACHEGEO', 'ACHEGEO.png'),
    ('ANDES', 'ANDES.png'),
    ('APOCE', 'APOCE.png'),
    ('AUTO PARQUIME', 'AUTO PARQUIME.png'),
    ('AUTOORDEN', 'AUTOORDEN.png'),
    ('CAFETAL', 'CAFETAL.png'),
    ('CAJA EXPRESS', 'CAJA EXPRESS.png'),
    ('CLEVER PARK SPA', 'CLEVER PARK SPA.png'),
    ('CONCESIONES PROVIDENCIA', 'CONCESIONES PROVIDENCIA.png'),
    ('CONCESIONES PUNTA ARENAS', 'CONCESIONES PUNTA ARENAS.png'),
    ('CONCESIONES SANTIAGO', 'CONCESIONES SANTIAGO.png'),
    ('CONCESIONES RECOLETA', 'CONCESIONES RECOLETA.png'),
    ('CONSORCIO VALPARAISO', 'CONSORCIO VALPARAISO.png'),
    ('COQUIMBO', 'COQUIMBO.png'),
    ('CP LATINA CHILE', 'CP LATINA CHILE.png'),
    ('DENSITOSEA', 'DENSITOSEA.png'),
    ('DIAGNOSIS', 'DIAGNOSIS.png'),
    ('ECM GEOTERMIA', 'ECM GEOTERMIA.png'),
    ('ESTACIONAR', 'ESTACIONAR.png'),
    ('IMAGEN', 'IMAGEN.png'),
    ('INFINERGEO', 'INFINERGEO.png'),
    ('INGENIEROS', 'INGENIEROS.png'),
    ('INVER EST', 'INVER EST.png'),
    ('INVERREST', 'INVERREST.png'),
    ('INVERTERRA', 'INVERTERRA.png'),
    ('IQUIQUE', 'IQUIQUE.png'),
    ('MEDCONSUL', 'MEDCONSUL.png'),
    ('PIEDMONT', 'PIEDMONT.png'),
    ('POLLO ABRAZADO', 'POLLO ABRAZADO.png'),
    ('RENAL', 'RENAL.png'),
    ('SEC', 'SEC.png'),
    ('SERVILAND', 'SERVILAND.png'),
    ('TALCA', 'TALCA.png'),
    ('TOMO IMAGEN', 'TOMO IMAGEN.png'),
    ('VALDIVIA', 'VALDIVIA.png'),
    ('VALSEGUR', 'VALSEGUR.png'),
    ('VALSEGUR-ETV', 'VALSEGUR-ETV.png')
]


def main():
    # Función Principal

    logging.info("Programa iniciado. Presione Ctrl-C para abortar.")


    miDicc = dict()

    for i in nombres:
        miDicc[i[0]] = funcion1(i[1], i[0], "1", "9999999", "1", "31122021")

    valorTest = Box(left=737, top=522, width=357, height=39)

    keyList = [ k for k in miDicc.keys() if miDicc[k] == valorTest ]

    logging.info("Programa Finalizado. Empresas con Flex tomado : {}".format(keyList))



def imPath(filename):
    """A shortcut for joining the 'images/'' file path, since it is used so often. Returns the filename with 'images/' prepended."""
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



def funcion1(coords, nom_modulo, cta_inicial, cta_final, fecha_inicial, fecha_final):
    fecha = date.today().strftime("%d-%m-%Y")

    while True:
        empresas1 = ag.locateCenterOnScreen(imPath(coords))
        if empresas1 is not None:
            break

    ag.click(empresas1, clicks=2)
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
    flex_cuentas = ag.locateCenterOnScreen(imPath('Flex_Cuentas.png'))
    ag.click(flex_cuentas)
    flex_cuentas_ficha = ag.locateCenterOnScreen(imPath('Flex_Cuenta-ficha.png'))
    ag.click(flex_cuentas_ficha)
    time.sleep(5)
    ag.click(206, 127)
    ag.write(cta_inicial, interval=.25)
    ag.click(385, 128)
    ag.write(cta_final, interval=.25)
    ag.click(205, 265)
    ag.write(fecha_inicial, interval=.25)
    ag.click(384, 261)
    ag.write(fecha_final, interval=.25)
    ag.click(49, 182)
    ag.click(63, 149)
    ag.click(195, 392)
    time.sleep(10)

    while True:
        a = ag.locateCenterOnScreen(imPath('Barra_Completado.png'))
        if a is not None:
            checkFlex = ag.locateOnScreen(imPath("Flex_ocupado.png"))
            if checkFlex == None:   
                proc = conectarVentana()
                excel = Application().connect(process=int(proc))
                excel[u"excflx.txt - Excel"].maximize()  
                ventanaEncontrada(nom_modulo, fecha)
                break
            else:
                ag.click(1137, 549)
                ag.click(1904, 13)
                ag.click(272, 401)
                ag.click(41, 38)
                ag.click(41, 125)
                break
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
    ag.write(r"W:\Cuentas-ficha\Cuentas-ficha")
    ag.press("enter")
    ag.click(770, 904)
    caps3 = GetKeyState(VK_CAPITAL)
    if caps3 == 0:
        ag.write('Cuenta-ficha_' + nom_modulo + '_' + str(fecha) + '.xlsx')
    else:
        ag.write('cUENTA-FICHA_' + nom_modulo.lower() + '_' + str(fecha) + '.XLSX') 
    ag.click(771, 925)
    ag.write("ll")
    ag.press("enter")
    ag.click(1748, 1008)
    time.sleep(10)
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