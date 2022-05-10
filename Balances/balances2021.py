#!"C:\Cmder\SII\Scripts\python.exe"

"""

Programa de descarga de Balances ECM


"""

import time, os, logging, glob
import shutil
import ctypes
import subprocess
import re    
import pyautogui as ag
import pandas as pd
from datetime import date
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
from win32api import GetKeyState 
from win32con import VK_CAPITAL 


# VARIABLES GLOBALES

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
ag.PAUSE = 2

meses = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre"
}

fecha = date.today().strftime("%d-%m-%Y")

logging.basicConfig(filename=r"W:\Logs\Balances\Log_Balances_" + str(fecha) + '.txt',
                    level=logging.INFO,
                    format='%(asctime)s.%(msecs)03d: %(message)s',
                    datefmt='%H:%M:%S')

x = [
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


def timeis(func):
    '''Decorator that reports the execution time.'''

    def wrap(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        print(func.__name__, end - start)
        return result

    return wrap


@timeis
def main():
    # Función Principal

    logging.info('Programa iniciado. Presione Ctrl-C para abortar.')
    a = dict()

    for i in x:
        logging.info('Comienza Loop Empresa: %s' % i[0])
        a[i[0]] = mainLoop(i[1], "1", "9999999", i[0])

    logging.info('Diccionario final : {}'.format(a))

    # Flag to check if all elements are same 
    res = True 
  
    # extracting value to compare
    test_val = list(a.values())[0]
  
    for ele in a:
        if a[ele] != test_val:
            res = False 
            break

    if res == True:
        logging.info("Se descargaron todos los balances correctamente, se procede a consolidar.")
        consolidar()
    else:
        logging.info("No se descargaron todos los balances. No se consolida")



def preparacion():
    logging.info('Comienza Preparación')
    ag.hotkey("winleft", "r")
    ag.write(r"C:\Documents and Settings\dental02\Escritorio\Contabilidad")
    ag.press('enter')
    ag.click(20, 130, clicks=2)
    while True:
        user_box = ag.locateCenterOnScreen(imPath('Aceptar_Cancelar.png'))
        if user_box is not None:
            print(user_box)
            break
    ag.click(user_box[0], user_box[1] + 15)
    ag.hotkey("winleft", "r")
    ag.write(r"V:\\")
    ag.press('enter')
    ag.write("Asecm.Finan6", interval=.25)
    time.sleep(3)
    ag.click(671, 609)
    while True:
        EQUIS = ag.locateOnScreen(imPath('Boton_Equis.png'))
        if EQUIS is not None:
            break

    ag.click(EQUIS[0] + 45, EQUIS[1] + 9)


def mainLoop(coords, cta_inicial, cta_final, nom_modulo):
    """Loop principal donde se descarga el Balance de cada empresa """
    time.sleep(3)
    while True:
        empresas1 = ag.locateCenterOnScreen(imPath(coords))
        if empresas1 is not None:
            break
    ag.click(empresas1, clicks=2)
    logging.info('Se hizo click en el Modulo: {}'.format(nom_modulo))
    time.sleep(3)
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
    ag.click(253, 41)
    ag.click(285, 101)
    time.sleep(4)
    ag.click(110, 223)
    ag.write(cta_inicial, interval=.25)
    ag.click(288, 222)
    ag.write(cta_final, interval=.25)
    ag.click(365, 302)
    ag.write("d", interval=.25)
    ag.click(305, 138)
    ag.click(50, 162)
    ag.click(62, 131)
    ag.click(452, 106)
    ag.click(1081, 593)
    time.sleep(10)
    logging.info('Buscando ventana Excel')
    checkFlex = ag.locateOnScreen(imPath("Flex_ocupado.png"))
    if checkFlex == None:
        flexNoTomado(nom_modulo)
    else:
        flexTomado()
    return checkFlex




def flexNoTomado(nom):
    while True:
        a = conectarVentana()
        if a is not None:
            break
        
    excel = Application().connect(process=int(a))
    excel[u"excflx.txt - Excel"].maximize()        
    ag.click(498, 644)
    ag.click(603, 529)
    time.sleep(3)
    ag.moveTo(65, 200)
    ag.dragTo(144, 200, 1, button='left')
    ag.click(button='right')
    time.sleep(2)
    ag.click(200, 337)
    ag.press("alt")
    ag.press("a")
    ag.press("v")
    ag.press("1")
    ag.click(763, 42)
    ag.write(r"W:\Balances")
    ag.press("enter")
    ag.click(770, 904)
    caps3 = GetKeyState(VK_CAPITAL)
    if caps3 == 0:
        ag.write('Balance_2021_' + nom + '_' + str(fecha) + '.xlsx')
    else:
        ag.write('bALANCE_2021_' + nom.lower() + '_' + str(fecha) + '.xlsx') 
    ag.click(771, 925)
    ag.write("ll")
    ag.press("enter")
    ag.click(1748, 1008)
    ag.click(1904, 14)
    ag.click(458, 140)
    ag.click(40, 39)
    ag.click(43, 125)
    logging.info("Archivo descargado, empresa: {}".format(nom))


def flexTomado():
    ag.click(1137, 549)
    ag.click(1904, 13)
    ag.click(452, 139)
    ag.click(41, 38)
    ag.click(41, 125)


def conectarVentana():
    df = get_processes_running()
    df = pd.DataFrame(df)
    proc = df.loc[df["image"] == "EXCEL.EXE"]["pid"].values[0]
    return proc


def consolidar():
    ag.press("winleft")
    ag.write("ejecutar")
    ag.press("enter")
    ag.write(r"W:\Balances\CONSOLIDADO")
    ag.press('enter')
    ag.hotkey("winleft", "up")
    ag.click(742, 37)
    ag.write(r"W:\Balances\CONSOLIDADO\CONSOLIDADO_2021.xlsm")
    ag.press("enter")
    while True:
        botonSi = ag.locateCenterOnScreen(imPath('Boton_si.png'))
        if botonSi is not None:
            break
    ag.click(botonSi[0] - 123, botonSi[1] + 64)
    ag.click(1902, 12)
    time.sleep(3)
    ag.moveTo(431, 171)
    ag.dragTo(431, 127, 1, button='left')
    ag.click(989, 548)
    ag.click(1893, 10)
    source = 'W:\\Balances'
    dest = 'W:\\Balances\\CONSOLIDADO\\Balances_historicos' 
    os.chdir(source)
    for f in glob.glob("*.xlsx"):
        shutil.move(f, dest)



if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logging.info("Programa Finalizado por KeyboardInterrupt")
    except FailSafeException as f:
        logging.info("Programa Finalizado por FailSafe,", f)
        pass