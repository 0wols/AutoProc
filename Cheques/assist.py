import pyautogui as ag
import time, os, logging
import ctypes
import subprocess
import re    
import pandas as pd
from datetime import date, timedelta
from pyautogui import FailSafeException
from pywinauto.application import Application


ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
ag.PAUSE = 2

hoy = date.today()
hoy_formato_dia = date.today().strftime("%a")

logging.basicConfig(
    filename='Log_Cheques_' + date.today().strftime("%d-%m-%Y") + '.txt',
    level=logging.INFO,
    format='%(asctime)s.%(msecs)03d: %(message)s',
    datefmt='%H:%M:%S')

if hoy_formato_dia == 'Mon':
    fechaFlex = hoy - timedelta(days=3)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")
else:
    fechaFlex = hoy - timedelta(days=1)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")

fecha = date.today().strftime("%d-%m-%Y")

print(fecha)

empresas = [
    ('ACHEGEO', '1102002', 'ACHEGEO.png'),
    ('ANDES', '1102002', 'ANDES.png'),
    ('APOCE', '1102002', 'APOCE.png'),
    ('AUTO PARQUIME', '1102000', 'AUTO PARQUIME.png'),
    ('AUTOORDEN', '1102008', 'AUTOORDEN.png'),
    ('CAFETAL', '1102001', 'CAFETAL.png'),
    ('CAJA EXPRESS', '1102005', 'CAJA EXPRESS.png'),
    ('CLEVER PARK SPA', '1102002', 'CLEVER PARK SPA.png'),
    ('CONCESIONES PROVIDENCIA', '5111042', 'CONCESIONES PROVIDENCIA.png'),
    ('CONCESIONES PUNTA ARENAS', '1102004', 'CONCESIONES PUNTA ARENAS.png'),
    ('CONCESIONES SANTIAGO', '1102003', 'CONCESIONES SANTIAGO.png'),
    ('CONCESIONES RECOLETA', '1102002', 'CONCESIONES RECOLETA.png'),
    ('CONSORCIO VALPARAISO', '1102004', 'CONSORCIO VALPARAISO.png'),
    ('COQUIMBO', '1102001', 'COQUIMBO.png'),
    ('CP LATINA CHILE', '1102000', 'CP LATINA CHILE.png'),
    ('DENSITOSEA', '1102004', 'DENSITOSEA.png'),
    ('DIAGNOSIS', '1102000', 'DIAGNOSIS.png'),
    ('ECM GEOTERM1IA', '1102000', 'ECM GEOTERMIA.png'),
    ('ESTACIONAR', '1102003', 'ESTACIONAR.png'),
    ('IMAGEN', '1102000', 'IMAGEN.png'),
    ('INFINERGEO', '1110132', 'INFINERGEO.png'),
    ('INGENIEROS', '1102003', 'INGENIEROS.png'),
    ('INVER EST', '1102002', 'INVER EST.png'),
    ('INVERREST', '1102008', 'INVERREST.png'),
    ('INVERTERRA', '1102004', 'INVERTERRA.png'),
    ('IQUIQUE', '1102004', 'IQUIQUE.png'),
    ('MEDCONSUL', '1102003', 'MEDCONSUL.png'),
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
    # preparacion()
    for i in empresas:
        buscarVentana(i[2], i[1], fechaFlex, fecha, i[0])


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


def preparacion():
    logging.info('Comienza Preparación')
    # ag.click(786, 94, clicks=2)
    ag.hotkey("winleft", "r")
    ag.write(r"C:\Documents and Settings\dental02\Escritorio\Contabilidad")
    ag.press('enter')
    # ag.click(20, 130, clicks=2)


    while True:
        empresa1 = ag.locateCenterOnScreen(imPath('Achegeo_Selec.png'))
        if empresa1 is not None:
            break
    ag.click(empresa1, clicks=2)


    while True:
        user_box = ag.locateCenterOnScreen(imPath('Aceptar_Cancelar.png'))
        if user_box is not None:
            print(user_box)
            break
    ag.click(user_box[0], user_box[1] + 15)
    # ag.click(454, 231)
    # ag.click(1160, 96, clicks=2)
    ag.hotkey("winleft", "r")
    ag.write(r"V:\\")
    ag.press('enter')
    ag.write("Asecm.Finan6", interval=.25)
    time.sleep(3)
    ag.click(671, 609)
    # ag.press('enter')
    while True:
        EQUIS = ag.locateOnScreen(imPath('Boton_Equis.png'))
        if EQUIS is not None:
            break

    ag.click(EQUIS[0] + 45, EQUIS[1] + 9)


# def mainLoop(cta_inicial, cta_final, nom_imagen, nom_modulo):
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
    ag.write('CCORR')
    # Password
    ag.click(user_box[0] - 160, user_box[1] + 30)
    ag.write('CCORR')
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


    time.sleep(15)
    # ag.screenshot('Barra_Excel_Error.png', region=(0, 1043, 753, 36))
    
    a = ag.locateCenterOnScreen(imPath('Barra_Excel_Error.png'))
    if a is not None:
        logging.info("Se encontro Barra de Error")
        ag.click(1081, 604)
        time.sleep(1)
        proc = conectarVentana()
        excel = Application().connect(process=int(proc))
        excel[u"excflx.txt - Excel"].maximize()
        sinDatosEnConsulta()
    else:
        time.sleep(15)
        logging.info("No se encontro Barra de Error")
        proc = conectarVentana()
        excel = Application().connect(process=int(proc))
        excel[u"excflx.txt - Excel"].maximize()  
        # ag.moveTo(512, 1062, duration=2)
        # ag.moveTo(525, 961, duration=2)
        # ag.click(525, 961)
        ventanaEncontrada(nom_modulo, fecha)


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
    # time.sleep(2)
    ag.click(770, 904)
    ag.write('Cheques_' + nom_modulo + '_' + fecha + '.xlsx')
    ag.click(771, 925)
    ag.write("ll")
    ag.press("enter")
    ag.click(1748, 1008)
    # time.sleep(2)
    # time.sleep(2)
    ag.click(1904, 14)
    # ag.press('enter')

    # time.sleep(2)

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

    # opciones = ag.locateCenterOnScreen(imPath('Opciones.png'))
    # ag.click(opciones, duration=.25)

    # salir1 = ag.locateCenterOnScreen(imPath('Salir1.png'))
    # ag.click(salir1, duration=.25)

    # archivo = ag.locateCenterOnScreen(imPath('Archivo.png'))
    # ag.click(archivo, duration=.25)

    # salir2 = ag.locateCenterOnScreen(imPath('Salir2.png'))
    # ag.click(salir2, duration=.25)
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
