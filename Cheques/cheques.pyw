#!"C:\Python34\python.exe"

import pyautogui as ag
import time, os, logging
from datetime import date, timedelta
from pyautogui import FailSafeException

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
    # ('ACHEGEO', '1102002', (20, 130)),
    # ('ANDES', '1102002', (20, 147)),
    # ('APOCE', '1102002', (20, 164)),
    # ('AUTO PARQUIME', '1102000', (20, 181)),
    # ('AUTOORDEN', '1102008', (20, 198)),
    # ('CAFETAL', '1102001', (20, 215)),
    # ('CAJA EXPRESS', '1102005', (20, 232)),
    # ('CLEVER PARK SPA', '1102002', (20, 249)),
    # ('CONCESIONES PROVIDENCIA', '5111042', (20, 266)),
    # ('CONCESIONES PUNTA ARENAS', '1102004', (20, 283)),
    ('CONCESIONES SANTIAGO', '1102003', (20, 300)),
    # ('CONCESIONES RECOLETA', '1102002', (20, 317)),
    # ('CONSORCIO VALPARAISO', '1102004', (20, 334)),
    # ('COQUIMBO', '1102001', (20, 351)),
    # ('CP LATINA CHILE', '1102000', (20, 368)),
    # ('DENSITOSEA', '1102004', (20, 385)),
    # ('DIAGNOSIS', '1102000', (20, 402)),
    # ('ECM GEOTERMIA', '1102000', (20, 419)),
    # ('ESTACIONAR', '1102003', (20, 436)),
    # ('IMAGEN', '1102000', (20, 453)),
    # ('INFINERGEO', '1110132', (20, 470)),
    # ('INGENIEROS', '1102003', (20, 487)),
    # ('INVER EST', '1102002', (20, 504)),
    # ('INVERREST', '1102008', (20, 521)),
    # ('INVERTERRA', '1102004', (20, 538)),
    # ('IQUIQUE', '1102004', (20, 555)),
    # ('MEDCONSUL', '1102003', (20, 572)),
    # ('PIEDMONT', '1102005', (20, 589)),
    # ('POLLO ABRAZADO', '1102001', (20, 606)),
    # ('RENAL', '1102003', (20, 623)),
    # ('SEC', '1102004', (20, 640)),
    # ('SERVILAND', '1102008', (20, 657)),
    # ('TALCA', '1102002', (20, 674)),
    # ('TOMO IMAGEN', '1102005', (20, 691)),
    # ('VALDIVIA', '1102001', (20, 708)),
    # ('VALSEGUR', '1102005', (20, 725)),
    # ('VALSEGUR-ETV', '1102005', (20, 742))
]


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended.
    """
    return os.path.join('Images', filename)

def preparacion():
    logging.info('Comienza Preparación')
    ag.click(786, 94, clicks=2)
    ag.click(20, 130, clicks=2)
    while True:
        user_box = ag.locateCenterOnScreen(imPath('Aceptar_Cancelar.png'))
        if user_box is not None:
            print(user_box)
            break
    ag.click(user_box[0], user_box[1] + 15)
    # ag.click(454, 231)
    ag.click(1160, 96, clicks=2)
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

    ag.click(coords[0], coords[1], clicks=2)
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
    
    flex_cuentas = ag.locateCenterOnScreen(imPath('Flex_Cuentas.png'))
    ag.click(flex_cuentas)
    flex_cuentas_ficha = ag.locateCenterOnScreen(imPath('Flex_Cuenta-ficha.png'))
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
    
 	b = ag.screenshot('Barra_Excel_Cuenta.png', region=(2, 998, 914, 24))

    # error_consultas = ag.locateCenterOnScreen(imPath('Error_consultas.png'))
    # if error_consultas is not None:
    #         sinDatosEnConsulta()
    # else:
    #     ventanaEncontrada()    944, 999


    while True:
        error_consultas = ag.locateCenterOnScreen(imPath('Barra_Excel_Error.png'))
        excel1 = ag.locateCenterOnScreen(imPath('Barra_Excel_Cuenta.png'))
        if error_consultas is not None:
        	sinDatosEnConsulta()
        	break
        if excel1 is not None:
        	ag.click(700, 1008, button='right', duration=.25)
        	ag.click(748, 967)
        	ag.click(498, 644)
        	ventanaEncontrada(nom_modulo, fecha)
        	break


def ventanaEncontrada(nom_modulo, fecha):
    # logging.info('Ventana encontrada, coordenadas: {}'.format(coords))

    ag.click(23, 23)

    # time.sleep(3)

    ag.click(81, 203)

    # time.sleep(3)

    ag.click(327, 195)
    ag.write('Cheques_' + nom_modulo + '_' + fecha + '.xlsx')

    # time.sleep(2)

    ag.click(410, 513)
    ag.press('enter')

    # time.sleep(2)

    ag.click(956, 689)
    ag.click(956, 716)

    ag.click(632, 704)
    ag.press('enter')

    # time.sleep(3)

    ag.click(1270, 9)

    time.sleep(2)
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

    logging.info("Archivo descargado, empresa: {}".format(nom_modulo))


def sinDatosEnConsulta():
    error = ag.locateCenterOnScreen(imPath('Error_consultas.png'))
    ag.click(error, duration=.25)
    ag.press('enter')
    # while True:
    #     excel1 = ag.locateCenterOnScreen(imPath('Barra_Excel_Cuenta.png'))
    #     if excel1 is not None:
    #         ag.click(700, 1008, button='right', duration=.25)
    #         ag.click(748, 967)
    #         ag.click(498, 644)
    #         break
    time.sleep(3)
    ag.click(700, 1008, button='right', duration=.25)
    ag.click(748, 967)
    ag.click(498, 644)
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

867, 1007

    logging.info("No hay datos en la consulta seleccionada")


def main():
    logging.info('Programa iniciado. Presione Ctrl-C para abortar.')
    # preparacion()
    for i in empresas:
        buscarVentana(i[2], i[1], fechaFlex, fecha, i[0])


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logging.info("Programa Finalizado por KeyboardInterrupt")
    except FailSafeException as f:
    	logging.info("Programa Finalizado por FailSafe,", f)
    	pass
