import pyautogui as ag
import time, os
from datetime import date
from datetime import datetime

ag.PAUSE = 1

x = [
    ('ACHEGEO.png', 'ACHEGEO'), ('ANDES.png', 'ANDES'),
    ('APOCE.png', 'APOCE'),
    ('AUTO PARQUIME.png', 'AUTO PARQUIME'),
    ('AUTOORDEN.png', 'AUTOORDEN'),
    ('CAFETAL.png', 'CAFETAL'), ('CAJA EXPRESS.png', 'CAJA EXPRESS'),
    ('CLEVER PARK SPA.png', 'CLEVER PARK SPA'),
    ('CONCESIONES PROVIDENCIA.png', 'CONCESIONES PROVIDENCIA'),
    ('CONCESIONES PUNTA ARENAS.png', 'CONCESIONES PUNTA ARENAS'),
    ('CONCESIONES SANTIAGO.png', 'CONCESIONES SANTIAGO'),
    ('CONCESIONES RECOLETA.png', 'CONSECIONES RECOLETA'),
    ('CONSORCIO VALPARAISO.png', 'CONSORCIO VALPARAISO'),
    ('COQUIMBO.png', 'COQUIMBO'),
    ('CP LATINA CHILE.png', 'CP LATINA CHILE'),
    ('DENSITOSEA.png', 'DENSITOSEA'), ('DIAGNOSIS.png', 'DIAGNOSIS'),
    ('ECM GEOTERMIA.png', 'ECM GEOTERMIA'), ('ESTACIONAR.png', 'ESTACIONAR'),
    ('IMAGEN.png', 'IMAGEN'), ('INFINERGEO.png', 'INFINERGEO'),
    ('INGENIEROS.png', 'INGENIEROS'), ('INVER EST.png', 'INVER EST'),
    ('INVERREST.png', 'INVERREST'), ('INVERTERRA.png', 'INVERTERRA'),
    ('IQUIQUE.png', 'IQUIQUE'), ('MEDCONSUL.png', 'MEDCONSUL'),
    ('PIEDMONT.png', 'PIEDMONT'),
    ('POLLO ABRAZADO.png', 'POLLO ABRAZADO'), ('RENAL.png', 'RENAL'),
    ('SEC.png', 'SEC'), ('SERVILAND.png', 'SERVILAND'), ('TALCA.png', 'TALCA'),
    ('TOMO IMAGEN.png', 'TOMO IMAGEN'), ('VALDIVIA.png', 'VALDIVIA'),
    ('VALSEGUR.png', 'VALSEGUR'), ('VALSEGUR-ETV.png', 'VALSEGUR-ETV')
]

meses = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}


def main():
    for i in x:
        loop("1", "9999999", i[0], i[1])
    consolidar()


def imPath(filename):
    """A shortcut for joining the 'images/'' file path,
    since it is used so often.
    Returns the filename with 'images/' prepended."""
    return os.path.join('Images', filename)


def loop(cta_inicial, cta_final, nom_imagen, nom_modulo):

    fecha = date.today().strftime("%d-%m-%Y")
    mesActual = meses[datetime.now().month]

    #time.sleep(15)

    #ag.click(230, 906)

    modulo = ag.locateCenterOnScreen(imPath(nom_imagen))
    if modulo is None:
        raise Exception('Modulo no encontrado en pantalla')
    ag.click(modulo, duration=.25, clicks=2)

    time.sleep(3)

    ag.click(292, 210)
    ag.write("CCORR", interval=.25)

    ag.click(292, 246)
    ag.write("CCORR", interval=.25)

    ag.click(447, 195)

    ag.click(242, 30)

    ag.click(273, 87)

    time.sleep(4)

    ag.click(110, 223)
    ag.write(cta_inicial, interval=.25)

    ag.click(288, 222)
    ag.write(cta_final, interval=.25)

    ag.click(365, 302)
    ag.write(mesActual, interval=.25)

    ag.click(305, 138)
    ag.click(50, 162)

    ag.click(62, 131)
    ag.click(452, 106)

    ag.click(600, 551)

    time.sleep(15)

    ag.click(603, 529)

    ag.moveTo(65, 170)

    ag.dragTo(144, 172, 1, button='left')

    ag.click(button='right')

    time.sleep(2)

    ag.click(200, 279)

    ag.click(25, 21)

    time.sleep(2)

    ag.click(68, 198)

    ag.write('Balance_' + nom_modulo + '_' + str(fecha) + '.xlsx')

    ag.click(318, 190)

    time.sleep(2)

    #ag.click(411, 615)
    #ag.press('enter')

    time.sleep(2)

    ag.click(409, 495)
    ag.press('enter')

    #ag.click(994, 734)

    time.sleep(2)

    ## ag.click(556, 692)

    ag.click(953, 685)

    ##time.sleep(6)

    ag.click(953, 715)
    ag.click(565, 700)

    ag.click(927, 738)

    ag.click(1270, 9)

    time.sleep(2)

    ag.click(49, 59)

    ag.click(49, 132)

    ag.click(35, 32)

    ag.click(42, 117)


def consolidar():
    ag.click(24, 1012)

    ag.click(249, 650)

    time.sleep(2)

    ag.click(277, 94)
    ag.write(r'V:\Balances\CONSOLIDADO')
    ag.press('enter')

    while True:
        consol = ag.locateCenterOnScreen(imPath('Consolidado_2022.png'))
        if consol is not None:
            break
    ag.click(consol, clicks=2)

    #ag.click(39, 132)
    #ag.press('enter')

    #ag.click(37, 165)
    #ag.press('enter')

    time.sleep(30)

#def main3():
    ag.click(351, 911)
    #ag.click('enter')

    ag.click(1270, 12)

    ag.moveTo(41, 204)
    ag.dragTo(41, 136, 1, button='left')

    ag.click(682, 538, clicks=2)
    time.sleep(3)
    ag.press('enter')

    ag.click(165, 69)

    time.sleep(3)

    ag.moveTo(46, 790)

    ag.dragTo(154, 145, 2, button='left')

    time.sleep(3)

    ag.click(154, 145, button='right')

    time.sleep(3)

    ag.click(187, 408)

    ag.click(41, 132)
    ag.press('enter')

    time.sleep(3)
    ag.click(39, 151)
    ag.press('enter')

    time.sleep(3)

    ag.click(364, 962, button='right')

    time.sleep(3)
    ag.click(398, 861)

    time.sleep(3)

    ag.click(664, 12)
    time.sleep(3)
    ag.click(804, 12)


if __name__ == '__main__':
    main()
