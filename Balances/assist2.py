import pyautogui as ag
import os, time, logging
from datetime import date

nombres = [
    ('ACHEGEO', (20, 130)),
    ('ANDES', (20, 147)),
    ('APOCE', (20, 164)),
    ('AUTO PARQUIME', (20, 181)),
    ('AUTOORDEN', (20, 198)),
    ('CAFETAL', (20, 215)),
    ('CAJA EXPRESS', (20, 232)),
    ('CLEVER PARK SPA', (20, 249)),
    ('CONCESIONES PROVIDENCIA', (20, 266)),
    ('CONCESIONES PUNTA ARENAS', (20, 283)),
    ('CONCESIONES SANTIAGO', (20, 300)),
    ('CONCESIONES RECOLETA', (20, 317)),
    ('CONSORCIO VALPARAISO', (20, 334)),
    ('COQUIMBO', (20, 351)),
    ('CP LATINA CHILE', (20, 368)),
    ('DENSITOSEA', (20, 385)),
    ('DIAGNOSIS', (20, 402)),
    ('ECM GEOTERMIA', (20, 419)),
    ('ESTACIONAR', (20, 436)),
    ('IMAGEN', (20, 453)),
    ('INFINERGEO', (20, 470)),
    ('INGENIEROS', (20, 487)),
    ('INVER EST', (20, 504)),
    ('INVERREST', (20, 521)),
    ('INVERTERRA', (20, 538)),
    ('IQUIQUE', (20, 555)),
    ('MEDCONSUL', (20, 572)),
    ('PIEDMONT', (20, 589)),
    ('POLLO ABRAZADO', (20, 606)),
    ('RENAL', (20, 623)),
    ('SEC', (20, 640)),
    ('SERVILAND', (20, 657)),
    ('TALCA', (20, 674)),
    ('TOMO IMAGEN', (20, 691)),
    ('VALDIVIA', (20, 708)),
    ('VALSEGUR', (20, 725)),
    ('VALSEGUR-ETV', (20, 742))
]

# nombres2 = ('ESTACIONAR.png', 'ESTACIONAR')
ag.PAUSE = 1

def imPath(filename):
    """A shortcut for joining the 'images/'' file path, since it is used so often. Returns the filename with 'images/' prepended."""
    return os.path.join('Images', filename)


def consolidar():
    # ag.click(1161, 98, clicks=2)
    ag.hotkey("winleft", "r")
    ag.write(r'W:\Balances\CONSOLIDADO')
    ag.press('enter')

    while True:
        consol = ag.locateCenterOnScreen(imPath('Consolidado_2021.png'))
        if consol is not None:
            break
    ag.click(consol, clicks=2)

    ag.click(856, 70)
    ag.click(474, 155)

    while True:
        botonSi = ag.locateCenterOnScreen(imPath('Boton_si.png'))
        if botonSi is not None:
            break
    ag.click(botonSi)

if __name__ == '__main__':
    consolidar()