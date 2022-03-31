from cheques import buscarVentana
from tkinter import Tk, OptionMenu, Button, StringVar, mainloop
from datetime import date, timedelta
import logging

OPTIONS = [
"Jan",
"Feb",
"Mar"
] #etc

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



master = Tk()

variable = StringVar(master)
variable.set("Seleccione empresas a descargar:") # default value

w = OptionMenu(master, variable, *empresas)
w.pack()

DICC_OPT = []


def ok():
    print ("value is:" + variable.get())
    a = variable.get()
    b = []
    b.append(a)
    return b

def loop(x):
	for i in x:
		buscarVentana(i[2], i[1], fechaFlex, fecha, i[0])

button = Button(master, text="OK", command=ok)
button.pack()


def main():
	mainloop()
	empresas = ok()
	print(empresas)
	# loop(x=empresas)


if __name__ == '__main__':
	main()